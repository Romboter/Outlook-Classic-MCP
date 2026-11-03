//! Outlook side of confirmed rule changes (`crate::action`).
//!
//! Rules save immediately and act on mail automatically, so the prompt says
//! what the rule does. The state is read back after saving; any doubt after
//! the change has begun is an unknown outcome.
use super::compose::{fnv1a, new_id};
use super::mailbox::{bool_property, optional_text};
use super::win;
use crate::action::{ActionPreview, ActionTarget, Fact, Located};
use crate::ticket::{NOT_SENT, UNKNOWN_OUTCOME};
use anyhow::{Result, bail};
use serde_json::{Value, json};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const MAX_RULES: i32 = 500;

/// `(rule action property, plain wording, acts on other people or deletes)`.
const EFFECTS: [(&str, &str, bool); 11] = [
    ("Delete", "deletes the message", true),
    ("DeletePermanently", "PERMANENTLY deletes the message", true),
    ("Forward", "forwards the message to other people", true),
    (
        "ForwardAsAttachment",
        "forwards the message as an attachment",
        true,
    ),
    ("Redirect", "redirects the message to other people", true),
    ("CC", "copies the message to other people", true),
    ("MoveToFolder", "moves the message to a folder", false),
    ("CopyToFolder", "copies the message to a folder", false),
    ("MarkAsTask", "flags the message as a task", false),
    ("AssignToCategory", "assigns a category", false),
    ("Stop", "stops processing later rules", false),
];

fn not_done(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

fn rule_name(target: &ActionTarget) -> &str {
    target.entry_id.trim()
}

/// The collection and the one rule with this exact name.
fn find_rule(namespace: &IDispatch, name: &str) -> Result<(IDispatch, IDispatch)> {
    let store = win::object_property(namespace, "DefaultStore")?;
    let rules = win::method_object(&store, "GetRules", vec![])?;
    let count = win::int_property(&rules, "Count")?.clamp(0, MAX_RULES);
    let mut found = Vec::new();
    for i in 1..=count {
        let rule = win::method_object(&rules, "Item", vec![VARIANT::from(i)])?;
        if optional_text(&rule, "Name") == name {
            found.push(rule);
        }
    }
    match found.len() {
        0 => bail!("no rule is named that; use outlook_list_rules to see the names"),
        1 => Ok((rules, found.remove(0))),
        _ => bail!("more than one rule has that name, so it cannot be identified safely"),
    }
}

/// What the rule does, read from Outlook, and whether any of it acts on
/// other people or deletes.
fn effects_of(rule: &IDispatch) -> (Vec<String>, bool) {
    let mut described = Vec::new();
    let mut risky = false;
    if let Ok(actions) = win::object_property(rule, "Actions") {
        for (property, wording, danger) in EFFECTS {
            if let Ok(action) = win::object_property(&actions, property)
                && bool_property(&action, "Enabled")
            {
                described.push(wording.to_string());
                risky |= danger;
            }
        }
    }
    (described, risky)
}

fn state_word(on: bool) -> &'static str {
    if on { "ON" } else { "OFF" }
}

fn hash_of(name: &str, current: bool, effects: &[String]) -> String {
    fnv1a(&format!("{name}|{current}|{}", effects.join(";")))
}

pub(super) fn prepare(namespace: &IDispatch, target: &ActionTarget) -> Result<Value> {
    let name = rule_name(target);
    let want = target.enabled.unwrap_or(true);
    let (_, rule) = find_rule(namespace, name)?;
    let current = bool_property(&rule, "Enabled");
    if current == want {
        bail!(
            "that rule is already {}; nothing to change",
            state_word(current)
        );
    }
    let (effects, risky) = effects_of(&rule);
    let kind = match win::int_property(&rule, "RuleType").unwrap_or(0) {
        1 => "outgoing mail",
        _ => "incoming mail",
    };
    let mut warnings = Vec::new();
    if want && risky {
        warnings.push(format!(
            "Once on, this rule acts on mail automatically and {}.",
            effects.join(" and ")
        ));
    }
    let what = if effects.is_empty() {
        "no actions that could be read".to_string()
    } else {
        effects.join("; ")
    };
    Ok(json!({"preview": ActionPreview {
        action_id: new_id(),
        summary: format!(
            "Turn the mail rule \"{name}\" {}. It takes effect immediately.",
            state_word(want)
        ),
        facts: vec![
            Fact { label: "Rule".into(), value: name.to_string() },
            Fact { label: "Applies to".into(), value: kind.into() },
            Fact { label: "Now".into(), value: state_word(current).into() },
            Fact { label: "Will be".into(), value: state_word(want).into() },
            Fact { label: "What it does".into(), value: what },
        ],
        warnings,
        hash: hash_of(name, current, &effects),
    }}))
}

pub(super) fn execute(
    namespace: &IDispatch,
    target: &ActionTarget,
    preview: &ActionPreview,
) -> Result<Value> {
    let name = rule_name(target);
    let want = target.enabled.unwrap_or(true);
    // Everything before the change can fail safely: the rule is untouched.
    let (rules, rule) = find_rule(namespace, name).map_err(|e| not_done(&format!("{e:#}")))?;
    let current = bool_property(&rule, "Enabled");
    if current == want {
        return Err(not_done("the rule is already in that state"));
    }
    let (effects, _) = effects_of(&rule);
    if hash_of(name, current, &effects) != preview.hash {
        return Err(not_done("the rule changed after it was shown to you"));
    }
    // From here the outcome is unknown unless the change saves and reads back.
    let unknown = |why: String| anyhow::anyhow!("{UNKNOWN_OUTCOME}{why}");
    win::put_property(&rule, "Enabled", VARIANT::from(want))
        .map_err(|e| unknown(format!("setting the rule failed ({e:#})")))?;
    // No progress dialog: it would block the Outlook worker.
    win::invoke(&rules, "Save", DISPATCH_METHOD, vec![VARIANT::from(false)]).map_err(|e| {
        unknown(format!(
            "saving the rules failed ({e:#}); the change may have applied"
        ))
    })?;
    if bool_property(&rule, "Enabled") != want {
        return Err(unknown(
            "Outlook reports a different state after saving".into(),
        ));
    }
    Ok(json!({"done": true}))
}

pub(super) fn locate(namespace: &IDispatch, target: Option<&ActionTarget>) -> Result<Value> {
    let located = match target {
        Some(target) => match find_rule(namespace, rule_name(target)) {
            Ok((_, rule)) => {
                if bool_property(&rule, "Enabled") == target.enabled.unwrap_or(true) {
                    Located::Done
                } else {
                    Located::NotDone
                }
            }
            Err(_) => Located::Unknown,
        },
        // Without the original request there is no way to know which rule was meant.
        None => Located::Unknown,
    };
    Ok(json!({"located": match located {
        Located::Done => "done",
        Located::NotDone => "not_done",
        Located::Unknown => "unknown",
    }}))
}
