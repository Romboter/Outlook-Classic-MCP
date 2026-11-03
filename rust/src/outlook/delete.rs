//! Outlook side of confirmed deletes (`crate::action`): read the facts the
//! person will see, refuse what must never be deleted, stamp the item, then
//! move it to Deleted Items once.
//!
//! `Delete()` on an item outside Deleted Items moves it there (verify live on
//! each store type). Permanent deletion is never offered, and nothing here
//! touches the item after `Delete()` starts.
use super::compose::{fnv1a, new_id};
use super::mailbox::{default_folder, optional_text};
use super::schedule::date_iso;
use super::win;
use crate::action::{ActionKind, ActionPreview, ActionTarget, Fact, Located};
use crate::ticket::{NOT_SENT, UNKNOWN_OUTCOME};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const OL_FOLDER_DELETED: i32 = 3;
const STAMP: &str = "OutlookMcpActionId";
const OL_TEXT: i32 = 1;
const SCAN_NEWEST: i32 = 200;

fn not_done(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

fn item_by_id(namespace: &IDispatch, target: &ActionTarget) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(target.entry_id.as_str())];
    if let Some(store) = &target.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Item no longer available; EntryID may be stale")
}

pub(super) fn stamp_of(item: &IDispatch) -> Option<String> {
    let props = win::object_property(item, "UserProperties").ok()?;
    let prop = win::method_object(&props, "Find", vec![VARIANT::from(STAMP)]).ok()?;
    win::string_property(&prop, "Value").ok()
}

/// True when `item` is in Deleted Items or any folder below it. Unknown
/// counts as "yes": when in doubt, do not delete.
pub(super) fn in_deleted_items(namespace: &IDispatch, item: &IDispatch) -> bool {
    let Ok(deleted) = default_folder(namespace, OL_FOLDER_DELETED) else {
        return true;
    };
    let deleted_id = optional_text(&deleted, "EntryID");
    if deleted_id.is_empty() {
        return true;
    }
    let Ok(mut current) = win::object_property(item, "Parent") else {
        return true;
    };
    for _ in 0..16 {
        if optional_text(&current, "EntryID") == deleted_id {
            return true;
        }
        match win::object_property(&current, "Parent") {
            Ok(parent) => current = parent,
            // Reached the store root without meeting Deleted Items.
            Err(_) => return false,
        }
    }
    true
}

fn check_kind(item: &IDispatch, kind: ActionKind) -> Result<()> {
    let class = win::int_property(item, "Class")?;
    match kind {
        ActionKind::DeleteMail if class == 43 || class == 53 => Ok(()),
        ActionKind::DeleteEvent if class == 26 => Ok(()),
        ActionKind::DeleteMail => bail!("that EntryID does not identify a mail item"),
        ActionKind::DeleteEvent => bail!("that EntryID does not identify a calendar event"),
        ActionKind::ToggleRule | ActionKind::RespondEvent => bail!("not an item to delete"),
    }
}

/// Pre-prompt refusals that never reach the person.
fn check_allowed(namespace: &IDispatch, item: &IDispatch, target: &ActionTarget) -> Result<()> {
    check_kind(item, target.kind)?;
    if in_deleted_items(namespace, item) {
        bail!(
            "that item is already in Deleted Items (or cannot be placed); permanent deletion is never offered"
        );
    }
    if target.kind == ActionKind::DeleteEvent
        && win::int_property(item, "MeetingStatus").unwrap_or(0) != 0
    {
        bail!(
            "that event is a meeting with attendees; deleting it alone would leave it on their calendars, and cancelling meetings is not available yet"
        );
    }
    Ok(())
}

fn facts_of(item: &IDispatch, kind: ActionKind) -> (String, Vec<Fact>, Vec<String>, String) {
    let fact = |label: &str, value: String| Fact {
        label: label.into(),
        value,
    };
    let folder = win::object_property(item, "Parent")
        .map(|p| optional_text(&p, "Name"))
        .unwrap_or_default();
    match kind {
        ActionKind::DeleteMail => {
            let sender = optional_text(item, "SenderName");
            let address = super::table::public_address(&optional_text(item, "SenderEmailAddress"));
            let subject = optional_text(item, "Subject");
            let received = date_iso(item, "ReceivedTime").unwrap_or_default();
            let identity = format!("mail|{subject}|{received}|{address}");
            (
                "Move this email to Deleted Items (it can be restored from there).".into(),
                vec![
                    fact("From", format!("{sender} <{address}>")),
                    fact("Subject", subject),
                    fact("Received", received),
                    fact("Folder", folder),
                ],
                vec![],
                identity,
            )
        }
        ActionKind::DeleteEvent => {
            let subject = optional_text(item, "Subject");
            let start = date_iso(item, "Start").unwrap_or_default();
            let end = date_iso(item, "End").unwrap_or_default();
            let state = win::int_property(item, "RecurrenceState").unwrap_or(0);
            let (scope, warnings) = match state {
                1 => (
                    "the whole recurring series",
                    vec!["Every occurrence of this series will be removed.".to_string()],
                ),
                2 | 3 => ("this occurrence only", vec![]),
                _ => ("this event", vec![]),
            };
            let identity = format!("event|{subject}|{start}|{state}");
            (
                "Move this calendar event to Deleted Items (it can be restored from there).".into(),
                vec![
                    fact("Event", subject),
                    fact("Starts", start),
                    fact("Ends", end),
                    fact("Location", optional_text(item, "Location")),
                    fact("Applies to", scope.into()),
                    fact("Folder", folder),
                ],
                warnings,
                identity,
            )
        }
        // Rule changes are handled in rule.rs and never reach this function.
        ActionKind::ToggleRule | ActionKind::RespondEvent => {
            (String::new(), vec![], vec![], String::new())
        }
    }
}

fn hash_of(raw_entry_id: &str, identity: &str) -> String {
    fnv1a(&format!("{raw_entry_id}|{identity}"))
}

pub(super) fn prepare(namespace: &IDispatch, target: &ActionTarget) -> Result<Value> {
    let item = item_by_id(namespace, target)?;
    check_allowed(namespace, &item, target)?;
    let (summary, facts, warnings, identity) = facts_of(&item, target.kind);
    let action_id = new_id();
    // The stamp lets a later check find the item in Deleted Items even though
    // its EntryID changes when it moves.
    let props = win::object_property(&item, "UserProperties")?;
    let stamp = win::method_object(
        &props,
        "Add",
        vec![
            VARIANT::from(STAMP),
            VARIANT::from(OL_TEXT),
            VARIANT::from(false),
        ],
    )?;
    win::put_property(&stamp, "Value", VARIANT::from(action_id.as_str()))?;
    win::invoke(&item, "Save", DISPATCH_METHOD, vec![])?;
    Ok(json!({"preview": ActionPreview {
        action_id,
        summary,
        facts,
        warnings,
        hash: hash_of(&target.entry_id, &identity),
    }}))
}

pub(super) fn execute(
    namespace: &IDispatch,
    target: &ActionTarget,
    preview: &ActionPreview,
) -> Result<Value> {
    // Everything before Delete() can fail safely: the item stays where it is.
    let item = item_by_id(namespace, target).map_err(|e| not_done(&format!("{e:#}")))?;
    check_allowed(namespace, &item, target).map_err(|e| not_done(&format!("{e:#}")))?;
    if stamp_of(&item).as_deref() != Some(preview.action_id.as_str()) {
        return Err(not_done("the item does not carry this action id"));
    }
    let (_, _, _, identity) = facts_of(&item, target.kind);
    if hash_of(&target.entry_id, &identity) != preview.hash {
        return Err(not_done("the item changed after it was shown to you"));
    }
    // From here the outcome is unknown unless Delete() returns cleanly, and
    // the item must not be touched again.
    match win::invoke(&item, "Delete", DISPATCH_METHOD, vec![]) {
        Ok(_) => Ok(json!({"done": true})),
        Err(error) => Err(anyhow::anyhow!(
            "{UNKNOWN_OUTCOME}Delete() reported an error ({error:#}); the item may still have been deleted"
        )),
    }
}

pub(super) fn locate(
    namespace: &IDispatch,
    action_id: &str,
    target: Option<&ActionTarget>,
) -> Result<Value> {
    if action_id.is_empty()
        || action_id.len() > 80
        || !action_id
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-')
    {
        bail!("invalid action_id");
    }
    let result = |located: Located| {
        json!({"located": match located {
            Located::Done => "done",
            Located::NotDone => "not_done",
            Located::Unknown => "unknown",
        }})
    };
    // The original id still resolves: either untouched, or it moved within the store.
    if let Some(target) = target
        && let Ok(item) = item_by_id(namespace, target)
        && stamp_of(&item).as_deref() == Some(action_id)
    {
        return Ok(result(if in_deleted_items(namespace, &item) {
            Located::Done
        } else {
            Located::NotDone
        }));
    }
    // Otherwise look for the stamp among the newest Deleted Items.
    if let Ok(items) =
        default_folder(namespace, OL_FOLDER_DELETED).and_then(|f| win::object_property(&f, "Items"))
    {
        let _ = win::invoke(
            &items,
            "Sort",
            DISPATCH_METHOD,
            vec![VARIANT::from("[LastModificationTime]"), VARIANT::from(true)],
        );
        let count = win::int_property(&items, "Count")
            .unwrap_or(0)
            .clamp(0, SCAN_NEWEST);
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(20);
        for i in 1..=count {
            if std::time::Instant::now() > deadline {
                break;
            }
            if let Ok(item) = win::method_object(&items, "Item", vec![VARIANT::from(i)])
                && stamp_of(&item).as_deref() == Some(action_id)
            {
                return Ok(result(Located::Done));
            }
        }
    }
    Ok(result(Located::Unknown))
}
