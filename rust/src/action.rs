//! Confirmed actions that are not sends: deleting mail and events, and
//! switching mail rules. See `docs/CONFIRMED-ACTIONS-DESIGN.md`.
//!
//! No mode or setting skips the person's approval: the prompt is built from
//! facts read from Outlook, a client that cannot show it means nothing
//! happens, and an action that may have started is reported as an unknown
//! outcome and never replayed. Outlook and the person are reached through
//! `ActionPort` and `Confirmer`, so every path is tested with fakes.
#![allow(async_fn_in_trait)]

use crate::policy::Capability;
use crate::send::{Confirmation, Confirmer, Outcome, PortError};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::{
    collections::HashMap,
    sync::{Arc, Mutex},
    time::{Duration, Instant},
};

const HOUR: Duration = Duration::from_secs(3600);

/// Thresholds are speed bumps (they add a section to the prompt); the hard
/// ceilings only stop runaways. Strict turns a threshold into a refusal.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActionSettings {
    pub delete_per_hour: usize,
    pub delete_hard_max_per_hour: usize,
    pub rule_per_hour: usize,
    pub rule_hard_max_per_hour: usize,
    pub response_per_hour: usize,
    pub response_hard_max_per_hour: usize,
    pub strict: bool,
}

impl Default for ActionSettings {
    fn default() -> Self {
        Self {
            delete_per_hour: 10,
            delete_hard_max_per_hour: 50,
            rule_per_hour: 5,
            rule_hard_max_per_hour: 20,
            response_per_hour: 10,
            response_hard_max_per_hour: 50,
            strict: false,
        }
    }
}

impl ActionSettings {
    pub fn from_env() -> Result<Self, String> {
        Self::from_values(|name| std::env::var(name).ok())
    }

    pub fn from_values(get: impl Fn(&str) -> Option<String>) -> Result<Self, String> {
        let number = |name: &str, default: usize| -> Result<usize, String> {
            match get(name).map(|v| v.trim().to_string()) {
                None => Ok(default),
                Some(v) if v.is_empty() => Ok(default),
                Some(v) => v
                    .parse::<usize>()
                    .ok()
                    .filter(|n| *n >= 1)
                    .ok_or_else(|| format!("{name} must be a whole number of at least 1")),
            }
        };
        let defaults = Self::default();
        let settings = Self {
            delete_per_hour: number("OUTLOOK_MCP_DELETE_PER_HOUR", defaults.delete_per_hour)?,
            delete_hard_max_per_hour: number(
                "OUTLOOK_MCP_DELETE_HARD_MAX_PER_HOUR",
                defaults.delete_hard_max_per_hour,
            )?,
            rule_per_hour: number("OUTLOOK_MCP_RULE_PER_HOUR", defaults.rule_per_hour)?,
            rule_hard_max_per_hour: number(
                "OUTLOOK_MCP_RULE_HARD_MAX_PER_HOUR",
                defaults.rule_hard_max_per_hour,
            )?,
            response_per_hour: number("OUTLOOK_MCP_RESPONSE_PER_HOUR", defaults.response_per_hour)?,
            response_hard_max_per_hour: number(
                "OUTLOOK_MCP_RESPONSE_HARD_MAX_PER_HOUR",
                defaults.response_hard_max_per_hour,
            )?,
            strict: matches!(
                get("OUTLOOK_MCP_ACTION_STRICT")
                    .map(|v| v.trim().to_ascii_lowercase())
                    .as_deref(),
                Some("1" | "true" | "yes")
            ),
        };
        if settings.delete_hard_max_per_hour < settings.delete_per_hour
            || settings.rule_hard_max_per_hour < settings.rule_per_hour
            || settings.response_hard_max_per_hour < settings.response_per_hour
        {
            return Err("hard action ceilings must not be below the usual limits".into());
        }
        Ok(settings)
    }

    /// `(usual limit, hard ceiling)` for one kind of action.
    fn limits(&self, kind: ActionKind) -> (usize, usize) {
        match kind.class() {
            "rules" => (self.rule_per_hour, self.rule_hard_max_per_hour),
            "responses" => (self.response_per_hour, self.response_hard_max_per_hour),
            _ => (self.delete_per_hour, self.delete_hard_max_per_hour),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ActionKind {
    DeleteMail,
    DeleteEvent,
    ToggleRule,
    RespondEvent,
}

impl ActionKind {
    pub fn name(self) -> &'static str {
        match self {
            ActionKind::DeleteMail => "delete_mail",
            ActionKind::DeleteEvent => "delete_event",
            ActionKind::ToggleRule => "toggle_rule",
            ActionKind::RespondEvent => "respond_event",
        }
    }

    /// Actions in one class share an hourly counter.
    fn class(self) -> &'static str {
        match self {
            ActionKind::DeleteMail | ActionKind::DeleteEvent => "deletes",
            ActionKind::ToggleRule => "rules",
            ActionKind::RespondEvent => "responses",
        }
    }

    fn label(self) -> &'static str {
        match self.class() {
            "rules" => "rule changes",
            "responses" => "meeting responses",
            _ => "deletes",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ActionTarget {
    pub kind: ActionKind,
    /// Short session handle or raw EntryID, or the rule name for rule changes.
    pub entry_id: String,
    pub store_id: Option<String>,
    /// For rule changes: the state being asked for.
    pub enabled: Option<bool>,
    /// For meeting responses: `accept`, `tentative` or `decline` (`enabled` then
    /// says whether a reply is emailed to the organizer).
    pub response: Option<String>,
}

impl ActionTarget {
    /// The permission this request needs. Telling an organizer needs `send`; a
    /// local-only answer only changes the user's own calendar.
    pub fn capability(&self) -> Capability {
        match self.kind {
            ActionKind::DeleteMail | ActionKind::DeleteEvent => Capability::Delete,
            ActionKind::ToggleRule => Capability::Rules,
            ActionKind::RespondEvent if self.enabled == Some(false) => Capability::Events,
            ActionKind::RespondEvent => Capability::Send,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Fact {
    pub label: String,
    pub value: String,
}

/// What the person sees, read from Outlook (never from the model).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ActionPreview {
    pub action_id: String,
    pub summary: String,
    pub facts: Vec<Fact>,
    pub warnings: Vec<String>,
    /// Identity and facts of the target, rechecked just before acting.
    pub hash: String,
}

#[derive(Debug, Clone)]
pub struct PreparedAction {
    pub target: ActionTarget,
    pub preview: ActionPreview,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Located {
    Done,
    NotDone,
    Unknown,
}

pub trait ActionPort {
    /// Read the facts, apply the pre-prompt refusals (for example "already in
    /// Deleted Items"), stamp the item and build the preview. Errors here mean
    /// nothing happened and no prompt is shown.
    async fn prepare(&self, target: &ActionTarget) -> Result<PreparedAction, String>;
    /// Recheck the hash and perform the one action.
    async fn execute(&self, prepared: &PreparedAction) -> Result<(), PortError>;
    /// Find out what happened to a stamped action.
    async fn locate(
        &self,
        action_id: &str,
        target: Option<&ActionTarget>,
    ) -> Result<Located, String>;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Acting,
    Done,
    Unknown,
}

#[derive(Default)]
struct Inner {
    state: HashMap<String, State>,
    targets: HashMap<String, ActionTarget>,
    /// When each action was attempted, per class, for the hourly window.
    attempts: HashMap<&'static str, Vec<Instant>>,
    /// Human-approved extra actions: (when approved, how many), per class.
    grants: HashMap<&'static str, Vec<(Instant, usize)>>,
}

/// Shared across tool calls so one process has one action history.
#[derive(Default)]
pub struct ActionLedger(Mutex<Inner>);

pub struct ActionEngine<P, C> {
    port: P,
    confirmer: C,
    settings: ActionSettings,
    ledger: Arc<ActionLedger>,
    clock: Box<dyn Fn() -> Instant + Send + Sync>,
}

fn outcome(error: bool, value: serde_json::Value) -> Outcome {
    Outcome { value, error }
}

pub fn build_prompt(preview: &ActionPreview, issues: &[String]) -> String {
    let mut text = format!("Approve this action?\n\n{}\n", preview.summary);
    if !issues.is_empty() {
        text.push_str("\nOUTSIDE THE USUAL LIMITS:\n");
        for issue in issues {
            text.push_str(&format!("  - {issue}\n"));
        }
    }
    text.push('\n');
    for fact in &preview.facts {
        text.push_str(&format!("{}: {}\n", fact.label, fact.value));
    }
    for warning in &preview.warnings {
        text.push_str(&format!("\nNote: {warning}\n"));
    }
    text
}

impl<P: ActionPort, C: Confirmer> ActionEngine<P, C> {
    pub fn new(port: P, confirmer: C, settings: ActionSettings) -> Self {
        Self::build(
            port,
            confirmer,
            settings,
            Arc::default(),
            Box::new(Instant::now),
        )
    }

    /// Production constructor: the ledger outlives any one request.
    pub fn with_ledger(
        port: P,
        confirmer: C,
        settings: ActionSettings,
        ledger: Arc<ActionLedger>,
    ) -> Self {
        Self::build(port, confirmer, settings, ledger, Box::new(Instant::now))
    }

    pub fn with_clock(
        port: P,
        confirmer: C,
        settings: ActionSettings,
        clock: Box<dyn Fn() -> Instant + Send + Sync>,
    ) -> Self {
        Self::build(port, confirmer, settings, Arc::default(), clock)
    }

    fn build(
        port: P,
        confirmer: C,
        settings: ActionSettings,
        ledger: Arc<ActionLedger>,
        clock: Box<dyn Fn() -> Instant + Send + Sync>,
    ) -> Self {
        Self {
            port,
            confirmer,
            settings,
            ledger,
            clock,
        }
    }

    pub async fn run(&self, target: ActionTarget) -> Outcome {
        let kind = target.kind;
        let class = kind.class();
        let (usual, hard) = self.settings.limits(kind);

        // A target whose last attempt has an unknown outcome stays blocked
        // until outlook_check_action says what happened.
        {
            let ledger = self.ledger.0.lock().unwrap();
            let blocked = ledger.state.iter().any(|(id, state)| {
                *state == State::Unknown
                    && ledger
                        .targets
                        .get(id)
                        .is_some_and(|t| t.kind == kind && t.entry_id == target.entry_id)
            });
            if blocked {
                return outcome(
                    true,
                    json!({
                        "status": "refused", "done": false,
                        "reason": "an earlier attempt on this item has an unknown outcome; call outlook_check_action and ask the user first",
                    }),
                );
            }
        }

        let prepared = match self.port.prepare(&target).await {
            Ok(p) => p,
            Err(reason) => {
                return outcome(
                    true,
                    json!({"status": "refused", "done": false, "reason": reason}),
                );
            }
        };
        let id = prepared.preview.action_id.clone();
        let now = (self.clock)();
        let (used, allowance) = {
            let mut ledger = self.ledger.0.lock().unwrap();
            let attempts = ledger.attempts.entry(class).or_default();
            attempts.retain(|t| now.duration_since(*t) < HOUR);
            let used = attempts.len();
            let grants = ledger.grants.entry(class).or_default();
            grants.retain(|(t, _)| now.duration_since(*t) < HOUR);
            let extra: usize = grants.iter().map(|(_, n)| n).sum();
            (used, (usual + extra).min(hard))
        };
        if used >= hard {
            return outcome(
                true,
                json!({"status": "refused", "done": false, "action_id": id,
                    "reason": format!("{used} {} this hour reached the hard ceiling of {hard}; change the server setting if intended", kind.label())}),
            );
        }
        let mut issues = Vec::new();
        let needs_grant = used >= allowance;
        if needs_grant {
            issues.push(format!(
                "{used} {} already this hour (usual limit {usual}); approving allows {usual} more",
                kind.label()
            ));
        }
        if self.settings.strict && !issues.is_empty() {
            return outcome(
                true,
                json!({"status": "refused", "done": false, "action_id": id,
                    "reason": format!("strict mode refuses: {}", issues.join("; "))}),
            );
        }

        match self
            .confirmer
            .confirm(&build_prompt(&prepared.preview, &issues))
            .await
        {
            Confirmation::Accepted => {}
            Confirmation::Declined => {
                return outcome(
                    false,
                    json!({"status": "not_done", "done": false, "action_id": id, "reason": "declined_by_user"}),
                );
            }
            Confirmation::Unsupported => {
                return outcome(
                    false,
                    json!({"status": "not_done", "done": false, "action_id": id,
                        "reason": "client_cannot_confirm",
                        "note": "This client cannot show the approval prompt, so nothing was changed."}),
                );
            }
        }

        let attempt_at = (self.clock)();
        {
            let mut ledger = self.ledger.0.lock().unwrap();
            if ledger.state.contains_key(&id) {
                return outcome(
                    true,
                    json!({"status": "refused", "done": false, "reason": "this action is already being handled"}),
                );
            }
            ledger.state.insert(id.clone(), State::Acting);
            ledger.targets.insert(id.clone(), target.clone());
            ledger.attempts.entry(class).or_default().push(attempt_at);
            if needs_grant {
                ledger
                    .grants
                    .entry(class)
                    .or_default()
                    .push((attempt_at, usual));
            }
        }

        let result = self.port.execute(&prepared).await;
        let mut ledger = self.ledger.0.lock().unwrap();
        match result {
            Ok(()) => {
                ledger.state.insert(id.clone(), State::Done);
                outcome(
                    false,
                    json!({"status": "done", "done": true, "action_id": id,
                        "kind": kind.name(), "summary": prepared.preview.summary}),
                )
            }
            Err(PortError::NotSent(why)) => {
                // Provably not done: free the slot so a fresh request can try again.
                ledger.state.remove(&id);
                ledger.targets.remove(&id);
                ledger.attempts.entry(class).or_default().pop();
                outcome(
                    true,
                    json!({"status": "not_done", "done": false, "action_id": id, "reason": why}),
                )
            }
            Err(PortError::Unknown(why)) => {
                ledger.state.insert(id.clone(), State::Unknown);
                outcome(
                    true,
                    json!({"status": "unknown_outcome", "done": null, "action_id": id, "detail": why,
                        "next_step": "Call outlook_check_action with this action_id, then ask the user. Do not try again automatically."}),
                )
            }
        }
    }

    pub async fn check(&self, action_id: &str) -> Outcome {
        let target = self
            .ledger
            .0
            .lock()
            .unwrap()
            .targets
            .get(action_id)
            .cloned();
        match self.port.locate(action_id, target.as_ref()).await {
            Ok(located) => {
                let (status, note) = match located {
                    Located::Done => ("done", "The action took effect."),
                    Located::NotDone => (
                        "not_done",
                        "The item is unchanged; the action did not happen.",
                    ),
                    Located::Unknown => (
                        "unknown",
                        "Outlook shows no clear result. Ask the user to look in Outlook.",
                    ),
                };
                let mut ledger = self.ledger.0.lock().unwrap();
                match located {
                    Located::Done => {
                        ledger.state.insert(action_id.to_string(), State::Done);
                    }
                    // Proven not done: the item may be tried again.
                    Located::NotDone => {
                        ledger.state.remove(action_id);
                        ledger.targets.remove(action_id);
                    }
                    Located::Unknown => {}
                }
                outcome(
                    false,
                    json!({"status": status, "action_id": action_id, "note": note}),
                )
            }
            Err(e) => outcome(true, json!({"status": "failed", "reason": e})),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[derive(Clone)]
    struct FakePort {
        prepare_error: Option<String>,
        execute_calls: Arc<AtomicUsize>,
        execute_result: Arc<Mutex<Result<(), PortError>>>,
        located: Arc<Mutex<Located>>,
        counter: Arc<AtomicUsize>,
    }

    impl FakePort {
        fn new() -> Self {
            Self {
                prepare_error: None,
                execute_calls: Arc::default(),
                execute_result: Arc::new(Mutex::new(Ok(()))),
                located: Arc::new(Mutex::new(Located::Done)),
                counter: Arc::default(),
            }
        }
    }

    impl ActionPort for FakePort {
        async fn prepare(&self, target: &ActionTarget) -> Result<PreparedAction, String> {
            if let Some(e) = &self.prepare_error {
                return Err(e.clone());
            }
            let n = self.counter.fetch_add(1, Ordering::SeqCst);
            Ok(PreparedAction {
                target: target.clone(),
                preview: ActionPreview {
                    action_id: format!("act-{n}"),
                    summary: "Move this email to Deleted Items".into(),
                    facts: vec![
                        Fact {
                            label: "From".into(),
                            value: "billing@vendor.example".into(),
                        },
                        Fact {
                            label: "Subject".into(),
                            value: "Invoice 42".into(),
                        },
                    ],
                    warnings: vec![],
                    hash: "h".into(),
                },
            })
        }
        async fn execute(&self, _: &PreparedAction) -> Result<(), PortError> {
            self.execute_calls.fetch_add(1, Ordering::SeqCst);
            self.execute_result.lock().unwrap().clone()
        }
        async fn locate(&self, _: &str, _: Option<&ActionTarget>) -> Result<Located, String> {
            Ok(*self.located.lock().unwrap())
        }
    }

    struct FakeConfirmer {
        answer: Confirmation,
        prompts: Arc<Mutex<Vec<String>>>,
    }
    impl Confirmer for FakeConfirmer {
        async fn confirm(&self, prompt: &str) -> Confirmation {
            self.prompts.lock().unwrap().push(prompt.to_string());
            self.answer
        }
    }

    fn engine(
        port: FakePort,
        answer: Confirmation,
        settings: ActionSettings,
    ) -> (
        ActionEngine<FakePort, FakeConfirmer>,
        Arc<Mutex<Vec<String>>>,
    ) {
        let prompts = Arc::new(Mutex::new(Vec::new()));
        let confirmer = FakeConfirmer {
            answer,
            prompts: Arc::clone(&prompts),
        };
        (ActionEngine::new(port, confirmer, settings), prompts)
    }

    fn target(id: &str) -> ActionTarget {
        ActionTarget {
            kind: ActionKind::DeleteMail,
            entry_id: id.into(),
            store_id: None,
            enabled: None,
            response: None,
        }
    }

    fn status(o: &Outcome) -> &str {
        o.value["status"].as_str().unwrap_or("")
    }

    #[tokio::test]
    async fn approval_performs_the_action_exactly_once() {
        let port = FakePort::new();
        let calls = Arc::clone(&port.execute_calls);
        let (e, prompts) = engine(port, Confirmation::Accepted, ActionSettings::default());
        let o = e.run(target("m1")).await;
        assert_eq!(status(&o), "done");
        assert_eq!(o.value["done"], true);
        assert!(!o.error);
        assert_eq!(calls.load(Ordering::SeqCst), 1);
        let prompt = prompts.lock().unwrap()[0].clone();
        assert!(prompt.contains("From: billing@vendor.example"));
        assert!(prompt.contains("Subject: Invoice 42"));
    }

    #[tokio::test]
    async fn declined_or_unsupported_confirmation_does_nothing() {
        for answer in [Confirmation::Declined, Confirmation::Unsupported] {
            let port = FakePort::new();
            let calls = Arc::clone(&port.execute_calls);
            let (e, _) = engine(port, answer, ActionSettings::default());
            let o = e.run(target("m1")).await;
            assert_eq!(status(&o), "not_done", "{answer:?}");
            assert_eq!(o.value["done"], false);
            assert_eq!(calls.load(Ordering::SeqCst), 0);
        }
    }

    #[tokio::test]
    async fn pre_prompt_refusal_shows_no_prompt_and_does_nothing() {
        let mut port = FakePort::new();
        port.prepare_error =
            Some("already in Deleted Items; permanent deletion is never offered".into());
        let calls = Arc::clone(&port.execute_calls);
        let (e, prompts) = engine(port, Confirmation::Accepted, ActionSettings::default());
        let o = e.run(target("m1")).await;
        assert_eq!(status(&o), "refused");
        assert!(
            o.value["reason"]
                .as_str()
                .unwrap()
                .contains("permanent deletion")
        );
        assert!(prompts.lock().unwrap().is_empty());
        assert_eq!(calls.load(Ordering::SeqCst), 0);
    }

    #[tokio::test]
    async fn unknown_outcome_blocks_the_item_until_check_proves_it_was_not_done() {
        let port = FakePort::new();
        *port.execute_result.lock().unwrap() = Err(PortError::Unknown("connection lost".into()));
        let calls = Arc::clone(&port.execute_calls);
        let located = Arc::clone(&port.located);
        let (e, _) = engine(
            port.clone(),
            Confirmation::Accepted,
            ActionSettings::default(),
        );
        let o = e.run(target("m1")).await;
        assert_eq!(status(&o), "unknown_outcome");
        assert!(o.error && o.value["done"].is_null());
        assert_eq!(calls.load(Ordering::SeqCst), 1);

        // The same item is refused without another prompt or attempt.
        let again = e.run(target("m1")).await;
        assert_eq!(status(&again), "refused");
        assert_eq!(calls.load(Ordering::SeqCst), 1);
        // Other items are unaffected.
        *port.execute_result.lock().unwrap() = Ok(());
        assert_eq!(status(&e.run(target("m2")).await), "done");

        // Checking proves it did not happen, so the item may be tried again.
        *located.lock().unwrap() = Located::NotDone;
        assert_eq!(status(&e.check("act-0").await), "not_done");
        assert_eq!(status(&e.run(target("m1")).await), "done");
    }

    #[tokio::test]
    async fn changed_target_is_not_done_and_frees_the_slot() {
        let port = FakePort::new();
        *port.execute_result.lock().unwrap() = Err(PortError::NotSent("item changed".into()));
        let (e, _) = engine(
            port.clone(),
            Confirmation::Accepted,
            ActionSettings::default(),
        );
        let o = e.run(target("m1")).await;
        assert_eq!(status(&o), "not_done");
        assert_eq!(o.value["done"], false);
        *port.execute_result.lock().unwrap() = Ok(());
        assert_eq!(status(&e.run(target("m1")).await), "done");
    }

    #[tokio::test]
    async fn hourly_threshold_prompts_grants_in_steps_and_stops_at_the_ceiling() {
        let settings = ActionSettings {
            delete_per_hour: 2,
            delete_hard_max_per_hour: 5,
            strict: false,
            ..ActionSettings::default()
        };
        let port = FakePort::new();
        let calls = Arc::clone(&port.execute_calls);
        let (e, prompts) = engine(port, Confirmation::Accepted, settings);
        let mut statuses = Vec::new();
        for i in 0..6 {
            statuses.push(status(&e.run(target(&format!("m{i}"))).await).to_string());
        }
        assert_eq!(
            statuses,
            ["done", "done", "done", "done", "done", "refused"]
        );
        assert_eq!(calls.load(Ordering::SeqCst), 5);
        let prompts = prompts.lock().unwrap();
        assert!(!prompts[0].contains("OUTSIDE THE USUAL LIMITS"));
        assert!(prompts[2].contains("deletes already this hour"));
    }

    fn rule_target(name: &str, on: bool) -> ActionTarget {
        ActionTarget {
            kind: ActionKind::ToggleRule,
            entry_id: name.into(),
            store_id: None,
            enabled: Some(on),
            response: None,
        }
    }

    #[tokio::test]
    async fn rule_changes_have_their_own_counter_and_thresholds() {
        let settings = ActionSettings {
            delete_per_hour: 1,
            delete_hard_max_per_hour: 5,
            rule_per_hour: 2,
            rule_hard_max_per_hour: 3,
            strict: false,
            ..ActionSettings::default()
        };
        let port = FakePort::new();
        let calls = Arc::clone(&port.execute_calls);
        let (e, prompts) = engine(port, Confirmation::Accepted, settings);
        // Deleting does not use up rule changes.
        assert_eq!(status(&e.run(target("m1")).await), "done");
        let mut statuses = Vec::new();
        for i in 0..4 {
            statuses
                .push(status(&e.run(rule_target(&format!("rule {i}"), true)).await).to_string());
        }
        assert_eq!(statuses, ["done", "done", "done", "refused"]);
        assert_eq!(calls.load(Ordering::SeqCst), 4);
        let prompts = prompts.lock().unwrap();
        // Prompt 0 is the delete; prompt 3 is the third rule change, over the usual 2.
        assert!(prompts[3].contains("rule changes already this hour"));
        assert!(!prompts[1].contains("OUTSIDE THE USUAL LIMITS"));
    }

    #[tokio::test]
    async fn meeting_responses_have_their_own_counter() {
        let settings = ActionSettings {
            response_per_hour: 1,
            response_hard_max_per_hour: 2,
            ..ActionSettings::default()
        };
        let port = FakePort::new();
        let (e, prompts) = engine(port, Confirmation::Accepted, settings);
        let respond = |n: usize| ActionTarget {
            kind: ActionKind::RespondEvent,
            entry_id: format!("m{n}"),
            store_id: None,
            enabled: Some(true),
            response: Some("accept".into()),
        };
        // Deletes and rule changes do not use up meeting responses.
        assert_eq!(status(&e.run(target("m9")).await), "done");
        assert_eq!(status(&e.run(rule_target("r", true)).await), "done");
        assert_eq!(status(&e.run(respond(1)).await), "done");
        assert_eq!(status(&e.run(respond(2)).await), "done");
        assert_eq!(status(&e.run(respond(3)).await), "refused");
        let prompts = prompts.lock().unwrap();
        assert!(prompts[3].contains("meeting responses already this hour"));
    }

    #[test]
    fn kinds_map_to_their_capabilities() {
        let of = |kind, enabled| {
            ActionTarget {
                kind,
                entry_id: "x".into(),
                store_id: None,
                enabled,
                response: None,
            }
            .capability()
        };
        assert_eq!(of(ActionKind::DeleteMail, None), Capability::Delete);
        assert_eq!(of(ActionKind::DeleteEvent, None), Capability::Delete);
        assert_eq!(of(ActionKind::ToggleRule, Some(true)), Capability::Rules);
        assert_eq!(of(ActionKind::RespondEvent, Some(true)), Capability::Send);
        assert_eq!(of(ActionKind::RespondEvent, None), Capability::Send);
        assert_eq!(
            of(ActionKind::RespondEvent, Some(false)),
            Capability::Events
        );
    }

    #[tokio::test]
    async fn strict_mode_refuses_at_the_threshold_without_prompting() {
        let settings = ActionSettings {
            delete_per_hour: 1,
            delete_hard_max_per_hour: 5,
            strict: true,
            ..ActionSettings::default()
        };
        let port = FakePort::new();
        let (e, prompts) = engine(port, Confirmation::Accepted, settings);
        assert_eq!(status(&e.run(target("m1")).await), "done");
        assert_eq!(status(&e.run(target("m2")).await), "refused");
        assert_eq!(prompts.lock().unwrap().len(), 1);
    }

    #[tokio::test]
    async fn hourly_window_expires() {
        let settings = ActionSettings {
            delete_per_hour: 1,
            delete_hard_max_per_hour: 5,
            strict: false,
            ..ActionSettings::default()
        };
        let now = Arc::new(Mutex::new(Instant::now()));
        let clock_now = Arc::clone(&now);
        let prompts = Arc::new(Mutex::new(Vec::new()));
        let engine = ActionEngine::with_clock(
            FakePort::new(),
            FakeConfirmer {
                answer: Confirmation::Accepted,
                prompts: Arc::clone(&prompts),
            },
            settings,
            Box::new(move || *clock_now.lock().unwrap()),
        );
        engine.run(target("m1")).await;
        *now.lock().unwrap() += Duration::from_secs(3601);
        engine.run(target("m2")).await;
        assert!(!prompts.lock().unwrap()[1].contains("already this hour"));
    }

    #[tokio::test]
    async fn check_reports_each_state() {
        for (located, expected) in [
            (Located::Done, "done"),
            (Located::NotDone, "not_done"),
            (Located::Unknown, "unknown"),
        ] {
            let port = FakePort::new();
            *port.located.lock().unwrap() = located;
            let (e, _) = engine(port, Confirmation::Accepted, ActionSettings::default());
            assert_eq!(status(&e.check("act-0").await), expected);
        }
    }

    #[test]
    fn settings_parse_and_fail_closed() {
        let get = |pairs: &'static [(&str, &str)]| {
            move |name: &str| {
                pairs
                    .iter()
                    .find(|(k, _)| *k == name)
                    .map(|(_, v)| v.to_string())
            }
        };
        assert_eq!(
            ActionSettings::from_values(get(&[])).unwrap(),
            ActionSettings::default()
        );
        let s = ActionSettings::from_values(get(&[
            ("OUTLOOK_MCP_DELETE_PER_HOUR", "20"),
            ("OUTLOOK_MCP_ACTION_STRICT", "yes"),
        ]))
        .unwrap();
        assert_eq!(s.delete_per_hour, 20);
        assert!(s.strict);
        assert!(ActionSettings::from_values(get(&[("OUTLOOK_MCP_DELETE_PER_HOUR", "0")])).is_err());
        assert!(ActionSettings::from_values(get(&[("OUTLOOK_MCP_DELETE_PER_HOUR", "x")])).is_err());
        assert!(
            ActionSettings::from_values(get(&[("OUTLOOK_MCP_DELETE_PER_HOUR", "60")])).is_err(),
            "limit above the hard ceiling"
        );
    }
}
