//! Local Outlook Classic MCP tools.
//!
//! All mailbox operations execute on the dedicated Windows COM STA worker.
//! Mutating tools exist only behind the server-side permission policy, and no
//! Outlook data is logged to stderr.
pub mod action;
pub mod approval;
pub mod contracts;
#[cfg(windows)]
pub mod dialog;
pub mod localtime;
pub mod outlook;
pub mod policy;
pub mod safe_path;
pub mod send;
pub mod thread_text;
pub mod ticket;
pub mod ui;

use action::{
    ActionEngine, ActionKind, ActionLedger, ActionPort, ActionPreview, ActionSettings,
    ActionTarget, Located, PreparedAction,
};
use approval::{ApprovalMode, Backend};
use contracts::{
    CheckActionParams, CheckSendParams, CompleteTaskParams, CreateEventParams, CreateFolderParams,
    CreateTaskParams, DeleteItemParams, FolderParams, ForwardMailParams, GetContactParams,
    GetEventParams, GetMailParams, ListContactsParams, ListEventsParams, ListMailsParams,
    ListTasksParams, MailCommand, MarkMailParams, MoveMailParams, ReplyMailParams,
    ResolveNameParams, RespondEventParams, SaveAttachmentsParams, SearchContactsParams,
    SearchMailsParams, SendDraftParams, SendMailParams, SetCategoryParams, ThreadParams,
    ToggleRuleParams, UpdateEventParams, checked_send_id, prepare_cancel_meeting, prepare_delete,
    prepare_forward_mail, prepare_invite_meeting, prepare_reply_mail, prepare_respond_event,
    prepare_save_attachments, prepare_send_mail, prepare_toggle_rule, prepare_update_meeting,
};
use outlook::OutlookBridge;
use policy::{Capability, Policy};
use rmcp::{
    Peer, RoleServer, ServerHandler, elicit_safe,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock},
    schemars::JsonSchema,
    service::ElicitationError,
    tool, tool_handler, tool_router,
};
use send::{
    Confirmation, Confirmer, Delivery, DraftRef, Engine, Location, MailPort, Outcome, PortError,
    PrepareRequest, Prepared, Preview, SendLedger, SendSettings,
};
use serde::{Deserialize, Serialize};
use std::sync::Arc;
use ticket::NOT_SENT;
use ticket::UNKNOWN_OUTCOME;

/// The Rust implementation is experimental until tested against a real
/// Outlook Classic instance. A successful CI build is not a live COM test.
#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub struct FoundationStatus {
    pub implementation: &'static str,
    pub version: &'static str,
    pub stage: &'static str,
    pub mailbox_tools_available: bool,
    pub outlook_connected: bool,
}

pub fn foundation_status() -> FoundationStatus {
    FoundationStatus {
        implementation: "rust",
        version: env!("CARGO_PKG_VERSION"),
        stage: "read-only-experimental",
        mailbox_tools_available: true,
        // An unconnected status probe must not claim Outlook is available.
        outlook_connected: false,
    }
}

#[derive(Clone)]
pub struct OutlookMcp {
    bridge: Arc<OutlookBridge>,
    policy: Policy,
    send_settings: SendSettings,
    ledger: Arc<SendLedger>,
    action_settings: ActionSettings,
    action_ledger: Arc<ActionLedger>,
    approval: ApprovalMode,
}

/// The form the host shows the human. Only a person at the host can fill it in.
#[derive(Debug, Deserialize, JsonSchema)]
struct SendDecision {
    /// Tick to send this email now. Leave unticked to keep it as a draft.
    confirm: bool,
}
elicit_safe!(SendDecision);

const APPROVAL_WAIT: std::time::Duration = std::time::Duration::from_secs(120);

/// Asks a person through the backend chosen by `OUTLOOK_MCP_APPROVAL`.
struct ApprovalConfirmer {
    peer: Peer<RoleServer>,
    mode: ApprovalMode,
}

impl ApprovalConfirmer {
    async fn elicit(&self, prompt: &str) -> Confirmation {
        match self
            .peer
            .elicit_with_timeout::<SendDecision>(prompt, Some(APPROVAL_WAIT))
            .await
        {
            Ok(Some(decision)) if decision.confirm => Confirmation::Accepted,
            Err(ElicitationError::CapabilityNotSupported) => Confirmation::Unsupported,
            // Declined, cancelled, timed out or malformed: never send.
            _ => Confirmation::Declined,
        }
    }

    #[cfg(windows)]
    async fn dialog(prompt: &str) -> Confirmation {
        // One window at a time, so a second request cannot be answered by mistake.
        static ONE_AT_A_TIME: std::sync::Mutex<()> = std::sync::Mutex::new(());
        let (summary, details) = approval::split_prompt(prompt);
        let request = dialog::Request {
            title: "Outlook MCP: approval needed".into(),
            headline: "Approve this Outlook action?".into(),
            summary,
            details,
            timeout: APPROVAL_WAIT,
        };
        let decision = tokio::task::spawn_blocking(move || {
            let _guard = ONE_AT_A_TIME.lock().unwrap_or_else(|e| e.into_inner());
            dialog::approve(request)
        })
        .await;
        match decision {
            Ok(dialog::Decision::Approved) => Confirmation::Accepted,
            // Closed, declined, timed out or failed to show: never act.
            _ => Confirmation::Declined,
        }
    }

    #[cfg(not(windows))]
    async fn dialog(_prompt: &str) -> Confirmation {
        Confirmation::Unsupported
    }
}

impl Confirmer for ApprovalConfirmer {
    async fn confirm(&self, prompt: &str) -> Confirmation {
        match self.mode.choose(client_elicits(&self.peer)) {
            Backend::Elicitation => self.elicit(prompt).await,
            Backend::Dialog => Self::dialog(prompt).await,
            Backend::Unavailable => Confirmation::Unsupported,
        }
    }
}

fn client_elicits(peer: &Peer<RoleServer>) -> bool {
    peer.supported_elicitation_modes()
        .contains(&rmcp::service::ElicitationMode::Form)
}

/// The send pipeline's view of Outlook: each step is one short call on the COM worker.
struct BridgePort<'a>(&'a OutlookMcp);

/// The confirmed-action engine's view of Outlook.
struct BridgeActionPort<'a>(&'a OutlookMcp);

impl ActionPort for BridgeActionPort<'_> {
    async fn prepare(&self, target: &ActionTarget) -> Result<PreparedAction, String> {
        let value = self
            .0
            .raw_call(MailCommand::ActionPrepare(target.clone()))
            .await?;
        let preview: ActionPreview =
            serde_json::from_value(value["preview"].clone()).map_err(|e| e.to_string())?;
        Ok(PreparedAction {
            target: target.clone(),
            preview,
        })
    }

    async fn execute(&self, prepared: &PreparedAction) -> Result<(), PortError> {
        let command = MailCommand::ActionExecute {
            target: prepared.target.clone(),
            preview: prepared.preview.clone(),
        };
        match self.0.raw_call(command).await {
            Ok(_) => Ok(()),
            Err(text) => {
                if let Some(why) = text.strip_prefix(NOT_SENT) {
                    Err(PortError::NotSent(why.to_string()))
                } else if text.starts_with("permission denied") {
                    Err(PortError::NotSent(text))
                } else if let Some(why) = text.strip_prefix(UNKNOWN_OUTCOME) {
                    Err(PortError::Unknown(why.to_string()))
                } else {
                    // Anything else after dispatch is treated as unknown, never as a clean failure.
                    Err(PortError::Unknown(text))
                }
            }
        }
    }

    async fn locate(
        &self,
        action_id: &str,
        target: Option<&ActionTarget>,
    ) -> Result<Located, String> {
        let value = self
            .0
            .raw_call(MailCommand::ActionLocate {
                action_id: action_id.to_string(),
                target: target.cloned(),
            })
            .await?;
        Ok(match value["located"].as_str() {
            Some("done") => Located::Done,
            Some("not_done") => Located::NotDone,
            _ => Located::Unknown,
        })
    }
}

fn parse_prepared(value: &serde_json::Value) -> Result<Prepared, String> {
    let entry_id = value["entry_id"]
        .as_str()
        .ok_or("missing draft id")?
        .to_string();
    let preview: Preview =
        serde_json::from_value(value["preview"].clone()).map_err(|e| e.to_string())?;
    Ok(Prepared {
        draft: DraftRef {
            entry_id,
            store_id: None,
        },
        preview,
    })
}

impl MailPort for BridgePort<'_> {
    async fn prepare(&self, request: PrepareRequest) -> Result<Prepared, String> {
        parse_prepared(
            &self
                .0
                .raw_call(MailCommand::ComposePrepare(request))
                .await?,
        )
    }

    async fn inspect(&self, draft: &DraftRef) -> Result<Prepared, String> {
        parse_prepared(
            &self
                .0
                .raw_call(MailCommand::ComposeInspect(draft.clone()))
                .await?,
        )
    }

    async fn send(&self, draft: &DraftRef, preview: &Preview) -> Result<(), PortError> {
        let command = MailCommand::ComposeSend {
            draft: draft.clone(),
            preview: preview.clone(),
        };
        match self.0.raw_call(command).await {
            Ok(_) => Ok(()),
            Err(text) => {
                if let Some(why) = text.strip_prefix(NOT_SENT) {
                    Err(PortError::NotSent(why.to_string()))
                } else if text.starts_with("permission denied") {
                    Err(PortError::NotSent(text))
                } else if let Some(why) = text.strip_prefix(UNKNOWN_OUTCOME) {
                    Err(PortError::Unknown(why.to_string()))
                } else {
                    // Anything else after dispatch is treated as unknown, never as a clean failure.
                    Err(PortError::Unknown(text))
                }
            }
        }
    }

    async fn locate(&self, send_id: &str) -> Result<Location, String> {
        let value = self
            .0
            .raw_call(MailCommand::ComposeLocate(send_id.to_string()))
            .await?;
        Ok(match value["location"].as_str() {
            Some("sent") => Location::Sent,
            Some("queued") => Location::Queued,
            Some("drafts") => Location::Drafts,
            _ => Location::Unknown,
        })
    }
}

fn outcome_result(outcome: Outcome) -> CallToolResult {
    let content = vec![ContentBlock::text(outcome.value.to_string())];
    if outcome.error {
        CallToolResult::error(content)
    } else {
        CallToolResult::success(content)
    }
}

fn text_error(message: impl Into<String>) -> CallToolResult {
    CallToolResult::error(vec![ContentBlock::text(message.into())])
}

/// Sending needs `send`. Without it, `draft` still allows saving the message
/// for the user to send from Outlook; otherwise nothing is allowed.
pub fn delivery_for(policy: &Policy, save_only: bool) -> Result<Delivery, String> {
    if !save_only && policy.allows(Capability::Send) {
        return Ok(Delivery::Send);
    }
    if policy.allows(Capability::Draft) {
        return Ok(Delivery::SaveOnly);
    }
    policy
        .check(if save_only {
            Capability::Draft
        } else {
            Capability::Send
        })
        .map(|_| Delivery::SaveOnly)
}

impl OutlookMcp {
    /// Read-only, the safest mode. Use `with_policy` to opt into more.
    pub fn new() -> Self {
        Self::with_policy(Policy::read_only())
    }

    pub fn with_policy(policy: Policy) -> Self {
        Self::with_settings(policy, SendSettings::default())
    }

    pub fn with_settings(policy: Policy, send_settings: SendSettings) -> Self {
        Self {
            bridge: Arc::new(OutlookBridge::start()),
            policy,
            send_settings,
            ledger: Arc::default(),
            action_settings: ActionSettings::default(),
            action_ledger: Arc::default(),
            approval: ApprovalMode::Dialog,
        }
    }

    /// Replace the delete thresholds (they are prompts, with hard ceilings).
    pub fn with_action_settings(mut self, action_settings: ActionSettings) -> Self {
        self.action_settings = action_settings;
        self
    }

    /// Choose how approval prompts are shown (`OUTLOOK_MCP_APPROVAL`).
    pub fn with_approval(mut self, approval: ApprovalMode) -> Self {
        self.approval = approval;
        self
    }

    fn action_engine(
        &self,
        peer: Peer<RoleServer>,
    ) -> ActionEngine<BridgeActionPort<'_>, ApprovalConfirmer> {
        ActionEngine::with_ledger(
            BridgeActionPort(self),
            ApprovalConfirmer {
                peer,
                mode: self.approval,
            },
            self.action_settings.clone(),
            Arc::clone(&self.action_ledger),
        )
    }

    fn engine(&self, peer: Peer<RoleServer>) -> Engine<BridgePort<'_>, ApprovalConfirmer> {
        Engine::with_ledger(
            BridgePort(self),
            ApprovalConfirmer {
                peer,
                mode: self.approval,
            },
            self.send_settings.clone(),
            Arc::clone(&self.ledger),
        )
    }

    /// Policy check, then one request on the COM worker. Errors are plain text.
    async fn raw_call(&self, cmd: MailCommand) -> Result<serde_json::Value, String> {
        // Server-side gate: nothing the model sends can change the policy.
        self.policy.check(cmd.capability())?;
        let bridge = Arc::clone(&self.bridge);
        // COM/Windows calls must NEVER run directly on the async MCP executor.
        match tokio::task::spawn_blocking(move || bridge.execute(cmd)).await {
            Ok(Ok(value)) => Ok(value),
            Ok(Err(error)) => Err(error.to_string()),
            Err(error) => Err(format!("Outlook worker task failed: {error}")),
        }
    }

    async fn confirmed_delete(
        &self,
        peer: Peer<RoleServer>,
        params: DeleteItemParams,
        kind: ActionKind,
    ) -> CallToolResult {
        match prepare_delete(&params, kind) {
            Ok(target) => self.confirmed(peer, target).await,
            Err(error) => text_error(error),
        }
    }

    /// Permission for this kind of action, then the approval-gated engine.
    async fn confirmed(&self, peer: Peer<RoleServer>, target: ActionTarget) -> CallToolResult {
        if let Err(error) = self.policy.check(target.capability()) {
            return text_error(error);
        }
        outcome_result(self.action_engine(peer).run(target).await)
    }

    /// Like `call`, plus the parsed result as `structuredContent` for MCP Apps views.
    async fn call_ui(&self, cmd: MailCommand) -> CallToolResult {
        let mut result = self.call(cmd).await;
        if result.is_error != Some(true) {
            let parsed = result
                .content
                .first()
                .and_then(|c| c.as_text())
                .and_then(|t| serde_json::from_str::<serde_json::Value>(&t.text).ok());
            if let Some(value @ serde_json::Value::Object(_)) = parsed {
                result.structured_content = Some(value);
            }
        }
        result
    }

    async fn call(&self, cmd: MailCommand) -> CallToolResult {
        match self.raw_call(cmd).await {
            Ok(value) => CallToolResult::success(vec![ContentBlock::text(value.to_string())]),
            Err(text) => {
                match text.strip_prefix(UNKNOWN_OUTCOME) {
                    // A write may or may not have happened; never present it as a plain failure.
                    Some(detail) => CallToolResult::error(vec![ContentBlock::text(
                        serde_json::json!({
                            "status": "unknown_outcome",
                            "detail": detail,
                            "next_step": "Check Outlook to see whether the operation took effect. Do not retry automatically."
                        })
                        .to_string(),
                    )]),
                    None => CallToolResult::error(vec![ContentBlock::text(text)]),
                }
            }
        }
    }
}

impl Default for OutlookMcp {
    fn default() -> Self {
        Self::new()
    }
}

#[tool_router]
impl OutlookMcp {
    #[tool(
        description = "Return the Rust server stage and capability flags; does not check the connected Outlook session."
    )]
    fn outlook_rust_status(&self, peer: Peer<RoleServer>) -> String {
        let mut status = serde_json::to_value(foundation_status())
            .expect("serialization of static fields cannot fail");
        status["permission_mode"] = serde_json::json!(self.policy.describe());
        // Without this the host cannot show send confirmations; sends then save drafts.
        status["client_supports_elicitation"] = serde_json::json!(client_elicits(&peer));
        status["approval_mode"] = serde_json::json!(self.approval.name());
        status["approval_backend"] =
            serde_json::json!(format!("{:?}", self.approval.choose(client_elicits(&peer))));
        status.to_string()
    }

    #[tool(
        description = "Read the signed-in Outlook Classic user's display name. Experimental, read-only."
    )]
    async fn outlook_whoami(&self) -> CallToolResult {
        self.call(MailCommand::WhoAmI).await
    }

    #[tool(
        description = "List Outlook Classic mailbox folders. Bounded traversal; defaults to mailbox root with max_depth 2."
    )]
    async fn outlook_list_folders(
        &self,
        Parameters(params): Parameters<FolderParams>,
    ) -> CallToolResult {
        self.call(MailCommand::ListFolders(params)).await
    }

    #[tool(
        description = "List up to 30 metadata-only mail summaries in one folder. EntryID fields are short session-scoped m1-style handles; call get_mail with a returned handle. Backends auto/table/legacy.",
        meta = ui::tool_meta("mail-list")
    )]
    async fn outlook_list_mails(
        &self,
        Parameters(params): Parameters<ListMailsParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::ListMails(params)).await
    }

    #[tool(
        description = "Search one Outlook Classic folder and return compact metadata-only hits with session-scoped m1-style handles instead of long EntryIDs/StoreIDs. Backends auto/table/legacy/indexed (indexed is experimental, requires Instant Search and waits for completion event); elapsed_ms is measured.",
        meta = ui::tool_meta("mail-list")
    )]
    async fn outlook_search_mails(
        &self,
        Parameters(params): Parameters<SearchMailsParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::SearchMails(params)).await
    }

    #[tool(
        description = "Fetch a bounded Outlook conversation containing a mail handle. Uses exact ConversationID (not subject matching); chronological messages, conservative quoted-history trimming, per-message and total character ceilings. Raw originals remain accessible via outlook_get_mail."
    )]
    async fn outlook_get_thread(
        &self,
        Parameters(params): Parameters<ThreadParams>,
    ) -> CallToolResult {
        self.call(MailCommand::GetThread(params)).await
    }

    #[tool(
        description = "Read one Outlook Classic email. Pass the short entry_id handle returned by search/list (e.g. m1); legacy raw EntryIDs also work. Bounded plain-text body; HTML disabled.",
        meta = ui::tool_meta("mail-view")
    )]
    async fn outlook_get_mail(
        &self,
        Parameters(params): Parameters<GetMailParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::GetMail(params)).await
    }

    #[tool(
        description = "List contacts from every contact folder in every store, sorted by name. Bounded (default 50, max 200; scan_truncated reports hitting the scan budget). entry_id is a short session handle; call outlook_get_contact.",
        meta = ui::tool_meta("contacts")
    )]
    async fn outlook_list_contacts(
        &self,
        Parameters(params): Parameters<ListContactsParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::ListContacts(params)).await
    }

    #[tool(
        description = "Search contact folders and, by default, the Exchange Global Address List. All query words must match name, email, company or job title. Bounded scans; scan_truncated reports a partial search.",
        meta = ui::tool_meta("contacts")
    )]
    async fn outlook_search_contacts(
        &self,
        Parameters(params): Parameters<SearchContactsParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::SearchContacts(params)).await
    }

    #[tool(
        description = "Read one contact by handle from list/search (or raw EntryID). Notes are capped at 4000 characters by default (max_notes_chars up to 50000; notes_truncated)."
    )]
    async fn outlook_get_contact(
        &self,
        Parameters(params): Parameters<GetContactParams>,
    ) -> CallToolResult {
        self.call(MailCommand::GetContact(params)).await
    }

    #[tool(
        description = "Resolve a name or address against the address book, like Ctrl+K in Outlook. Unique matches resolve; ambiguous ones return resolved=false. No UI is shown."
    )]
    async fn outlook_resolve_name(
        &self,
        Parameters(params): Parameters<ResolveNameParams>,
    ) -> CallToolResult {
        self.call(MailCommand::ResolveName(params)).await
    }

    #[tool(
        description = "List calendar events in a bounded local-time window (default now to +14 days; limit default 50, max 200). Recurring series are expanded by default. Pass local ISO-8601 times or times with a UTC offset (converted to local); results carry the local offset. entry_id is a short session handle.",
        meta = ui::tool_meta("calendar")
    )]
    async fn outlook_list_events(
        &self,
        Parameters(params): Parameters<ListEventsParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::ListEvents(params)).await
    }

    #[tool(
        description = "Read one calendar event by handle from list_events (or raw EntryID), including attendees. Body capped at 4000 characters by default (max_body_chars up to 50000; body_truncated)."
    )]
    async fn outlook_get_event(
        &self,
        Parameters(params): Parameters<GetEventParams>,
    ) -> CallToolResult {
        self.call(MailCommand::GetEvent(params)).await
    }

    #[tool(
        description = "List tasks from the default Tasks folder sorted by due date (limit default 50, max 200). Completed tasks are excluded unless include_completed is true. Tasks without a date report null.",
        meta = ui::tool_meta("tasks")
    )]
    async fn outlook_list_tasks(
        &self,
        Parameters(params): Parameters<ListTasksParams>,
    ) -> CallToolResult {
        self.call_ui(MailCommand::ListTasks(params)).await
    }

    #[tool(
        description = "Mark a mail item read/unread and/or flagged. Requires modify permission (server mode full). Reports the state Outlook holds afterwards; status is unverified when it differs from the request.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_mark_mail(
        &self,
        Parameters(params): Parameters<MarkMailParams>,
    ) -> CallToolResult {
        self.call(MailCommand::MarkMail(params)).await
    }

    #[tool(
        description = "Replace an item's categories with a semicolon-separated list (empty clears). Requires modify permission (server mode full). Reports the categories Outlook holds afterwards.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_set_category(
        &self,
        Parameters(params): Parameters<SetCategoryParams>,
    ) -> CallToolResult {
        self.call(MailCommand::SetCategory(params)).await
    }

    #[tool(
        description = "Create a mail folder under a parent (default Inbox). Requires modify permission (server mode full). Fails if the name already exists.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_create_folder(
        &self,
        Parameters(params): Parameters<CreateFolderParams>,
    ) -> CallToolResult {
        self.call(MailCommand::CreateFolder(params)).await
    }

    #[tool(
        description = "Move a mail item to another folder. Requires modify permission (server mode full). The returned entry_id is the item's NEW handle; the old one is stale. Moving into Deleted Items is refused.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_move_mail(
        &self,
        Parameters(params): Parameters<MoveMailParams>,
    ) -> CallToolResult {
        self.call(MailCommand::MoveMail(params)).await
    }

    #[tool(
        description = "Compose a plain-text email and send it ONLY after the user approves a preview shown by their client (recipients, subject, body). With save_only, or when the server is not in full mode, it saves to Drafts and says so; check `sent` in the result. Attachments and HTML are not supported yet. Limits (many recipients, many sends per hour) trigger an extra user prompt.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_send_mail(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<SendMailParams>,
    ) -> CallToolResult {
        let request = match prepare_send_mail(&params) {
            Ok(request) => request,
            Err(error) => return text_error(error),
        };
        let delivery = match delivery_for(&self.policy, params.save_only.unwrap_or(false)) {
            Ok(delivery) => delivery,
            Err(error) => return text_error(error),
        };
        outcome_result(self.engine(peer).submit(request, delivery).await)
    }

    #[tool(
        description = "Reply to a message (reply_all optional) in plain text and send it ONLY after the user approves a preview. Otherwise saves to Drafts; check `sent` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_reply_mail(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<ReplyMailParams>,
    ) -> CallToolResult {
        let request = match prepare_reply_mail(&params) {
            Ok(request) => request,
            Err(error) => return text_error(error),
        };
        let delivery = match delivery_for(&self.policy, false) {
            Ok(delivery) => delivery,
            Err(error) => return text_error(error),
        };
        outcome_result(self.engine(peer).submit(request, delivery).await)
    }

    #[tool(
        description = "Forward a message to new recipients in plain text and send it ONLY after the user approves a preview that notes existing content is leaving the mailbox. Otherwise saves to Drafts; check `sent` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_forward_mail(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<ForwardMailParams>,
    ) -> CallToolResult {
        let request = match prepare_forward_mail(&params) {
            Ok(request) => request,
            Err(error) => return text_error(error),
        };
        let delivery = match delivery_for(&self.policy, false) {
            Ok(delivery) => delivery,
            Err(error) => return text_error(error),
        };
        outcome_result(self.engine(peer).submit(request, delivery).await)
    }

    #[tool(
        description = "Send a draft this server saved earlier (use the entry_id it returned), after the user approves a preview. Requires send permission (server mode full). A draft already sent, or one with an unknown outcome, is refused.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_send_draft(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<SendDraftParams>,
    ) -> CallToolResult {
        if let Err(error) = self.policy.check(Capability::Send) {
            return text_error(error);
        }
        if params.entry_id.is_empty() || params.entry_id.len() > 4096 {
            return text_error("invalid EntryID");
        }
        let draft = DraftRef {
            entry_id: params.entry_id,
            store_id: params.store_id,
        };
        outcome_result(self.engine(peer).send_draft(draft).await)
    }

    #[tool(
        description = "After a send reported unknown_outcome (or to confirm one), look for the message by its send_id in Sent Items, Outbox and Drafts. Read-only. Never retry a send before checking.",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_check_send(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<CheckSendParams>,
    ) -> CallToolResult {
        if let Err(error) = self
            .policy
            .check(Capability::Read)
            .and(checked_send_id(&params.send_id))
        {
            return text_error(error);
        }
        outcome_result(self.engine(peer).check(&params.send_id).await)
    }

    #[tool(
        description = "Create a task (subject, optional due_date, body, importance, reminder). Requires modify permission (server mode full). Reads the task back and reports the stored values.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_create_task(
        &self,
        Parameters(params): Parameters<CreateTaskParams>,
    ) -> CallToolResult {
        self.call(MailCommand::CreateTask(params)).await
    }

    #[tool(
        description = "Mark a task complete by handle from list_tasks. Requires modify permission (server mode full). Reports whether Outlook now shows it complete.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_complete_task(
        &self,
        Parameters(params): Parameters<CompleteTaskParams>,
    ) -> CallToolResult {
        self.call(MailCommand::CompleteTask(params)).await
    }

    #[tool(
        description = "Create a calendar event. Without attendees it is on your own calendar only (events permission, server mode full). With `attendees` (plain email addresses) it is a meeting: invitations are emailed only after the user approves a prompt listing every attendee (outside-organization ones marked), and it needs send permission; recurring meetings and online meetings are not supported yet. Nothing is created if the prompt is declined or the client cannot show it. Times are local ISO-8601 or carry a UTC offset.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_create_event(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<CreateEventParams>,
    ) -> CallToolResult {
        if params.attendees.as_ref().is_some_and(|a| !a.is_empty()) {
            return match prepare_invite_meeting(&params) {
                Ok(target) => self.confirmed(peer, target).await,
                Err(error) => text_error(error),
            };
        }
        self.call(MailCommand::CreateEvent(params)).await
    }

    #[tool(
        description = "Change the subject, start, end, location or description of a meeting you organized and email the update to every attendee, only after the user approves a prompt showing each change and every attendee (outside-organization attendees are marked). Refused for events without attendees (use outlook_update_event), meetings you did not organize, cancelled and recurring meetings. Cannot add or remove attendees. Requires send permission (server mode full). Check `done` in the result; on unknown_outcome call outlook_check_action and never retry.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_update_meeting(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<UpdateEventParams>,
    ) -> CallToolResult {
        match prepare_update_meeting(&params) {
            Ok(target) => self.confirmed(peer, target).await,
            Err(error) => text_error(error),
        }
    }

    #[tool(
        description = "Change an event's subject, start, end, location or body. Requires events permission (server mode full). Refused for meetings with attendees (use outlook_update_meeting). Reads the event back.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_update_event(
        &self,
        Parameters(params): Parameters<UpdateEventParams>,
    ) -> CallToolResult {
        self.call(MailCommand::UpdateEvent(params)).await
    }

    #[tool(
        description = "Move ONE mail item to Deleted Items, only after the user approves a prompt showing its sender, subject, date and folder. Nothing is deleted if the client cannot show the prompt or the user declines. Items already in Deleted Items are refused (permanent deletion is never offered). Requires delete permission (server mode full). Check `done` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = true,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_delete_mail(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<DeleteItemParams>,
    ) -> CallToolResult {
        self.confirmed_delete(peer, params, ActionKind::DeleteMail)
            .await
    }

    #[tool(
        description = "Move ONE calendar event to Deleted Items, only after the user approves a prompt showing its subject, time and whether it is a whole series. Refused for meetings with attendees (they would not be told) and for items already in Deleted Items. Requires delete permission (server mode full). Check `done` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = true,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_delete_event(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<DeleteItemParams>,
    ) -> CallToolResult {
        self.confirmed_delete(peer, params, ActionKind::DeleteEvent)
            .await
    }

    #[tool(
        description = "Turn ONE mail rule on or off by exact name, only after the user approves a prompt showing the rule, its current and new state and what it does (rules that delete or forward mail get an extra warning). Rules take effect immediately. Nothing changes if the client cannot show the prompt or the user declines. Requires rules permission (server mode full). Check `done` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_toggle_rule(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<ToggleRuleParams>,
    ) -> CallToolResult {
        match prepare_toggle_rule(&params) {
            Ok(target) => self.confirmed(peer, target).await,
            Err(error) => text_error(error),
        }
    }

    #[tool(
        description = "Answer a meeting invitation you received (accept, tentative or decline), only after the user approves a prompt showing the meeting, organizer, your current and new answer and whether a reply is emailed. By default the answer is emailed to the organizer (requires send permission); with send_response=false only your own calendar changes (requires events permission). Accepting or tentatively accepting replaces your calendar item with a new one and declining removes it, so use the handle only from a fresh list_events. Refused for meetings you organized, cancelled meetings and non-meetings. Nothing happens if the client cannot show the prompt. Check `done` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_respond_event(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<RespondEventParams>,
    ) -> CallToolResult {
        match prepare_respond_event(&params) {
            Ok(target) => self.confirmed(peer, target).await,
            Err(error) => text_error(error),
        }
    }

    #[tool(
        description = "Cancel ONE meeting you organized: after the user approves a prompt listing the meeting and EVERY attendee, cancellation emails are sent to them and the meeting moves to Deleted Items. Refused for plain events (use outlook_delete_event), invitations you received, already cancelled meetings and items in Deleted Items. Requires send permission (server mode full). Nothing happens if the client cannot show the prompt. Check `done` in the result; on unknown_outcome call outlook_check_action and never retry.",
        annotations(
            read_only_hint = false,
            destructive_hint = true,
            idempotent_hint = false,
            open_world_hint = true
        )
    )]
    async fn outlook_cancel_meeting(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<DeleteItemParams>,
    ) -> CallToolResult {
        match prepare_cancel_meeting(&params) {
            Ok(target) => self.confirmed(peer, target).await,
            Err(error) => text_error(error),
        }
    }

    #[tool(
        description = "Save one or all attachments of a mail into an existing folder under the user's profile, only after the user approves a prompt listing the message, every file name and size and the folder. Nothing is overwritten (a taken name gets ' (1)'). Executable and script files (.exe, .js, .ps1, .lnk, ...) are NOT saved; the prompt and summary name each refused file and say why. Requires files permission (server mode full). Nothing is saved if the client cannot show the prompt. Check `done` in the result.",
        annotations(
            read_only_hint = false,
            destructive_hint = false,
            idempotent_hint = false,
            open_world_hint = false
        )
    )]
    async fn outlook_save_attachments(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<SaveAttachmentsParams>,
    ) -> CallToolResult {
        match prepare_save_attachments(&params) {
            Ok(target) => self.confirmed(peer, target).await,
            Err(error) => text_error(error),
        }
    }

    #[tool(
        description = "After a delete, rule change or meeting answer reported unknown_outcome (or to confirm one), look for the item by its action_id: done, not_done or unknown. Read-only. Never try a delete again before checking.",
        annotations(
            read_only_hint = true,
            destructive_hint = false,
            idempotent_hint = true,
            open_world_hint = false
        )
    )]
    async fn outlook_check_action(
        &self,
        peer: Peer<RoleServer>,
        Parameters(params): Parameters<CheckActionParams>,
    ) -> CallToolResult {
        if let Err(error) = self
            .policy
            .check(Capability::Read)
            .and(checked_send_id(&params.action_id))
        {
            return text_error(error);
        }
        outcome_result(self.action_engine(peer).check(&params.action_id).await)
    }

    #[tool(description = "List the profile's color categories (name and Outlook color index).")]
    async fn outlook_list_categories(&self) -> CallToolResult {
        self.call(MailCommand::ListCategories).await
    }

    #[tool(description = "List mail rules on the default store (index, name, enabled). Read-only.")]
    async fn outlook_list_rules(&self) -> CallToolResult {
        self.call(MailCommand::ListRules).await
    }

    #[tool(
        description = "Read whether Out-of-Office is on. Read-only; status is unknown when the MAPI property is unreadable."
    )]
    async fn outlook_get_out_of_office(&self) -> CallToolResult {
        self.call(MailCommand::GetOutOfOffice).await
    }
}

const INSTRUCTIONS: &str = "Experimental Outlook Classic MCP; read-only unless the user started the server in a wider permission mode (see outlook_rust_status). Requires locally running Outlook Classic on Windows. No sends/deletes or OST parsing. Search/list results contain short session-scoped entry_id handles instead of raw COM IDs; fetch a selected message with outlook_get_mail. Responses are compact JSON. Do not claim this build has passed live Outlook testing.";

#[tool_handler]
impl ServerHandler for OutlookMcp {
    fn get_info(&self) -> rmcp::model::ServerConfig {
        rmcp::model::ServerConfig::new(
            rmcp::model::ServerCapabilities::builder()
                .enable_tools()
                .enable_resources()
                .build(),
        )
        .with_server_info(rmcp::model::Implementation::new(
            "outlook-classic-mcp-rs",
            "0.1.0",
        ))
        .with_instructions(INSTRUCTIONS.to_string())
    }

    async fn list_resources(
        &self,
        _request: Option<rmcp::model::PaginatedRequestParams>,
        _context: rmcp::service::RequestContext<RoleServer>,
    ) -> Result<rmcp::model::ListResourcesResult, rmcp::ErrorData> {
        Ok(rmcp::model::ListResourcesResult::with_all_items(
            ui::resources(),
        ))
    }

    async fn read_resource(
        &self,
        request: rmcp::model::ReadResourceRequestParams,
        _context: rmcp::service::RequestContext<RoleServer>,
    ) -> Result<rmcp::model::ReadResourceResponse, rmcp::ErrorData> {
        ui::read(&request.uri).map(Into::into).ok_or_else(|| {
            rmcp::ErrorData::resource_not_found(format!("unknown resource {}", request.uri), None)
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn status_is_explicitly_experimental() {
        let status = foundation_status();
        assert_eq!(status.stage, "read-only-experimental");
        assert!(status.mailbox_tools_available);
        assert!(!status.outlook_connected);
    }
    #[test]
    fn no_tool_accepts_a_parameter_that_could_grant_permission() {
        // Permission comes from the server environment only. If a tool ever
        // gained an argument like this, a model could try to authorize itself.
        let forbidden = [
            "mode",
            "confirm",
            "confirmed",
            "permission",
            "permissions",
            "allow",
            "force",
            "policy",
            "override",
            "approve",
            "approved",
            "authorize",
            "authorized",
        ];
        for tool in OutlookMcp::tool_router().list_all() {
            let schema = serde_json::to_value(&*tool.input_schema).unwrap();
            if let Some(props) = schema["properties"].as_object() {
                for name in props.keys() {
                    assert!(
                        !forbidden.contains(&name.to_ascii_lowercase().as_str()),
                        "{} exposes '{name}'",
                        tool.name
                    );
                }
            }
        }
    }

    #[test]
    fn delivery_follows_the_permission_mode() {
        let policy = |mode: &str| Policy::from_settings(Some(mode), None).unwrap();
        assert_eq!(delivery_for(&policy("full"), false), Ok(Delivery::Send));
        assert_eq!(delivery_for(&policy("full"), true), Ok(Delivery::SaveOnly));
        // Draft mode never sends, it saves for the user to send from Outlook.
        assert_eq!(
            delivery_for(&policy("draft"), false),
            Ok(Delivery::SaveOnly)
        );
        assert!(delivery_for(&policy("read-only"), false).is_err());
        assert!(delivery_for(&policy("read-only"), true).is_err());
        let no_send = Policy::from_settings(Some("full"), Some("send")).unwrap();
        // Send denied but drafts allowed: the message is saved, never sent.
        assert_eq!(delivery_for(&no_send, false), Ok(Delivery::SaveOnly));
    }

    #[test]
    fn default_server_is_read_only_and_reports_its_mode() {
        let server = OutlookMcp::default();
        assert_eq!(server.policy.mode(), policy::Mode::ReadOnly);
        assert_eq!(server.policy.describe(), "read-only");
    }

    #[test]
    fn registered_tools_exclude_not_yet_implemented_writes() {
        let router = OutlookMcp::tool_router();
        let tools = router.list_all();
        assert_eq!(tools.len(), 38);
        let names: Vec<_> = tools.iter().map(|tool| tool.name.as_ref()).collect();
        for expected in [
            "outlook_whoami",
            "outlook_list_folders",
            "outlook_list_mails",
            "outlook_search_mails",
            "outlook_get_mail",
            "outlook_get_thread",
            "outlook_list_contacts",
            "outlook_search_contacts",
            "outlook_get_contact",
            "outlook_resolve_name",
            "outlook_list_categories",
            "outlook_list_rules",
            "outlook_get_out_of_office",
            "outlook_list_events",
            "outlook_get_event",
            "outlook_list_tasks",
            "outlook_mark_mail",
            "outlook_set_category",
            "outlook_create_folder",
            "outlook_move_mail",
            "outlook_send_mail",
            "outlook_reply_mail",
            "outlook_forward_mail",
            "outlook_send_draft",
            "outlook_check_send",
            "outlook_create_task",
            "outlook_complete_task",
            "outlook_create_event",
            "outlook_update_event",
            "outlook_delete_mail",
            "outlook_delete_event",
            "outlook_check_action",
            "outlook_toggle_rule",
            "outlook_respond_event",
            "outlook_cancel_meeting",
            "outlook_update_meeting",
            "outlook_save_attachments",
        ] {
            assert!(names.contains(&expected));
        }
    }
}
