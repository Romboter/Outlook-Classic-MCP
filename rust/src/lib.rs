//! Local, read-only Outlook Classic MCP tools.
//!
//! All mailbox operations execute on the dedicated Windows COM STA worker.
//! No mutating tools are registered, and no Outlook data is logged to stderr.
pub mod contracts;
pub mod localtime;
pub mod outlook;
pub mod policy;
pub mod safe_path;
pub mod thread_text;
pub mod ticket;

use contracts::{
    FolderParams, GetContactParams, GetEventParams, GetMailParams, ListContactsParams,
    ListEventsParams, ListMailsParams, ListTasksParams, MailCommand, ResolveNameParams,
    SearchContactsParams, SearchMailsParams, ThreadParams,
};
use outlook::OutlookBridge;
use policy::Policy;
use rmcp::{
    ServerHandler,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock},
    tool, tool_handler, tool_router,
};
use serde::Serialize;
use std::sync::Arc;
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
}

impl OutlookMcp {
    /// Read-only, the safest mode. Use `with_policy` to opt into more.
    pub fn new() -> Self {
        Self::with_policy(Policy::read_only())
    }

    pub fn with_policy(policy: Policy) -> Self {
        Self {
            bridge: Arc::new(OutlookBridge::start()),
            policy,
        }
    }

    async fn call(&self, cmd: MailCommand) -> CallToolResult {
        // Server-side gate: nothing the model sends can change the policy.
        if let Err(message) = self.policy.check(cmd.capability()) {
            return CallToolResult::error(vec![ContentBlock::text(message)]);
        }
        let bridge = Arc::clone(&self.bridge);
        // COM/Windows calls must NEVER run directly on the async MCP executor.
        let result = tokio::task::spawn_blocking(move || bridge.execute(cmd)).await;
        match result {
            Ok(Ok(value)) => CallToolResult::success(vec![ContentBlock::text(value.to_string())]),
            Ok(Err(error)) => {
                let text = error.to_string();
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
            Err(error) => CallToolResult::error(vec![ContentBlock::text(format!(
                "Outlook worker task failed: {error}"
            ))]),
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
    fn outlook_rust_status(&self) -> String {
        let mut status = serde_json::to_value(foundation_status())
            .expect("serialization of static fields cannot fail");
        status["permission_mode"] = serde_json::json!(self.policy.describe());
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
        description = "List up to 30 metadata-only mail summaries in one folder. EntryID fields are short session-scoped m1-style handles; call get_mail with a returned handle. Backends auto/table/legacy."
    )]
    async fn outlook_list_mails(
        &self,
        Parameters(params): Parameters<ListMailsParams>,
    ) -> CallToolResult {
        self.call(MailCommand::ListMails(params)).await
    }

    #[tool(
        description = "Search one Outlook Classic folder and return compact metadata-only hits with session-scoped m1-style handles instead of long EntryIDs/StoreIDs. Backends auto/table/legacy/indexed (indexed is experimental, requires Instant Search and waits for completion event); elapsed_ms is measured."
    )]
    async fn outlook_search_mails(
        &self,
        Parameters(params): Parameters<SearchMailsParams>,
    ) -> CallToolResult {
        self.call(MailCommand::SearchMails(params)).await
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
        description = "Read one Outlook Classic email. Pass the short entry_id handle returned by search/list (e.g. m1); legacy raw EntryIDs also work. Bounded plain-text body; HTML disabled."
    )]
    async fn outlook_get_mail(
        &self,
        Parameters(params): Parameters<GetMailParams>,
    ) -> CallToolResult {
        self.call(MailCommand::GetMail(params)).await
    }

    #[tool(
        description = "List contacts from every contact folder in every store, sorted by name. Bounded (default 50, max 200; scan_truncated reports hitting the scan budget). entry_id is a short session handle; call outlook_get_contact."
    )]
    async fn outlook_list_contacts(
        &self,
        Parameters(params): Parameters<ListContactsParams>,
    ) -> CallToolResult {
        self.call(MailCommand::ListContacts(params)).await
    }

    #[tool(
        description = "Search contact folders and, by default, the Exchange Global Address List. All query words must match name, email, company or job title. Bounded scans; scan_truncated reports a partial search."
    )]
    async fn outlook_search_contacts(
        &self,
        Parameters(params): Parameters<SearchContactsParams>,
    ) -> CallToolResult {
        self.call(MailCommand::SearchContacts(params)).await
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
        description = "List calendar events in a bounded local-time window (default now to +14 days; limit default 50, max 200). Recurring series are expanded by default. Pass local ISO-8601 times or times with a UTC offset (converted to local); results carry the local offset. entry_id is a short session handle."
    )]
    async fn outlook_list_events(
        &self,
        Parameters(params): Parameters<ListEventsParams>,
    ) -> CallToolResult {
        self.call(MailCommand::ListEvents(params)).await
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
        description = "List tasks from the default Tasks folder sorted by due date (limit default 50, max 200). Completed tasks are excluded unless include_completed is true. Tasks without a date report null."
    )]
    async fn outlook_list_tasks(
        &self,
        Parameters(params): Parameters<ListTasksParams>,
    ) -> CallToolResult {
        self.call(MailCommand::ListTasks(params)).await
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

#[tool_handler(
    name = "outlook-classic-mcp-rs",
    version = "0.1.0",
    instructions = "Experimental read-only Outlook Classic MCP. Requires locally running Outlook Classic on Windows. No sends/deletes or OST parsing. Search/list results contain short session-scoped entry_id handles instead of raw COM IDs; fetch a selected message with outlook_get_mail. Responses are compact JSON. Do not claim this build has passed live Outlook testing."
)]
impl ServerHandler for OutlookMcp {}

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
    fn default_server_is_read_only_and_reports_its_mode() {
        let server = OutlookMcp::default();
        assert_eq!(server.policy.mode(), policy::Mode::ReadOnly);
        let status: serde_json::Value =
            serde_json::from_str(&server.outlook_rust_status()).unwrap();
        assert_eq!(status["permission_mode"], "read-only");
    }

    #[test]
    fn registered_tools_are_read_only() {
        let router = OutlookMcp::tool_router();
        let tools = router.list_all();
        assert_eq!(tools.len(), 17);
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
        ] {
            assert!(names.contains(&expected));
        }
        for forbidden in ["outlook_send_mail", "outlook_delete_mail"] {
            assert!(!names.contains(&forbidden));
        }
    }
}
