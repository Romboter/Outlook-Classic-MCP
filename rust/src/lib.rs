//! Local, read-only Outlook Classic MCP tools.
//!
//! All mailbox operations execute on the dedicated Windows COM STA worker.
//! No mutating tools are registered, and no Outlook data is logged to stderr.
pub mod contracts;
pub mod outlook;

use contracts::{FolderParams, GetMailParams, ListMailsParams, MailCommand, SearchMailsParams};
use outlook::OutlookBridge;
use rmcp::{
    ServerHandler,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock},
    tool, tool_handler, tool_router,
};
use serde::Serialize;
use std::sync::Arc;

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
}

impl OutlookMcp {
    pub fn new() -> Self {
        Self {
            bridge: Arc::new(OutlookBridge::start()),
        }
    }

    async fn call(&self, cmd: MailCommand) -> CallToolResult {
        let bridge = Arc::clone(&self.bridge);
        // COM/Windows calls must NEVER run directly on the async MCP executor.
        let result = tokio::task::spawn_blocking(move || bridge.execute(cmd)).await;
        match result {
            Ok(Ok(value)) => CallToolResult::success(vec![ContentBlock::text(value.to_string())]),
            Ok(Err(error)) => CallToolResult::error(vec![ContentBlock::text(error.to_string())]),
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
        serde_json::to_string(&foundation_status())
            .expect("serialization of static fields cannot fail")
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
        description = "Read one Outlook Classic email. Pass the short entry_id handle returned by search/list (e.g. m1); legacy raw EntryIDs also work. Bounded plain-text body; HTML disabled."
    )]
    async fn outlook_get_mail(
        &self,
        Parameters(params): Parameters<GetMailParams>,
    ) -> CallToolResult {
        self.call(MailCommand::GetMail(params)).await
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
    fn registered_tools_are_read_only() {
        let router = OutlookMcp::tool_router();
        let tools = router.list_all();
        assert_eq!(tools.len(), 6);
        let names: Vec<_> = tools.iter().map(|tool| tool.name.as_ref()).collect();
        for expected in [
            "outlook_whoami",
            "outlook_list_folders",
            "outlook_list_mails",
            "outlook_search_mails",
            "outlook_get_mail",
        ] {
            assert!(names.contains(&expected));
        }
        for forbidden in ["outlook_send_mail", "outlook_delete_mail"] {
            assert!(!names.contains(&forbidden));
        }
    }
}
