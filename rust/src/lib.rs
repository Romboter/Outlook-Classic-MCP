//! Foundation of the Rust Outlook Classic MCP server.
//!
//! This stage intentionally exposes **no mailbox tools**. It does not attempt
//! to start Outlook, read email, or access a COM session. See issues #3 and #4.

use rmcp::{ServerHandler, tool, tool_handler, tool_router};
use serde::Serialize;

/// Informational status only. Do not confuse with a live Outlook connection.
#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub struct FoundationStatus {
    pub implementation: &'static str,
    pub version: &'static str,
    pub stage: &'static str,
    pub outlook_connected: bool,
    pub mailbox_tools_available: bool,
}

pub fn foundation_status() -> FoundationStatus {
    FoundationStatus {
        implementation: "rust",
        version: env!("CARGO_PKG_VERSION"),
        stage: "foundation-only",
        outlook_connected: false,
        mailbox_tools_available: false,
    }
}

#[derive(Clone)]
pub struct OutlookMcp;

#[tool_router]
impl OutlookMcp {
    /// Verify this is the Rust scaffold, not a connected Outlook session.
    #[tool(
        description = "Return Rust scaffold status. This version does not connect to Outlook or expose mailbox data."
    )]
    fn outlook_rust_status(&self) -> String {
        serde_json::to_string(&foundation_status())
            .expect("serialization of a static status struct cannot fail")
    }
}

#[tool_handler(
    name = "outlook-classic-mcp-rs",
    version = "0.1.0",
    instructions = "Foundation build only. No Outlook COM connection or mailbox tools are implemented. Continue using the Python server for real mailbox operations."
)]
impl ServerHandler for OutlookMcp {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scaffold_does_not_claim_outlook_connection() {
        let status = foundation_status();
        assert!(!status.outlook_connected);
        assert!(!status.mailbox_tools_available);
        assert_eq!(status.stage, "foundation-only");
    }

    #[test]
    fn scaffold_status_is_compact_and_valid_json() {
        let result = serde_json::to_string(&foundation_status()).unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&result).unwrap();
        assert_eq!(parsed["implementation"], "rust");
        assert_eq!(parsed["outlook_connected"], false);
        assert!(result.len() < 256);
    }
}
