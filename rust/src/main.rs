//! Local stdio MCP entry point. Never print logs to stdout.

use anyhow::Result;
use outlook_classic_mcp_rs::{
    OutlookMcp, action::ActionSettings, approval::ApprovalMode, outlook::OutlookBridge,
    policy::Policy, send::SendSettings,
};
use rmcp::{ServiceExt, transport::stdio};

#[tokio::main]
async fn main() -> Result<()> {
    let mut args = std::env::args().skip(1);
    match (args.next(), args.next()) {
        (Some(flag), None) if flag == "--version" => {
            println!("outlook-mcp-rs {}", env!("CARGO_PKG_VERSION"));
            return Ok(());
        }
        (Some(flag), None) if flag == "--help" || flag == "-h" => {
            println!(
                "outlook-mcp-rs [--help|--version|--probe-outlook|--probe-first-mail]\n\nWith no arguments, serves a local MCP stdio connection.\nExperimental read-only MCP tools are enabled; write tools exist only in wider permission modes. Outlook Classic must already be running.\n\nPermission mode comes only from the server environment and is read-only by default:\n  OUTLOOK_MCP_MODE = read-only | draft | full\n  OUTLOOK_MCP_DENY = comma list from modify,draft,send,delete,rules,events,files\nSend limits (prompts, not walls): OUTLOOK_MCP_SEND_MAX_RECIPIENTS (20), OUTLOOK_MCP_SEND_PER_HOUR (5), OUTLOOK_MCP_SEND_ALLOW_DOMAINS, OUTLOOK_MCP_SEND_STRICT=1 (refuse instead of prompt); hard ceilings OUTLOOK_MCP_SEND_HARD_MAX_RECIPIENTS (200) and OUTLOOK_MCP_SEND_HARD_MAX_PER_HOUR (25).\nDelete limits (prompts, not walls): OUTLOOK_MCP_DELETE_PER_HOUR (10), OUTLOOK_MCP_DELETE_HARD_MAX_PER_HOUR (50), OUTLOOK_MCP_RULE_PER_HOUR (5), OUTLOOK_MCP_RULE_HARD_MAX_PER_HOUR (20), OUTLOOK_MCP_RESPONSE_PER_HOUR (10), OUTLOOK_MCP_RESPONSE_HARD_MAX_PER_HOUR (50), OUTLOOK_MCP_MEETING_PER_HOUR (10), OUTLOOK_MCP_MEETING_HARD_MAX_PER_HOUR (50), OUTLOOK_MCP_FILES_PER_HOUR (10), OUTLOOK_MCP_FILES_HARD_MAX_PER_HOUR (50), OUTLOOK_MCP_ACTION_STRICT=1 (refuse instead of prompt).
Approval: OUTLOOK_MCP_APPROVAL = dialog (default, native window) | auto (client prompt if offered, else window) | elicitation (client prompt only) | none (no approval possible: sends save drafts, other actions do nothing). Every write still needs a person to approve it.
Unrecognized values stop the server."
            );
            return Ok(());
        }
        (Some(flag), None) if flag == "--probe-outlook" || flag == "--probe-first-mail" => {
            let bridge = OutlookBridge::start();
            let probe = bridge.probe(flag == "--probe-first-mail")?;
            println!("{}", serde_json::to_string(&probe)?);
            return Ok(());
        }
        (None, None) => {}
        _ => {
            anyhow::bail!("unknown argument; run with --help");
        }
    }

    // On stdio transport, stdout is reserved for JSON-RPC. Diagnostics must
    // only use stderr. Outlook calls are dispatched to the dedicated STA.
    let policy = Policy::from_env().map_err(anyhow::Error::msg)?;
    eprintln!(
        "outlook-mcp-rs: experimental Outlook Classic MCP, permission mode: {}; no live COM validation yet",
        policy.describe()
    );
    let send_settings = SendSettings::from_env().map_err(anyhow::Error::msg)?;
    let action_settings = ActionSettings::from_env().map_err(anyhow::Error::msg)?;
    let approval = ApprovalMode::from_env().map_err(anyhow::Error::msg)?;
    let server = OutlookMcp::with_settings(policy, send_settings)
        .with_approval(approval)
        .with_action_settings(action_settings)
        .serve(stdio())
        .await?;
    server.waiting().await?;
    Ok(())
}
