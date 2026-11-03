//! Local stdio MCP entry point. Never print logs to stdout.

use anyhow::Result;
use outlook_classic_mcp_rs::{OutlookMcp, outlook::OutlookBridge, policy::Policy};
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
                "outlook-mcp-rs [--help|--version|--probe-outlook|--probe-first-mail]\n\nWith no arguments, serves a local MCP stdio connection.\nExperimental read-only MCP tools are enabled; no write operations exist yet. Outlook Classic must already be running.\n\nPermission mode comes only from the server environment and is read-only by default:\n  OUTLOOK_MCP_MODE = read-only | draft | full\n  OUTLOOK_MCP_DENY = comma list from modify,draft,send,delete,rules,events,files\nUnrecognized values stop the server."
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
    let server = OutlookMcp::with_policy(policy).serve(stdio()).await?;
    server.waiting().await?;
    Ok(())
}
