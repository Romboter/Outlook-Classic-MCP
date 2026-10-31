//! Local stdio MCP entry point. Never print logs to stdout.

use anyhow::Result;
use outlook_classic_mcp_rs::{OutlookMcp, outlook::OutlookBridge};
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
                "outlook-mcp-rs [--help|--version|--probe-outlook|--probe-first-mail]\n\nWith no arguments, serves a local MCP stdio connection.\nExperimental read-only MCP tools are enabled; no write operations. Outlook Classic must already be running."
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
    eprintln!(
        "outlook-mcp-rs: experimental read-only Outlook Classic MCP; no live COM validation yet"
    );
    let server = OutlookMcp::new().serve(stdio()).await?;
    server.waiting().await?;
    Ok(())
}
