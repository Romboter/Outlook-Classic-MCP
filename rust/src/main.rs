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
                "outlook-mcp-rs [--help|--version|--probe-outlook|--probe-first-mail]\n\nWith no arguments, serves a local MCP stdio connection.\nCOM probe stage: MCP mailbox tools are not implemented."
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
    // only use stderr. The server currently has no mailbox operations.
    eprintln!(
        "outlook-mcp-rs: MCP diagnostics only; Outlook COM is accessible through an explicit local --probe-outlook command"
    );
    let server = OutlookMcp.serve(stdio()).await?;
    server.waiting().await?;
    Ok(())
}
