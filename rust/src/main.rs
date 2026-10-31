//! Local stdio MCP entry point. Never print logs to stdout.

use anyhow::Result;
use outlook_classic_mcp_rs::OutlookMcp;
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
                "outlook-mcp-rs [--help|--version]\n\nWith no arguments, serves a local MCP stdio connection.\nFoundation stage: Outlook mailbox tools are not implemented."
            );
            return Ok(());
        }
        (None, None) => {}
        _ => {
            anyhow::bail!("unknown argument; run with --help");
        }
    }

    // On stdio transport, stdout is reserved for JSON-RPC. Diagnostics must
    // only use stderr. The server currently has no mailbox operations.
    eprintln!("outlook-mcp-rs: foundation-only MCP server (no Outlook COM yet)");
    let server = OutlookMcp.serve(stdio()).await?;
    server.waiting().await?;
    Ok(())
}
