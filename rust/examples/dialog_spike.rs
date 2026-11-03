//! Manual spike: shows one approval window with made-up test data. Usage: dialog_spike [timeout_secs]
//! Prints only the decision. Does nothing to Outlook.

use outlook_classic_mcp_rs::dialog::{Request, approve};

fn main() {
    let secs = std::env::args()
        .nth(1)
        .and_then(|s| s.parse().ok())
        .unwrap_or(20);
    let attendees: String = (1..=40)
        .map(|i| format!("test.person{i}@example.invalid\n"))
        .collect();
    let decision = approve(Request {
        title: "Outlook MCP: approval needed".into(),
        headline: "[outlook-mcp test] Cancel this meeting and notify 40 people?".into(),
        summary: "Subject: [outlook-mcp test] spike\nWhen: 2030-01-01 10:00\nThis is a dialog test; nothing will happen.".into(),
        details: format!("Attendees (40):\n{attendees}"),
        timeout: std::time::Duration::from_secs(secs),
    });
    println!("{decision:?}");
}
