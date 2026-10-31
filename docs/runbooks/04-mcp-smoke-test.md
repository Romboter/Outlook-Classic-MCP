# Runbook 04 — MCP protocol and read-only tool smoke test

**Use when:** Confirming that the experimental Rust executable speaks MCP over stdio and its tool definitions are usable. Use **MCP Inspector** before modifying an everyday Claude Code configuration.

## Prerequisites

- Built Rust debug executable from [runbook 02](02-build-test.md).
- Node.js **22.19+** and `npx` for recent MCP Inspector.
- Only for real mailbox calls: Outlook Classic running in the same interactive Windows session. For tool discovery and `outlook_rust_status`, Outlook is not required.

## Protocol checks (PowerShell from repo root)

```powershell
npx -y @modelcontextprotocol/inspector --cli .\rust\target\debug\outlook-mcp-rs.exe --method initialize
npx -y @modelcontextprotocol/inspector --cli .\rust\target\debug\outlook-mcp-rs.exe --method tools/list --format json
```

Expected: initialized local MCP server; seven registered tools (one `outlook_rust_status` diagnostic + six read-only Outlook tools: whoami, list_folders, list_mails, search_mails, get_mail, get_thread). Tool inputs should have JSON schemas.

Call the static diagnostic:

```powershell
npx -y @modelcontextprotocol/inspector --cli .\rust\target\debug\outlook-mcp-rs.exe --method tools/call --tool-name outlook_rust_status --format json
```

`outlook_connected:false` in this output is a static flag; **it does not test Outlook connectivity**.

## Interactive mailbox checks

Launch the Inspector UI and select the Rust executable as a **stdio** server:

```powershell
npx -y @modelcontextprotocol/inspector .\rust\target\debug\outlook-mcp-rs.exe
```

Once Outlook Classic is open and the Inspector is connected, invoke:

1. `outlook_whoami` — validates actual COM.
2. `outlook_list_folders` with `{"max_depth":1}` — ensure plausible folders.
3. `outlook_list_mails` with `{"folder":"Inbox","limit":5,"backend":"legacy"}`.
4. `outlook_search_mails` with an approved test word, `{"query":"test","scope":"subject","limit":5,"backend":"legacy"}`.
5. If a matching test message exists, use the **returned** handle in `outlook_get_mail` with `max_body_chars:500`. Don't assume `m1` always exists.
6. If that message belongs to a conversation, try `outlook_get_thread` with `max_messages:3`, `per_message_chars:250`, `total_body_chars:500`.

Check body budgets, short handles, error flags, false-positive tool claims and no writes. Sample query may return zero hits; that's not a server error.

## Negative-path checks

- An invalid `limit:31` should return an error.
- `outlook_get_mail` with `include_html:true` should return an explicit error.
- Search with `backend:"table"` and `scope:"subject_body"` should reject the unsupported combination.
- A stale handle from a prior server session should error and request a fresh search.

## Troubleshooting and cleanup

- Stdio output polluted: capture stderr separately; stdout belongs to the protocol, so do not run standalone `--help` as the server in Inspector.
- Inspector cannot start: verify Node/npx, executable path, version and antivirus execution policy.
- COM tool fails: follow [runbook 03](03-outlook-com-validation.md) then [runbook 06](06-troubleshooting.md).
- Stop the Inspector/client at the end. No permanent production MCP configuration should be changed.

Inspector CLI syntax reference: https://modelcontextprotocol.io/docs/tools/inspector (and https://github.com/modelcontextprotocol/inspector). **Inspector version/options may evolve; verify `--help` if an argument is rejected.**
