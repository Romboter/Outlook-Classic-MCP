# Runbook 09 — Trial Rust MCP configuration and rollback

**Use when:** A developer explicitly elects to test the experimental Rust binary from a separate MCP client or Claude Code environment.

**Not a production deployment.** The Python implementation remains the fallback. Do **not** replace its existing `outlook` entry until real Outlook validation and user approval.

## Preconditions

1. Build with [runbook 02](02-build-test.md) and pass local/CI checks.
2. Prove live COM access using [runbook 03](03-outlook-com-validation.md).
3. Understand that registering this MCP server in an LLM client may expose permitted mailbox data to that client/model. Use approved corporate AI policy and an isolated test profile where required.
4. Choose an independent name, e.g. `outlook-rust-test`, and a local/private scope. Confirm the existing Python server's name/config before changes.

## Preferred: Inspector without persistent configuration

```powershell
npx -y @modelcontextprotocol/inspector .\rust\target\debug\outlook-mcp-rs.exe
```

Test tool discovery and read-only operations; close the Inspector when finished. No Claude Code MCP entry is changed.

## Optional: isolated Claude Code registration

The following is an **example**, not a command to run in a shared production project. Use PowerShell; replace the path with the **absolute path to your locally built executable**. Verify `claude mcp add --help` on your installed CLI.

```powershell
$exe = (Resolve-Path .\rust\target\debug\outlook-mcp-rs.exe).Path
claude mcp add --scope local outlook-rust-test -- $exe
claude mcp list
```

Do not use `--scope project` because it would create a shared `.mcp.json` that could be committed. Keep the tested server's identity distinct from the Python MCP server. Restart only the test Claude Code session if needed. Avoid enabling both tools for unattended/automatic actions; inspect which tool namespace Claude is using.

## Verify

Run `outlook_rust_status` followed by `outlook_whoami` in the test client. Use safe/test mail only. Check for a bounded metadata-only list, short handles and explicit errors for unsupported options. Run [MCP smoke test](04-mcp-smoke-test.md).

## Rollback (remove only the test registration)

```powershell
claude mcp remove --scope local outlook-rust-test
claude mcp list
```

If your version of Claude CLI uses different remove syntax, check `claude mcp remove --help` first. Do not delete the Python MCP config or remove the pre-existing `outlook` server. Close running experimental MCP clients/processes. Session handles become invalid after restart.

## Post-test

Record the commit and success/failure with [validation template](../templates/validation-report.md). Do not paste tool responses, mailbox names, or identities. A successful test is **not** a release authorization.
