# Runbook 06 — Troubleshoot Outlook Classic COM, search and MCP failures

**Use when:** A Rust MCP tool returns an error, an indexed search times out, a handle expires, or the Inspector cannot connect.

## Triage order

1. **Stop and categorize:** Is this tool startup, MCP protocol, COM attach, provider search, message identity, or timeout? Do not reissue the same expensive request repeatedly.
2. **Verify branch and binary:**
   ```powershell
   git status --short --branch
   .\rust\target\debug\outlook-mcp-rs.exe --version
   .\rust\target\debug\outlook-mcp-rs.exe --help
   ```
3. **Verify Outlook Classic in the same interactive session:**
   ```powershell
   Get-Process OUTLOOK -ErrorAction SilentlyContinue |
     Select-Object Id,SessionId
   ```
4. **Try only the safe local probe** and inspect output privately:
   ```powershell
   .\rust\target\debug\outlook-mcp-rs.exe --probe-outlook
   ```
5. Reproduce the smallest failing read with Inspector, starting with `outlook_whoami` and `outlook_list_folders`. For search, force `backend:legacy` before testing `table` or `indexed`.
6. Record **error category and sanitized reproduction**, not full mailbox response or identifiers. Stop if Outlook prompts for credentials/permissions or appears to be targeting the wrong profile.

## Symptom / action matrix

| Symptom | Likely cause | Safe action |
| --- | --- | --- |
| Executable not found | Unbuilt binary or wrong working directory | Rebuild using [runbook 02](02-build-test.md); use explicit repo-root path |
| MCP initialize fails / hangs | Wrong server path or stdout pollution | Run `--help` separately; test Inspector `initialize`; keep stdout JSON-RPC only |
| Outlook not running / COM ROT missing | New Outlook, different account/session, profile not loaded | Start Outlook Classic manually in current session; re-run `--probe-outlook` |
| Wrong Outlook user/profile | COM attached to unexpected profile | **Stop**; verify account/session. Do not inspect email |
| RPC unavailable / COM disconnected | Outlook restarted/crashed or COM server busy | Restart Rust MCP client; verify Outlook state, then redo basic probe |
| 45-second request timeout | STA blocked by modal dialog or slow provider | Check Outlook UI; note that timeout **does not cancel COM**; avoid parallel retries |
| `backend:indexed` not available | `IsInstantSearchEnabled` not confirmed | Use `legacy`/auto; record store type; don't claim indexed timing |
| Indexed event wait 18 seconds | Event sink, index or query/provider issue | Stop or allow best-effort Stop; use legacy; record sanitized failure |
| Table query unsupported | Store rejects row properties/filter | Try `backend:legacy`; report fallback flag and error class |
| Unknown `mN` handle | Session restart, cache eviction or reconnect | Search/list again and use the new handle |
| Thread not available | Store conversation support disabled, deleted/inaccessible item, StoreID conversion issue | Fetch original via `outlook_get_mail` where possible; record omissions, do not join by subject |
| No results / missing older mail | Cache window, Outlook index, folder scope or query semantics | Verify directly in Outlook UI and compare same semantic filters |

## Escalation evidence

Fill out [validation template](../templates/validation-report.md): OS/Outlook build, Rust commit, tool name, error category, backend, sanitized counts, recurrence and safe resolution steps. **Do not include raw COM object dumps or actual mail content.**

## Recovery / rollback

- Restart only the test MCP process first; leave working Python server untouched.
- Switch tested queries back to `backend:legacy` or stop using Rust if correctness is uncertain.
- If test Claude Code configuration was added, remove only the test-specific server entry as in [runbook 09](09-test-deployment-rollback.md).
- Do **not** alter production Outlook profile, reset OST/index, disable security features or change permissions as a routine first response.

Where a defect is reproducible, create an issue/PR against **your fork**, not upstream. Keep sensitive logs local.
