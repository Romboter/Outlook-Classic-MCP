# Live test runbook: Rust Outlook MCP (start to finish)

One linear pass through everything that can be tested today. It condenses runbooks 02 to 07; use those for detail.
All tools are **read-only**. Nothing here sends, moves, edits or deletes mail. Use test mail you are happy to read on screen, and don't paste real subjects, addresses or IDs anywhere public.

Binary under test: `C:\sc\Outlook-Classic-MCP\rust\target\release\outlook-mcp-rs.exe` (branch `docs/project-operations-runbooks`, which contains the whole Rust stack through issue #7).

Set a shortcut for the session:

```powershell
cd C:\sc\Outlook-Classic-MCP
$exe = "$PWD\rust\target\release\outlook-mcp-rs.exe"
```

Record each result in the Result column or in a copy of [the validation template](../templates/validation-report.md).

## 0. Preconditions

| # | Check | Command / action | Pass |
| --- | --- | --- | --- |
| 0.1 | Outlook **Classic** (not New Outlook) is open, profile loaded, no dialogs showing | `Get-Process OUTLOOK \| select Id,SessionId` shows one process | |
| 0.2 | Same Windows user and session as your terminal | SessionId matches `(Get-Process -Id $PID).SessionId` | |
| 0.3 | Node 22.19+ for MCP Inspector | `node --version` | |
| 0.4 | You know a test conversation: 3+ messages, replies, quoted text, in your Inbox | choose one and note its distinctive subject word | |

## 1. Build and unit tests (no Outlook needed)

```powershell
cargo test --locked --manifest-path rust/Cargo.toml
cargo build --release --locked --manifest-path rust/Cargo.toml
& $exe --version
& $exe --help
```

**Expect:** 24 tests pass, build succeeds, `--help` prints the usage line listing `--probe-outlook` and `--probe-first-mail`.
Optional lint gate: `cargo fmt --manifest-path rust/Cargo.toml --all -- --check` and `cargo clippy --locked --manifest-path rust/Cargo.toml --all-targets -- -D warnings`.

## 2. Protocol check (no Outlook needed)

```powershell
npx -y @modelcontextprotocol/inspector --cli $exe --method tools/list
npx -y @modelcontextprotocol/inspector --cli $exe --method tools/call --tool-name outlook_rust_status
```

**Expect:** exactly 7 tools: `outlook_rust_status`, `outlook_whoami`, `outlook_list_folders`, `outlook_list_mails`, `outlook_search_mails`, `outlook_get_mail`, `outlook_get_thread`, each with an input schema.
`outlook_rust_status` is a static answer; its `outlook_connected:false` says nothing about COM.

## 3. COM attach (needs Outlook)

```powershell
& $exe --probe-outlook
```

**Expect:** JSON with your profile name, Inbox name, plausible `inbox_items` and `inbox_subfolders`. It must not launch Outlook.
Skip `--probe-first-mail` unless the Inbox's newest subject is fine to see on screen.

| # | Check | Pass |
| --- | --- | --- |
| 3.1 | Probe returns the right profile and plausible counts | |
| 3.2 | With Outlook **closed**, probe fails with a clear "not running" error and does not start Outlook | |
| 3.3 | Reopen Outlook, probe works again | |

## 4. Tool calls in the Inspector UI

```powershell
npx -y @modelcontextprotocol/inspector $exe
```

Connect over stdio, then run these in order. Substitute the real handle each time; handles look like `m1`, `m2` and only live for this server process.

| # | Tool | Arguments | Expect | Pass |
| --- | --- | --- | --- | --- |
| 4.1 | `outlook_whoami` | none | your display name | |
| 4.2 | `outlook_list_folders` | `{"max_depth":1}` | your real folders, 100 entries max | |
| 4.3 | `outlook_list_mails` | `{"folder":"Inbox","limit":5,"backend":"legacy"}` | 5 newest, metadata only, no bodies, short `entry_id` like `m1` | |
| 4.4 | `outlook_list_mails` | same, `"backend":"table"` | same messages in the same order as 4.3, `backend` reports `table` | |
| 4.5 | `outlook_search_mails` | `{"query":"<word>","scope":"subject","limit":5,"backend":"legacy"}` | hits that really contain the word | |
| 4.6 | `outlook_search_mails` | same, `"backend":"table"` | same set as 4.5 | |
| 4.7 | `outlook_search_mails` | `{"query":"<word>","scope":"from","limit":5}` | sender matches | |
| 4.8 | `outlook_get_mail` | `{"entry_id":"<handle from 4.5>","max_body_chars":500}` | the right message, body at most 500 chars, no long EntryID/StoreID in the output | |
| 4.9 | `outlook_get_mail` | same without `max_body_chars` | default 4,000 char cap | |

Note the `elapsed_ms` and `backend` fields on every list and search response.

## 5. Negative paths (each must fail cleanly, not crash or hang)

| # | Call | Expect | Pass |
| --- | --- | --- | --- |
| 5.1 | `outlook_list_mails` `{"limit":31}` | validation error | |
| 5.2 | `outlook_get_mail` `{"entry_id":"<handle>","include_html":true}` | explicit "not supported" error | |
| 5.3 | `outlook_search_mails` `{"query":"x","backend":"table","scope":"subject_body"}` | rejects the combination | |
| 5.4 | `outlook_search_mails` with `since` or `until` set | "unsupported" error, filter not silently ignored | |
| 5.5 | `outlook_get_mail` `{"entry_id":"m9999"}` | clear error telling you to search again | |
| 5.6 | Disconnect Inspector, reconnect, then call `outlook_get_mail` with an old handle | clear stale-handle error; a fresh search gives new handles | |
| 5.7 | Close Outlook, call `outlook_whoami`, reopen Outlook, call again | error while closed, recovers after reopen (the server may need restarting; note which) | |

## 6. Conversation retrieval (`outlook_get_thread`)

Use your test conversation from 0.4. Find a message in it with `outlook_search_mails`, then:

| # | Arguments | Expect | Pass |
| --- | --- | --- | --- |
| 6.1 | `{"entry_id":"<h>","max_messages":10,"include_bodies":true,"trim_quoted":true,"per_message_chars":1500,"total_body_chars":8000}` | `association` is `verified_conversation_id`; every message belongs to that conversation; oldest to newest; count matches Outlook | |
| 6.2 | same with `"trim_quoted":false` | longer bodies; quoted history is back | |
| 6.3 | Compare 6.1 against each original with `outlook_get_mail` | no newly written text lost by trimming | |
| 6.4 | `{"entry_id":"<h>","include_bodies":false}` | metadata only | |
| 6.5 | `{"entry_id":"<h>","max_messages":3,"per_message_chars":250,"total_body_chars":500}` | no body over 250, total at most 500, `older_messages_omitted` true if the thread is longer | |
| 6.6 | A message with non-English or emoji text, with low char budgets | budgets hold, no broken characters | |
| 6.7 | A single-message conversation | one message, no crash | |
| 6.8 | A different thread with the same subject | its messages do not appear | |
| 6.9 | Thread spanning Inbox and Sent Items | both folders' messages present | |

Check these flags are truthful: `older_messages_omitted`, `scan_truncated`, `inaccessible_messages`, `missing_store_identity`, `identity_mismatch`, `anchor_included`, `quote_history_omitted`, `body_truncated`.
Known limits: Deleted Items are not included, and the scan stops at 350 rows.

## 7. Search speed and correctness (optional, after 4 to 6 pass)

Same query, same folder, same limit; compare backends only when they return the same hits.

```json
{"query":"<word>","folder":"Inbox","scope":"subject","limit":10,"backend":"legacy"}
{"query":"<word>","folder":"Inbox","scope":"subject","limit":10,"backend":"table"}
{"query":"<word>","folder":"Inbox","scope":"subject","limit":10,"backend":"indexed","search_subfolders":false}
```

For each backend: 1 cold run, then 10 warm runs, noting `elapsed_ms`. Write down median and worst, and whether the hit set matched `legacy`.

| # | Check | Pass |
| --- | --- | --- |
| 7.1 | `table` hits equal `legacy` hits | |
| 7.2 | `indexed` returns, or errors clearly (Instant Search off); it never silently falls back | |
| 7.3 | `indexed` hits equal `legacy` hits (it uses phrase matching, so small differences are possible, not necessarily wrong) | |
| 7.4 | Large folder (thousands of items): all three backends still finish inside 45 s | |
| 7.5 | `backend:"auto"` labels the backend it actually used | |

`indexed` is the least proven piece (event sink, 18 s wait). If it hangs or errors, record it, then use `legacy` or `auto`. Don't retry repeatedly.

## 8. Try it from Claude Code (only after 4 to 6 pass)

Use a **separate server name** so the working Python MCP is untouched:

```powershell
claude mcp add outlook-rs --scope user -- "C:\sc\Outlook-Classic-MCP\rust\target\release\outlook-mcp-rs.exe"
```

In a new Claude Code session ask it to: list your 5 newest Inbox mails, search for the test word, read one message, and summarise the test thread. Confirm it uses the `outlook_*` tools from `outlook-rs`, handles are short, and token use is small.
Remove when done:

```powershell
claude mcp remove outlook-rs --scope user
```

## 9. Wrap-up

| Question | Answer |
| --- | --- |
| Did anything write to or change your mailbox? (should be no) | |
| Any hang over 45 s or an Outlook dialog left blocking? | |
| Which sections failed, with the exact error text (redacted)? | |
| Outlook build (File, Office Account, About Outlook) and store type (Exchange cached, PST) | |

## Not covered (not built yet)

Write operations and permission modes (issue #8), the other roughly 24 Python tools (#9), MCP Apps UI (#10), packaging and release (#11). Don't test or expect these.

## If something fails

Run [troubleshooting](06-troubleshooting.md). Quick rules: Outlook must be running, same user and session, no modal dialog open; a timeout does not cancel the COM call, so wait before retrying; restart the MCP process to reset handles.
