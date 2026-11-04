# Implementation status and release gates

**Updated 2026-11-03 (end of session 4).** The Rust work is merged into the fork's `rust/development` branch (PRs #23 to #46, then direct commits for G11, G10, G22, G25, G07, G08, G31, G26; no more PRs). It is **not** in `main`, and no client or plugin config points at it: the Python server is unaffected. Recheck issue states and the PR list at review time.

| Area | Code state | Verification state |
| --- | --- | --- |
| Rust crate, stdio MCP, Cargo.lock, Windows CI: #2 | In `rust/development` | 31 Python tool declarations and FastMCP runtime schemas captured and checked in CI. 15 read-only Python **client** calls matched source-derived key sets on owner-reported live testing. Wrapper envelopes and conditional outputs still outstanding |
| STA COM worker, attach/reconnect: #3 | In `rust/development` | Used in the owner's live read and write checks. Reconnect and COM error paths are fake-tested only |
| Read-only mail, folder and thread tools: #4 / #7 | In `rust/development` | Live read checks on an interactive Outlook Classic profile (shapes, counts, timings only). `list_mails` `since`/`until`, ISO `received` times and SMTP addresses for Exchange senders were checked live and read-only on 2026-11-03 |
| Table backend, indexed `AdvancedSearch`: #5 | In `rust/development` | Table is the default. Indexed is opt-in; in one test it returned the same results but took 10,968 ms vs 3,033 ms legacy. Do not change the auto backend without repeatable evidence |
| Short handles: #6 | In `rust/development` | Session-scoped identity references, not search ranks. Cross-store and reconnect cases untested |
| Permission modes, confirmed writes: #8 | In `rust/development` | Server-enforced `read-only` / `draft` / `full` modes and a person's approval for every write (MCP elicitation, or a native dialog by default per `OUTLOOK_MCP_APPROVAL`). Verified with scripts standing in for the person; **no real person has answered a real elicitation prompt**. Failure and timeout paths during `Send()`, `Delete()`, rule `Save()` and `Respond()` are fake-tested only |
| 31-tool Python parity: #9 | All 31 Python names registered (40 tools in total) | Remaining differences are ranked in [PARITY-GAPS.md](PARITY-GAPS.md). `outlook_respond_event` accept/tentative/decline has never run against a real invitation from someone else |
| MCP Apps views: #10 | Five `ui://outlook/*.html` views linked from the matching tools | Calendar view rendering was recorded; rendering of the other four is not recorded here |
| Windows distributable and release: #11 | CI groundwork only | No release artifact, signing, cutover or rollback validation |

## What is and is not verified

- [Owner-supplied sanitized report](../rust/compatibility/live-evidence-2026-11-02.md) and later session notes record **shapes, counts and timings only**; no mailbox content is kept. CI alone cannot verify Outlook COM behavior.
- Not tested live: shared or delegated mailboxes and calendars, PST or archive stores, non-default stores, all-day and daylight-saving-boundary events, live reply and forward, recurring-series deletes, attendees other than the owner.
- `outlook_rust_status.outlook_connected` is a static value, not a connection check.
- Session 4 changes were checked live and read-only (folder aliases and 11 stores, body paging, folder row fields, whoami, `scope=from` SMTP (weak sample, no Exchange sender confirmed), `next_offset`, `get_event` fields) except `no_reminder` (G26), which is unit-tested only. No live test changed the mailbox.
- Timings are single observations, not speed claims.

## Recommended next sequence

1. Work down [PARITY-GAPS.md](PARITY-GAPS.md), highest daily-use gaps first (done through G11, G10, G22 (partial), G25, G07, G08, G31 (partial), G26; open: G12, G09, G13 to G17, G27 to G29, G32 to G36 — see the ranked file; G16, G17 and G29 are probably deliberate).
2. Rebuild the release exe and try the tools in a normal session; confirm whether the client shows approval prompts.
3. [Benchmark searches](runbooks/05-search-benchmark.md) and [validate conversation reads](runbooks/07-thread-validation.md) on more stores.
4. Release gates under #11 (package, trial deploy, rollback) only after the owner says so; never infer release readiness from CI.

Source of truth: [roadmap](https://github.com/Romboter/Outlook-Classic-MCP/issues/1) and the [PR list](https://github.com/Romboter/Outlook-Classic-MCP/pulls).
