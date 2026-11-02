# Implementation status and release gates

**Status updated 2026-11-02 for Python contract evidence only.** Recheck issue/PR heads and results at review time. This branch is stacked on the existing fork-only read-only prototype and remains unmerged.

| Area | Code state | Verification state |
| --- | --- | --- |
| Rust crate, stdio MCP, Cargo.lock, Windows CI — #2 | Draft PR [#12](https://github.com/Romboter/Outlook-Classic-MCP/pull/12) | 31 machine-readable Python declaration and FastMCP runtime schemas recorded in draft PR #22; 15 read-only Python **client** calls matched source-derived key sets on owner-reported real Outlook testing. Wrappers, conditional outputs and full parity still outstanding |
| STA COM worker, attach/reconnect, probes — #3 | Draft PR [#13](https://github.com/Romboter/Outlook-Classic-MCP/pull/13) | Builds; **not live Outlook tested** |
| Six read-only mail/folder/thread tools — #4 / #7 | Draft PRs [#14](https://github.com/Romboter/Outlook-Classic-MCP/pull/14), [#18](https://github.com/Romboter/Outlook-Classic-MCP/pull/18) | Windows CI passed; actual mailbox behavior unverified |
| Table backend — #5 | Draft PR [#15](https://github.com/Romboter/Outlook-Classic-MCP/pull/15) | Windows CI passed; no measured real-mailbox benefit |
| Short handles — #6 | Draft PR [#16](https://github.com/Romboter/Outlook-Classic-MCP/pull/16) | Owner observed different handles for the same mail across search backends/thread. Investigate in one persistent session before claiming stable aliases; cross-store and reconnect cases remain outstanding |
| Indexed `AdvancedSearch` events — #5 | Draft PR [#17](https://github.com/Romboter/Outlook-Classic-MCP/pull/17) | Owner reported opt-in indexed search returned the same three matching messages but took 10,968 ms vs. 3,033 ms legacy in one test. Keep opt-in; broader benchmark outstanding |
| Permission and write-operation safety — #8 | Partial read-only groundwork | Permission modes and dangerous-operation protections not implemented |
| Full 31-tool Python parity — #9 | Not implemented | Not validated |
| MCP Apps UI parity — #10 | Not implemented | Not validated |
| Windows distributable/release — #11 | CI groundwork only | No release artifact or rollback validation |

## 2026-11-02 live evidence: what is and is not verified

- [Owner-supplied sanitized report](../rust/compatibility/live-evidence-2026-11-02.md): all **15 read-only Python client functions** returned without error when dispatched through the live OutlookBridge STA. Their observed key sets matched the source-derived inventory; reported types for nested accounts, calendar attendees, category color and OOF success result are recorded in `python-results.redacted.json`.
- **Not exercised:** Python FastMCP tool wrappers (`response_format`, `ui_result`, `list_folders` envelope), attachment elements, GAL rows, unresolved name branch, OOF error branch and all mutating tools.
- Rust parity gaps: `outlook_whoami` returns only current user; list/search/get mail omit several Python fields and use a different received-date representation. A legacy Exchange sender address can be empty when indexed search supplies SMTP; session handles differed in the reported tests; investigate without exposing raw IDs.
- `outlook_rust_status.outlook_connected` is a static false value, **not** an actual connection check. The Rust thread tool was exercised on a one-message thread but is not part of Python tool parity.
- Timings are **single observations**, not speed claims. In one comparable Rust search, indexed took approximately 3.6× legacy; do not change the auto backend without repeatable correctness and performance evidence.
- CI-only schema validation and compile/test results cannot verify Outlook COM behavior. No release or merge authorization is implied.

The existing Python MCP implementation is unaffected. **All listed Rust PRs are drafts in a stacked feature chain, not merged into fork `main`.**

## Recommended next sequence

1. [Run Windows environment and COM validation](runbooks/03-outlook-com-validation.md) with a real Outlook Classic profile.
2. [Use MCP Inspector](runbooks/04-mcp-smoke-test.md) to test tool discovery and read calls without replacing working client config.
3. [Benchmark searches](runbooks/05-search-benchmark.md), including forced `indexed` on a store with a valid index; verify result completeness, event delivery, timeouts and correctness.
4. [Test conversation reads](runbooks/07-thread-validation.md) against raw originals and multi-store examples.
5. Resolve missing schema/permission/security requirements in issues #2/#8 before any write tools.
6. After user approval and release gates, package and trial-deploy; never infer release readiness from CI alone.

Source of truth: [roadmap](https://github.com/Romboter/Outlook-Classic-MCP/issues/1), [issue #7](https://github.com/Romboter/Outlook-Classic-MCP/issues/7) and the [PR list](https://github.com/Romboter/Outlook-Classic-MCP/pulls).
