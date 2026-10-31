# Implementation status and release gates

**Snapshot date: 2026-10-31.** Recheck issue/PR states at the time of use.

| Area | Code state | Verification state |
| --- | --- | --- |
| Rust crate, stdio MCP, Cargo.lock, Windows CI — #2 | Draft PR [#12](https://github.com/Romboter/Outlook-Classic-MCP/pull/12) | Windows CI passed; complete Python schema snapshot outstanding |
| STA COM worker, attach/reconnect, probes — #3 | Draft PR [#13](https://github.com/Romboter/Outlook-Classic-MCP/pull/13) | Builds; **not live Outlook tested** |
| Six read-only mail/folder/thread tools — #4 / #7 | Draft PRs [#14](https://github.com/Romboter/Outlook-Classic-MCP/pull/14), [#18](https://github.com/Romboter/Outlook-Classic-MCP/pull/18) | Windows CI passed; actual mailbox behavior unverified |
| Table backend — #5 | Draft PR [#15](https://github.com/Romboter/Outlook-Classic-MCP/pull/15) | Windows CI passed; no measured real-mailbox benefit |
| Short handles — #6 | Draft PR [#16](https://github.com/Romboter/Outlook-Classic-MCP/pull/16) | Synthetic tests passed; real store/reconnect tests outstanding |
| Indexed `AdvancedSearch` events — #5 | Draft PR [#17](https://github.com/Romboter/Outlook-Classic-MCP/pull/17) | CI passed; opt-in and **unverified live** |
| Permission and write-operation safety — #8 | Partial read-only groundwork | Permission modes and dangerous-operation protections not implemented |
| Full 31-tool Python parity — #9 | Not implemented | Not validated |
| MCP Apps UI parity — #10 | Not implemented | Not validated |
| Windows distributable/release — #11 | CI groundwork only | No release artifact or rollback validation |

The existing Python MCP implementation is unaffected. **All listed Rust PRs are drafts in a stacked feature chain, not merged into fork `main`.**

## Recommended next sequence

1. [Run Windows environment and COM validation](runbooks/03-outlook-com-validation.md) with a real Outlook Classic profile.
2. [Use MCP Inspector](runbooks/04-mcp-smoke-test.md) to test tool discovery and read calls without replacing working client config.
3. [Benchmark searches](runbooks/05-search-benchmark.md), including forced `indexed` on a store with a valid index; verify result completeness, event delivery, timeouts and correctness.
4. [Test conversation reads](runbooks/07-thread-validation.md) against raw originals and multi-store examples.
5. Resolve missing schema/permission/security requirements in issues #2/#8 before any write tools.
6. After user approval and release gates, package and trial-deploy; never infer release readiness from CI alone.

Source of truth: [roadmap](https://github.com/Romboter/Outlook-Classic-MCP/issues/1), [issue #7](https://github.com/Romboter/Outlook-Classic-MCP/issues/7) and the [PR list](https://github.com/Romboter/Outlook-Classic-MCP/pulls).
