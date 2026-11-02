# Implementation status and release gates

**Repository snapshot: 2026-11-02.** Live observations below were reported by the owner; they have not been independently reproduced against the commit on PR #20. Recheck PR and issue states before merging.

| Area | Code state | Verification state |
| --- | --- | --- |
| Rust crate, stdio MCP, Cargo.lock, Windows CI — #2 | Draft PR [#12](https://github.com/Romboter/Outlook-Classic-MCP/pull/12) | Windows CI passed; complete Python schema snapshot outstanding |
| STA COM worker, attach/reconnect, probes — #3 | Draft PR [#13](https://github.com/Romboter/Outlook-Classic-MCP/pull/13) | Owner reports real Outlook attachment and reads succeeded for selected cases; reconnect and COM error behavior remain unverified on PR #20. |
| Six read-only mail/folder/thread tools — #4 / #7 | Draft PRs [#14](https://github.com/Romboter/Outlook-Classic-MCP/pull/14), [#18](https://github.com/Romboter/Outlook-Classic-MCP/pull/18) | Owner reports live thread and read results, but no test report has validated PR #20 head. Inspector protocol and negative-path matrix remain outstanding. |
| Table backend — #5 | Draft PR [#15](https://github.com/Romboter/Outlook-Classic-MCP/pull/15) | Owner reports equivalent results for a four-hit subject query across legacy, Table and indexed; controlled latency series outstanding. |
| Short handles — #6 | Draft PR [#16](https://github.com/Romboter/Outlook-Classic-MCP/pull/16) | Handles resolve in the reported live test; numbering reflects discovery order, not result rank. Cross-store and reconnect tests remain outstanding. |
| Indexed `AdvancedSearch` events — #5 | Draft PR [#17](https://github.com/Romboter/Outlook-Classic-MCP/pull/17) | Owner reports indexed subject search returned the same four results with ~6.5 s latency; **keep opt-in**. Broad correctness/performance cases remain unverified. |
| Permission and write-operation safety — #8 | Partial read-only groundwork | Permission modes and dangerous-operation protections not implemented |
| Full 31-tool Python parity — #9 | Not implemented | Not validated |
| MCP Apps UI parity — #10 | Not implemented | Not validated |
| Windows distributable/release — #11 | CI groundwork only | No release artifact or rollback validation |

### Pending verification of live-test fixes

- Draft PR [#20](https://github.com/Romboter/Outlook-Classic-MCP/pull/20) contains committed quote-trim, Table, thread and runbook fixes, stacked on documentation PR #19. No fresh Windows build or live Outlook report has been supplied for PR #20 head.
- Reported live observations on 2026-11-02: Outlook-native reply trimming omitted 5,696 of 11,153 characters with no newly authored text lost in the inspected sample; thread timing and `items_legacy` backend labeling returned; tested subject search yielded four matching hits across three backends without X.500-style `/O=` sender prefixes; a same-subject unrelated conversation was excluded; Unicode budgets (250 per item, 500 total) held in sampled emoji, Japanese, Russian and Arabic messages.
- `anchor_included` on PR #20 now compares candidate EntryID to anchor EntryID case-insensitively, rather than exact `(EntryID, StoreID)` text. This addresses a reported false negative but must be checked against cross-store identity and limited-result omission before the field can be trusted.
- The reported corrupted plain-text Microsoft 365 connector message is already corrupted in Outlook Body/HTMLBody (U+FFFD). Not attributable to the Rust transformation; no Rust workaround planned.
- PRs #12–#19 were found in the expected draft stack and GitHub Actions shows successful prior runs at their recorded heads; this does **not** verify PR #20 or satisfy interactive Outlook acceptance criteria.

The existing Python MCP implementation is unaffected. **All listed Rust PRs are drafts in a stacked feature chain, not merged into fork `main`.**

## Recommended next sequence

1. [Run Windows environment and COM validation](runbooks/03-outlook-com-validation.md) with a real Outlook Classic profile.
2. [Use MCP Inspector](runbooks/04-mcp-smoke-test.md) to test tool discovery and read calls without replacing working client config.
3. [Benchmark searches](runbooks/05-search-benchmark.md), including forced `indexed` on a store with a valid index; verify result completeness, event delivery, timeouts and correctness.
4. [Test conversation reads](runbooks/07-thread-validation.md) against raw originals and multi-store examples.
5. Resolve missing schema/permission/security requirements in issues #2/#8 before any write tools.
6. After user approval and release gates, package and trial-deploy; never infer release readiness from CI alone.

Source of truth: [roadmap](https://github.com/Romboter/Outlook-Classic-MCP/issues/1), [issue #7](https://github.com/Romboter/Outlook-Classic-MCP/issues/7) and the [PR list](https://github.com/Romboter/Outlook-Classic-MCP/pulls).
