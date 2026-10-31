# Runbook 05 — Measure search correctness and performance

**Use when:** Deciding whether `Folder.GetTable` or opt-in `AdvancedSearch` is faster and accurate enough to adopt for real mail.

**Detailed baseline:** [Rust search benchmark guide](../../rust/docs/search-benchmark.md). Do not report speed improvements based only on CI.

## Preflight

- Complete [real COM](03-outlook-com-validation.md) and [MCP](04-mcp-smoke-test.md) validation.
- Ensure Outlook Classic has finished initial indexing and cached the intended time period.
- Use a trusted test mailbox or preapproved nonsensitive query; never publish identifiers or message bodies.
- Record Outlook build, store category (Exchange cached, PST, shared store as applicable) and whether `instant_search_enabled` reports true/false/null.

## Test matrix

| Variable | Values |
| --- | --- |
| Backend | `legacy`, `table`, `indexed` (explicit opt-in only) |
| Search scope | `subject`, `from`, `subject_body` |
| Limit | 10 for normal comparison; 30 for stress |
| Query selectivity | Zero, few, many results |
| Cache | One cold, 10 warm iterations |
| Folder | Small and large Inbox-like folders; optional PST/delegated store |

**Important:** `table` does not support `subject_body`. Compare only backends that can execute the **same semantics**. `indexed` uses `ci_phrasematch`, which may differ from `LIKE`: verify result equivalence before comparing speed. Don't call non-equivalent queries identical.

## Steps

1. Use Inspector (see [runbook 04](04-mcp-smoke-test.md)) with the same query and folder. Example test-only inputs:
   ```json
   {"query":"test","folder":"Inbox","scope":"subject","limit":10,"backend":"legacy"}
   ```
   ```json
   {"query":"test","folder":"Inbox","scope":"subject","limit":10,"backend":"table"}
   ```
   ```json
   {"query":"test","folder":"Inbox","scope":"subject","limit":10,"backend":"indexed","search_subfolders":false}
   ```
2. Check `backend`, `count`, `has_more`, `scan_truncated`, `instant_search_enabled` and `elapsed_ms`. Never compare a successful Table rowset to a silently reported legacy result; backend labels matter.
3. Verify *which actual test messages* matched **locally**. Distinct filter semantics, missing entries, unexpected folders or stale indices invalidate a speed-only conclusion.
4. Record one cold run and 10 warm runs; compute median, p95, success rate and approximate returned bytes. If results are incomplete or different, document correctness first.
5. Repeat on a large folder and at least one unindexed store if available. `backend:indexed` should report an error when it cannot confirm an enabled search index, not silently fall back.
6. Inspect the 18-second indexed wait and 45-second caller timeouts on a safe test. An Outlook modal dialog can leave the STA blocked even when a caller has timed out.

## Expected result

A reproducible, anonymized timing/error table with correctness notes per backend and workload. There is **no predetermined fastest backend**. Only promote indexed into the auto path after reliable event delivery and correct results are established over representative workloads.

## Failure / rollback

- Indexed timeout/event error: switch to `backend:legacy` or `auto`; do not change production configuration or repeatedly hammer Outlook.
- Table column/filter unsupported: use explicit `legacy`; record the error class.
- Inconsistent results: flag the issue as a blocker and compare raw messages locally.

Record results in [validation template](../templates/validation-report.md); do not include query content if it is sensitive.
