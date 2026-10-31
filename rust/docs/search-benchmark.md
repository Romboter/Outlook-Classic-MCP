# Outlook Classic search backends — workstation benchmark

**Status: experimental / no live Outlook measurements yet.** Use these steps only on a trusted Windows workstation running Outlook **Classic** with an active cached profile. This is not an OST parser or an Exchange/Microsoft Graph client.

## Backends available in this branch

| Backend | Mode | What it does |
| --- | --- | --- |
| `folder_table` | `backend: "table"` | Calls `Folder.GetTable`, requests a small column set, sorts by `ReceivedTime`, and reads `Row.Item` values without opening a `MailItem` for each hit |
| `items_restrict_legacy` | `backend: "legacy"` | Existing Python-like `Items.Restrict` search and per-result item access |
| `auto` (default) | `backend: "auto"` | Tries Table for **subject** and **from** search; uses legacy for `subject_body` and falls back when the store does not support the requested Table columns or filters |

The list-mail tool also accepts `backend: "table"`, `"legacy"`, or `"auto"`. List results are always capped at 30. A Table query may scan up to **10,000 table rows**, returning `scan_truncated: true` when that safety ceiling is reached; in that case the result is **not exhaustive**. The limit bounds the number of *returned* items, not always the number of Table rows examined.

`Store.IsInstantSearchEnabled` is returned as `instant_search_enabled: true/false/null` where available. That property alone **does not mean a Folder.GetTable query was indexed**. The indexed multi-folder `AdvancedSearch` + `Search.GetTable` backend, including completion events, has not been implemented yet.

## Measure without exporting sensitive mail

1. Open Outlook Classic. Confirm it is connected and that mail from the expected time period is cached locally.
2. Start the Rust MCP server and connect through an MCP-compatible inspector/client. Do not configure it in place of the production Python server.
3. Choose one folder and several **non-sensitive** queries: exact subject word, sender substring, high-hit query, low-hit query. Avoid sharing returned messages/logs.
4. Invoke `outlook_search_mails` with identical `query`, `folder`, `scope: "subject"`, and `limit: 10`; vary only `backend: "table"` versus `backend: "legacy"`. Repeat with `scope: "from"`.
5. For folder listings, compare `outlook_list_mails` with identical `folder`, `limit: 10`, `offset: 0`, `unread_only` and sender filter.
6. Record only response `backend`, `elapsed_ms`, `count`, `has_more`, `scan_truncated`, and `instant_search_enabled`. Exclude real subjects, senders, EntryIDs, StoreIDs, email bodies, and anything identifying colleagues or customers.
7. Test 10 warm runs per backend after one cold run, then compare **median, p95 and first-run latency**. Repeat against a large folder and an unindexed/PST store if available. Compare hit correctness by visually matching returned items locally; do not upload identifiers.

Example parameters for a deliberately generic test:

```json
{"folder":"Inbox","query":"weekly","scope":"subject","limit":10,"backend":"table"}
```

```json
{"folder":"Inbox","query":"weekly","scope":"subject","limit":10,"backend":"legacy"}
```

If `backend: "table"` errors, record only the class of error and the Outlook build/store type. If `auto` falls back to legacy, the response reports `table_fallback: true`. Do **not** report timings from different scopes as apples-to-apples comparisons.

## Known caveats

- Outlook Table rows are a documented lightweight rowset, but performance depends on store provider, cache/indexing, query selectivity, and mailbox size.
- `subject_body` intentionally stays on the legacy path until a correctly event-driven indexed search is implemented. Table body-content filters may be unsupported.
- Some stores may not expose sender columns or the `PR_MESSAGE_FLAGS` field. `unread` may therefore be `null`, and the list tool may use its legacy fallback.
- Times in responses remain `received_local` display strings; do not interpret them as UTC.
- Existing handles are long Outlook EntryIDs. Compact server-side handles belong to issue #6.
- The COM timeout protects the caller but does not cancel an in-progress Outlook call.
- **No performance claims should be published until the same query has been run on real Outlook Classic.**

## References

- https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.gettable
- https://learn.microsoft.com/en-us/office/vba/api/outlook.table
- https://learn.microsoft.com/en-us/office/vba/api/outlook.row.item
- https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-filter-and-efficiently-enumerate-items-in-a-folder
- https://learn.microsoft.com/en-us/office/vba/api/outlook.search.gettable
