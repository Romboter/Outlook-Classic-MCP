# Read-only Outlook Classic contract evidence — 2026-11-02

**Source:** owner-supplied sanitized live evidence. These observations have **not** been reproduced by this documentation author. No mail subjects, addresses, bodies, or raw COM identifiers are retained here.

**Environment:** Interactive Outlook Classic on Windows; Python virtual environment with `mcp==1.29.0` and pywin32, outside this repository. The `--runtime` FastMCP registration comparison passed for 31 declarations. No changes to server configuration.

## Test boundary

The Python calls below executed `client/*` functions through the real `OutlookBridge` STA worker. This **does not** test FastMCP wrappers, their `response_format` conversion, the `list_folders` `{count, items}` envelope, MCP Apps `ui_result`/`structuredContent`, or network/protocol startup. Names, field types, counts and timings were compared; **no values** were retained. All **15 of 15** tested read-only client operations completed without error and had the source-derived key sets (including conditional omissions). The remaining **16 mutating Python tools were not called**. Python has no `outlook_get_thread`.

## Python client observations

| Tool suffix (prefix: `outlook_`) | One call (ms) | Notes |
| --- | ---: | --- |
| `whoami` | 16 | `accounts` items: `display_name`, `smtp_address`, `user_name` strings; `account_type` integer |
| `list_folders` | 1,007 | Bare list: 74 folders at depth 2; `default_item_type` integer; wrapper envelope **not** tested |
| `list_mails` | 95 | `next_offset` present because more results existed |
| `search_mails` | 2,864 | Client function only |
| `get_mail` | 27 | `html_body` absent as requested; empty attachments array, attachment element schema **not** tested |
| `list_events` | 101 | Client keys matched |
| `get_event` | 38 | Attendee `name`/`address` strings; `type`/`response` integers |
| `list_contacts` | 5,618 | One call only; slow on this profile |
| `search_contacts` | 5,957 | `include_directory=False`; GAL rows untested |
| `get_contact` | 14 | Client keys matched |
| `resolve_name` | 18 | Resolved branch only; unresolved `note` untested |
| `list_tasks` | 53 | Client keys matched |
| `list_categories` | 74 | `color` integer |
| `list_rules` | 775 | Client keys matched |
| `get_out_of_office` | 81 | Success branch: `out_of_office` bool, `status` string. Source error branch adds `note`, not observed live |

**Timing caution:** These are individual observed calls, not repeatable cold/warm benchmark results. Do not infer relative speed generally from this table. The typed observations and conditional fields are recorded in [`python-results.redacted.json`](python-results.redacted.json). Synthetic examples remain clearly labeled; no real returned values have been copied into fixtures.

## Rust comparison, separate read-only tests

- `outlook_rust_status` advertises `outlook_connected:false` statically even when mailbox tools work. This is **not** a connection check; avoid describing it as live status.
- `outlook_whoami` returns `current_user` only. Python additionally has `accounts`, `local_time`, `timezone`, `utc_offset`. **Output parity gap.**
- Same reported body-inclusive search, three matching messages in the same order for each successful backend:

  | Requested backend | Actual backend | Single observed `elapsed_ms` |
  | --- | --- | ---: |
  | `auto` | `items_restrict_legacy` | 3,024 |
  | `legacy` | `items_restrict_legacy` | 3,033 |
  | `indexed` | `advanced_search_indexed` | 10,968 |
  | `table` | Rejected body search (expected unsupported combination) | n/a |

  Python `search_mails` took 2,864 ms on a comparable query. These results **do not** justify promoting indexed to `auto`; it took approximately 3.6 times legacy's time in this one comparison. More representative benchmarks are needed for issue #5.
- Rust search item fields observed: `entry_id` (short handle), `from`, `from_address`, `received_local`, `subject`, `unread`. Python also provides `to`, `flagged`, `has_attachments`, `importance`, `preview`; Rust uses differently formatted `received_local` instead of Python `received` with an ISO-8601 local offset.
- Rust `outlook_get_mail` on a single item exposed `body`, `body_total_chars`, `body_truncated`, `cc`, `entry_id`, `from`, `from_address`, `received_local`, `subject`, `to`, `unread`. Missing versus Python: `conversation_id`, `bcc`, `sent`, `flagged`, `importance`, `categories`, `attachments`. Body truncation behaved as intended for this sample.
- Rust-only `outlook_get_thread`: one-message verified conversation-ID result in 29 ms, with the expected truncation and bookkeeping fields. Does not provide parity evidence for the Python 31-tool set.

## Defects or gaps requiring independent follow-up

1. **Sender address completeness (#5).** For some Exchange senders, Rust `legacy`/`auto` returns an empty `from_address`, whereas `indexed` returned an SMTP address for the same item. The Rust code currently suppresses Exchange X.500 `/O=` identifiers; ensure an optional **safe SMTP resolution** strategy rather than exposing X.500 addresses. Do not open every full MailItem on Table/indexed searches just to resolve senders.
2. **Handle stability (#6).** The owner observed the same message with different `mN` handles across legacy and indexed searches, and with yet another handle when returned as the thread anchor. The in-memory cache keys by exact `(EntryID, StoreID)` strings. Potential causes include different COM identity encodings or different MCP process sessions; neither is proven. Test all backends in **one persistent MCP session** before modifying aliasing. Do not deduplicate by EntryID alone without verifying store identity.
3. **Indexed latency (#5).** Approximately 10.97 s indexed vs. 3.03 s legacy in this one case; remain opt-in. Include cold/warm runs and full search semantics when benchmarking.
4. **Result parity (#4/#9).** Prioritize Python-compatible `whoami`, complete selected-mail metadata and date representation, with capped context sizes; preserve intentionally stricter Rust security defaults.
5. **Unverified conditional shapes (#2).** FastMCP wrapper envelopes, actual attached-file item fields, directory/GAL rows, unresolved name output and OOF exception `note` branch have not been live-tested.

## Suggested verification, still pending

Use one long-lived MCP Inspector session on Outlook Classic, collect only Boolean comparisons and elapsed times, not identifiers or mailbox values. For one approved message reached through legacy, Table/indexed where applicable, and thread:

- Verify whether `entry_id` short handles agree within the same process. Verify `outlook_get_mail` resolves every returned handle to the correct item. Do not infer identity from subject alone.
- Compare source StoreIDs and EntryIDs **only inside the COM STA**, report whether the values match (yes/no), not the values.
- Verify `anchor_included` on a one-message thread and a multi-store thread; check newest-N omission.
- Run the Python **FastMCP wrappers** separately with read-only calls (including one empty and one nonempty list), inspect the `content`/`structuredContent` envelope and markdown/json formatting locally, and report key names and types only.
- Keep the existing Python server and all live client/plugin config unchanged. Do not run any mutating tools until #8's server-side permission gates are implemented.

**Acceptance:** issue #2 has source declarations, real FastMCP registration metadata and client-only live field observations. It is **not closed**, because wrapper/output envelope validation and some conditional shapes are still outstanding; changes to the Rust compatibility surface belong in scoped, separately tested PRs. No merge without explicit owner instruction.
