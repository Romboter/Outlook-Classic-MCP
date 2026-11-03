# Rust MCP tool reference

**As of the issue #9 batch 2 branch (draft, stacked on batch 1):** One diagnostic tool and sixteen **experimental, read-only** Outlook Classic tools (six mail/folder/thread, seven contacts/categories/rules/out-of-office, and three calendar/task reads) are registered in `rust/src/lib.rs`. This inventory is grounded in `rust/src/contracts.rs`; it is not a live Outlook verification.

| Tool | Main inputs | Defaults / limits | Output or caveat |
| --- | --- | --- | --- |
| `outlook_rust_status` | none | — | Static implementation flags; `outlook_connected:false` is **not** an actual connection test |
| `outlook_whoami` | none | — | Outlook current user display name |
| `outlook_list_folders` | `root?`, `max_depth?` | depth 2 / max 4; max 100 folders | Default mailbox root, not a general delegated-store explorer |
| `outlook_list_mails` | `folder?`, `limit?`, `offset?`, `unread_only?`, `from_address?`, `backend?` | limit 10 / max 30; offset 0..1000 | Metadata-only, includes actual backend; `since` and `until` are **not implemented** and return errors |
| `outlook_search_mails` | `query`, `folder?`, `scope?`, `limit?`, `backend?`, `search_subfolders?` | query length 1..160; limit 10 / max 30; scope `subject_body`; backend `auto` | Backend: `auto`, `table`, `legacy`, `indexed`; only indexed accepts `search_subfolders` |
| `outlook_get_mail` | `entry_id`, `store_id?`, `include_body?`, `include_html?`, `max_body_chars?` | body on; 4,000 chars default / 12,000 max | Selected plain-text message; HTML bodies disabled |
| `outlook_list_contacts` | `limit?`, `offset?` | limit 50 / max 200 (matches Python); offset 0..5000 | Contact folders in every store, default store first; 15 s budget and 5,000-item scan cap, `scan_truncated` reports a partial scan. `entry_id` is a short handle |
| `outlook_search_contacts` | `query`, `limit?`, `include_directory?` | query 1..160; limit 25 / max 100 (matches Python); directory on | All words must match name/email/company/title; then the Exchange GAL by display name (38 s overall ceiling); `scan_truncated` when a budget was hit |
| `outlook_get_contact` | `entry_id`, `store_id?`, `max_notes_chars?` | notes 4,000 chars default / 50,000 max | Handle from list/search or raw EntryID; adds `notes_truncated` |
| `outlook_resolve_name` | `name` | 1..160 chars | Ctrl+K-style resolve; no UI; ambiguous returns `resolved:false` |
| `outlook_list_categories` | none | — | Profile color categories (name, color index) |
| `outlook_list_rules` | none | — | Default-store rules (index, name, enabled); read-only |
| `outlook_get_out_of_office` | none | — | `out_of_office`/`status` on, off, or unknown when the MAPI property is unreadable |
| `outlook_list_events` | `start?`, `end?`, `limit?`, `include_recurrences?` | window default now to +14 days, no maximum; limit 50 / max 200 (matches Python) | Default-store calendar. ISO-8601 local times or times with `Z`/an offset (converted to local); results carry the local UTC offset like Python. Recurrences expanded by default; `has_more` and `scan_truncated` added; end must be after start |
| `outlook_get_event` | `entry_id`, `store_id?`, `max_body_chars?` | body 4,000 chars default / 50,000 max | Handle from `list_events` or raw EntryID; includes attendees (max 500), reminder, categories; `body_truncated` added. Rejects non-appointment items |
| `outlook_list_tasks` | `limit?`, `include_completed?` | limit 50 / max 200 (matches Python); open tasks only by default | Default Tasks folder sorted by due date; dates carry the local UTC offset; unset dates are `null` (Python emitted the year-4501 placeholder) |
| `outlook_get_thread` | `entry_id`, `store_id?`, `max_messages?`, `include_bodies?`, `trim_quoted?`, `per_message_chars?`, `total_body_chars?` | 10 / max 30 messages; per-body 1,500 / max 6,000; total 8,000 / max 24,000 | Exact ConversationID; only recognized quote tails trimmed, omissions flagged |

All mail list/search/thread messages should return compact in-memory `entry_id` handles such as `m1` rather than raw EntryIDs/StoreIDs. Handles expire on restart and may be evicted; repeat search/list when `outlook_get_mail` says a reference is unknown. Raw IDs can still be supplied for compatibility but **do not share them in issues**.

## Example call sequence (for a trusted test profile)

1. Start the MCP server with no command-line flags; connect an MCP Inspector or test client.
2. Invoke `outlook_rust_status`; it only reports static capability flags.
3. Invoke `outlook_whoami`; verifies Outlook COM access.
4. Invoke `outlook_search_mails` with `{"query":"test","scope":"subject","limit":10,"backend":"legacy"}` using a **safe test query**.
5. Select a result's `entry_id` and invoke `outlook_get_mail` with `{"entry_id":"m1","max_body_chars":1000}`.
6. Optionally call `outlook_get_thread` with `{"entry_id":"m1","max_messages":5,"per_message_chars":500,"total_body_chars":2000}`. This depends on actual store conversation support.

Do not assume sample `m1` exists: use the handle returned in *your* current MCP session. A tool response may contain confidential mail; never paste it into GitHub.

## Search semantics worth knowing

- `scope=subject_body` may select `Items.Restrict` in auto mode; `table` mode explicitly rejects body search.
- `backend=indexed` is opt-in, requires an Outlook Instant Search-enabled store and event delivery; no performance claim yet.
- Returned `received_local` dates are Windows-locale text and must not be treated as UTC.
- Returned `has_more` / `scan_truncated` are not proof all mailbox results have been checked.
- No `outlook_send_mail`, delete, move, calendar writes, file saves, or similar Rust mutation tools are registered.

For the **existing Python** 31-tool inventory, consult [`rust/compatibility/tools.md`](../rust/compatibility/tools.md). That list is a migration target, not the number of currently implemented Rust tools.
