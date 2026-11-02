# Python-to-Rust MCP tool contract inventory

**Baseline:** Python server in this fork at `src/outlook_mcp/tools/` and `src/outlook_mcp/client/`, version 0.3.2 at the start of the rewrite. This inventory enumerates all **31** existing `outlook_*` tools; it is a **migration checklist**, not a statement of Rust support.

**Current draft implementation (2026-11-02, through PR #20):** five names from the existing Python inventory are implemented experimentally in Rust: `outlook_whoami`, `outlook_list_folders`, `outlook_list_mails`, `outlook_search_mails`, and `outlook_get_mail`. The Rust branch stack additionally registers `outlook_get_thread` and `outlook_rust_status` (not part of the 31-tool original inventory). **The owner reported selected live Outlook Classic results on 2026-11-02; these do not constitute a complete validated run against PR #20. No Rust changes are merged to `main`.**

| Existing tool | Category | Operation | Python source | Rust milestone |
| --- | --- | --- | --- | --- |
| `outlook_list_mails` | Mail | Read | `tools/mail.py` | #4 |
| `outlook_search_mails` | Mail | Read | `tools/mail.py` | #4, #5 |
| `outlook_get_mail` | Mail | Read | `tools/mail.py` | #4 |
| `outlook_send_mail` | Mail | Write/send | `tools/mail.py` | #8, #9 |
| `outlook_reply_mail` | Mail | Write/send | `tools/mail.py` | #8, #9 |
| `outlook_forward_mail` | Mail | Write/send | `tools/mail.py` | #8, #9 |
| `outlook_move_mail` | Mail | Write | `tools/mail.py` | #8, #9 |
| `outlook_delete_mail` | Mail | Destructive | `tools/mail.py` | #8, #9 |
| `outlook_mark_mail` | Mail | Write | `tools/mail.py` | #8, #9 |
| `outlook_save_attachments` | Mail | Filesystem write | `tools/mail.py` | #8, #9 |
| `outlook_list_folders` | Folders | Read | `tools/folders.py` | #4 |
| `outlook_create_folder` | Folders | Write | `tools/folders.py` | #9 |
| `outlook_list_events` | Calendar | Read | `tools/calendar.py` | #9 |
| `outlook_get_event` | Calendar | Read | `tools/calendar.py` | #9 |
| `outlook_create_event` | Calendar | Write / possibly send | `tools/calendar.py` | #8, #9 |
| `outlook_update_event` | Calendar | Write | `tools/calendar.py` | #9 |
| `outlook_delete_event` | Calendar | Destructive | `tools/calendar.py` | #8, #9 |
| `outlook_respond_event` | Calendar | Write / possibly send | `tools/calendar.py` | #8, #9 |
| `outlook_list_contacts` | Contacts | Read | `tools/contacts.py` | #9 |
| `outlook_search_contacts` | Contacts | Read | `tools/contacts.py` | #9 |
| `outlook_get_contact` | Contacts | Read | `tools/contacts.py` | #9 |
| `outlook_resolve_name` | Contacts | Read | `tools/contacts.py` | #9 |
| `outlook_list_tasks` | Tasks | Read | `tools/tasks.py` | #9 |
| `outlook_create_task` | Tasks | Write | `tools/tasks.py` | #9 |
| `outlook_complete_task` | Tasks | Write | `tools/tasks.py` | #9 |
| `outlook_list_categories` | Categories | Read | `tools/categories.py` | #9 |
| `outlook_set_category` | Categories | Write | `tools/categories.py` | #9 |
| `outlook_list_rules` | Rules | Read | `tools/rules.py` | #9 |
| `outlook_toggle_rule` | Rules | Write immediately | `tools/rules.py` | #8, #9 |
| `outlook_get_out_of_office` | OOF | Read | `tools/ooo.py` | #9 |
| `outlook_whoami` | Account | Read | `tools/account.py` | #4 |

## Reported observations and open compatibility gaps (2026-11-02)

- Live owner observations: five Python-named read-only Rust tools, plus `outlook_get_thread` and `outlook_rust_status`; four matching subject hits across legacy, Table and opt-in indexed; Outlook reply history trimming; Unicode budgets; independent conversations sharing a subject remain separate. These observations require a tested commit and a redacted signed-off validation report before closing issues.
- Short handles are session-scoped **identity references, not search ranks**. The first time an `(EntryID, StoreID)` identity is seen, it receives the next `mN`; the same identity should retain its handle in that process until eviction/restart. Different result order does not imply different message identity. A per-backend mismatch requires investigation only if the same canonical identity gets inconsistent handles within one live process.
- Indexed search took approximately **6.5 seconds** for one reported simple subject search. Remains explicit opt-in; no speed advantage claimed.
- The `anchor_included` result was falsely negative on a reported thread even when the anchor appeared. PR #20 changes comparison to EntryID-only; check cross-store safety and recency filtering before treating resolved.
- Upstream Python still has incompatible defaults/behavior (`limit` 25 vs Rust 10, some 100 vs Rust 30 caps, markdown/json options, Python raw DASL and HTML support vs intentionally restricted Rust APIs). Preserve documented deliberate safety differences rather than silently exposing unsafe capabilities.
- A corrupted body already present within Outlook itself (U+FFFD in both Body and HTMLBody after Microsoft 365 connector submission) is not a Rust decoding defect.

## Compatibility details requiring real tests

- **Mail:** list pagination, `EntryID` and `StoreID`, folder names/paths, unread/flagged filters, date ranges, plain vs. HTML bodies, attachment metadata, error formats.
- **Search:** current `Items.Restrict` + DASL; optimize through indexed `AdvancedSearch` / `GetTable` only after comparing accuracy. Explicitly test unindexed stores and cross-folder scope.
- **Contacts:** local contact folders, Exchange directory/GAL, SMTP normalization, duplicate contacts.
- **Calendar:** time zones, recurrence expansion, attendee resolution, meeting cancellations; the existing `is_online_meeting` parameter is accepted but not implemented in the Python client.
- **Rules:** `Rules.Save()` changes live rules; require confirmation/permission.
- **OOF:** current implementation is read-only.
- **MCP Apps:** five existing `ui://outlook/*.html` resources (mail list, reader, calendar, contacts, tasks) to preserve under issue #10.
- **Security:** server-side read-only/draft/full modes; never retry an ambiguous non-idempotent send.

## Contract capture checklist

Before closing #2, capture and check a **complete machine-readable schema snapshot** from the Python server, covering each tool's input JSON Schema, defaults, descriptions, annotations, and result shapes. Maintain redacted fixtures only, no actual company emails. The table above is inventory; it is **not** a full schema snapshot.

## Additional tools (NOT part of 31-tool compatibility set)

- `outlook_search`: metadata-only indexed search; ~10 results by default.
- `outlook_search_more`: request additional results by an opaque cursor.
- `outlook_get_thread`: implemented experimentally in draft PR #18; live conversation/store behavior not verified.
- `outlook_rust_status`: implemented static diagnostic tool; does not test a live Outlook connection.

Do not advertise the proposed tools as implemented before corresponding issues close.
