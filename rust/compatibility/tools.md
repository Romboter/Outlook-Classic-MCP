# Python-to-Rust MCP tool contract inventory

**Baseline:** Python server in this fork at `src/outlook_mcp/tools/` and `src/outlook_mcp/client/`, version 0.3.2 at the start of the rewrite. This inventory enumerates all **31** existing `outlook_*` tools; it is a **migration checklist**, not a statement of Rust support.

**Current draft implementation (2026-10-31, through PR #18):** five names from the existing Python inventory are implemented experimentally in Rust: `outlook_whoami`, `outlook_list_folders`, `outlook_list_mails`, `outlook_search_mails`, and `outlook_get_mail`. The Rust branch stack additionally registers `outlook_get_thread` and `outlook_rust_status` (not part of the 31-tool original inventory). **Live Outlook Classic validation is still outstanding; no Rust changes are merged to `main`.**

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

## Compatibility details requiring real tests

- **Mail:** list pagination, `EntryID` and `StoreID`, folder names/paths, unread/flagged filters, date ranges, plain vs. HTML bodies, attachment metadata, error formats.
- **Search:** current `Items.Restrict` + DASL; optimize through indexed `AdvancedSearch` / `GetTable` only after comparing accuracy. Explicitly test unindexed stores and cross-folder scope.
- **Contacts:** local contact folders, Exchange directory/GAL, SMTP normalization, duplicate contacts.
- **Calendar:** time zones, recurrence expansion, attendee resolution, meeting cancellations; the existing `is_online_meeting` parameter is accepted but not implemented in the Python client.
- **Rules:** `Rules.Save()` changes live rules; require confirmation/permission.
- **OOF:** current implementation is read-only.
- **MCP Apps:** five existing `ui://outlook/*.html` resources (mail list, reader, calendar, contacts, tasks) to preserve under issue #10.
- **Security:** server-side read-only/draft/full modes; never retry an ambiguous non-idempotent send.

## Machine-readable Python contract baseline (issue #2)

The following new **source-derived, unverified** artifacts cover all 31 Python tools and are committed without accessing Outlook:

- [Python tool declarations](python-tools.declarations.json) — names, 104 declared input parameters, types, defaults, field constraints, tool descriptions, read/write annotations, resource URI associations, and return envelopes. These are an offline interpretation of the committed decorators, **not** the official FastMCP JSON Schema wire response.
- [Python result field inventory and redacted fixtures](python-results.redacted.json) — observed code-level output keys, conditional-shape notes, and synthetic placeholder-only examples. These are **not** outputs from a real mailbox and are not strict verified result JSON Schemas.
- [Offline/registration validator](../../scripts/verify_python_contract_snapshot.py) — checks committed declarations against Python syntax; optionally compares the declarations with the actual local FastMCP \`tools/list\` registry without starting Outlook.
- [Contract regression tests](../../tests/test_python_contract_snapshot.py) — guard 31-tool coverage, representative defaults, nested recurrence, annotation flags, UI resource count and placeholder-only fixtures.
- [Contract CI](../../.github/workflows/python-contract-ci.yml) — static checks on Linux and FastMCP metadata comparison on Windows. A Windows CI pass **does not** prove Outlook COM behavior.

From the repository root, with Python 3.10+:

\`\`\`powershell
python scripts/verify_python_contract_snapshot.py --static
python -m unittest discover -s tests -p 'test_python_contract_snapshot.py' -v
\`\`\`

On a configured Windows development machine with the Python server's **existing** dependencies installed (but Outlook does not need to run):

\`\`\`powershell
python -m pip install -e .
python scripts/verify_python_contract_snapshot.py --runtime
python scripts/verify_python_contract_snapshot.py --capture-runtime
\`\`\`

The last command writes \`rust/compatibility/python-tools.runtime.json\` (actual FastMCP registration, containing **only** tool metadata and JSON Schemas). Review and commit this generated file after confirming it contains no mailbox content. If the runtime check fails, resolve the differences **without weakening compatibility checks** and recapture. No Python tool is invoked, and no COM thread is started by the metadata-only commands. Check the result payloads against *approved test fixtures only* on a controlled Windows Outlook Classic host; redact any subjects, addresses, content, identifiers and paths **before** proposing updates to the repository. Never add real mailbox samples to Git.

**Acceptance remains open:** the current snapshots document Python code declarations and possible result shapes; runtime metadata matching, result-shape reconciliation, and owner-reviewed live test evidence have not been established for this PR.

## Contract capture checklist

Before closing #2, capture and check a **complete machine-readable schema snapshot** from the Python server, covering each tool's input JSON Schema, defaults, descriptions, annotations, and result shapes. Maintain redacted fixtures only, no actual company emails. The table above is inventory; it is **not** a full schema snapshot.

## Additional tools (NOT part of 31-tool compatibility set)

- `outlook_search`: metadata-only indexed search; ~10 results by default.
- `outlook_search_more`: request additional results by an opaque cursor.
- `outlook_get_thread`: implemented experimentally in draft PR #18; live conversation/store behavior not verified.
- `outlook_rust_status`: implemented static diagnostic tool; does not test a live Outlook connection.

Do not advertise the proposed tools as implemented before corresponding issues close.
