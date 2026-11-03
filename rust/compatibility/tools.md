# Python-to-Rust MCP tool contract inventory

**Baseline:** Python server in this fork at `src/outlook_mcp/tools/` and `src/outlook_mcp/client/`, version 0.3.2 at the start of the rewrite. This inventory enumerates all **31** existing `outlook_*` tools; it is a **migration checklist**, not a statement of Rust support.

**Batch 1 (issue #9, draft):** `outlook_list_contacts`, `outlook_search_contacts`, `outlook_get_contact`, `outlook_resolve_name`, `outlook_list_categories`, `outlook_list_rules` and `outlook_get_out_of_office` are implemented experimentally in Rust (12 of 31 Python names now present). Deliberate deviations: scans are time- and item-bounded with `scan_truncated`, contact notes default to 4,000 characters (`max_notes_chars`), and contact `entry_id` values are short session handles; contact limits match Python (list 50/200, search 25/100). Owner workstation run on 2026-11-02 of the tools against live Outlook Classic returned the expected key shapes (counts and timings only, no mailbox content recorded); this is not a full parity test.

**Attachments (issue #8/#9, draft):** `outlook_save_attachments` (31 of 31 Python names). Same arguments as Python. Deviations: a person's approval of a prompt listing every file name and size and the folder; the folder must exist under the user profile; executable and script extensions (`.exe .js .ps1 .lnk` and similar) are refused per file with the file named and the reason, and the other files are still saved; linked (by-reference) attachments are refused; nothing is overwritten. Live-tested against an unsent draft with attachments.

**Meeting invitations and updates (issue #8, draft):** `outlook_create_event` accepts `attendees` (plain addresses) and needs approval; new `outlook_update_meeting` (deviation: the design said `outlook_update_event`, which still refuses meetings and now points at it). Python sent invitations immediately. Live-tested against a self-invite only; recurring meetings, attendee changes and online meetings are not supported.

**Cancel meeting (issue #8, draft):** new `outlook_cancel_meeting` (no Python equivalent; Python deleted meetings without telling attendees). Run live once against a self-invited test meeting (attendee is the owner only); other attendees, recurring series and failure paths untested. Inviting and updating meetings with attendees are still not implemented.

**Confirmed meeting responses (issue #8/#9, draft):** `outlook_respond_event` (30 of 31 names). Python answered immediately and mailed the organizer without anyone seeing it; Rust needs a person's approval of a prompt, refuses non-invitations, meetings you organized and cancelled meetings before any prompt, and `send_response=false` needs only the `events` capability. The accept, tentative and decline paths are verified against fakes and Microsoft's documentation only, not against a real received invitation. Still not implemented: `outlook_save_attachments`.

**Confirmed rule changes (issue #8/#9, draft):** `outlook_toggle_rule` (29 of 31 names). Python switched rules immediately with no explanation; Rust needs a person's approval of a prompt that says what the rule does, refuses a no-op or an ambiguous name, reads the state back and limits changes to 5 per hour as a prompt. Still not implemented: `outlook_respond_event` and `outlook_save_attachments`.

**Confirmed deletes (issue #8/#9, draft):** `outlook_delete_mail` and `outlook_delete_event` (28 of 31 names) plus the new `outlook_check_action`. Deviations from Python, which deleted immediately: every delete needs a person's approval of a prompt built from facts read from Outlook, items already in Deleted Items and meetings with attendees are refused, and an unsupported client means nothing is deleted. Still not implemented: `outlook_respond_event`, `outlook_toggle_rule`, `outlook_save_attachments`.

**Task and event writes (issue #9, draft, stacked on the send pipeline):** `outlook_create_task`, `outlook_complete_task`, `outlook_create_event` and `outlook_update_event` (26 of 31 names with the send tools). Event writes need the `events` capability, task writes `modify`. Deviations: `attendees` and `is_online_meeting=true` are rejected, `update_event` is refused for meetings with attendees, and `update_event` cannot change recurrence. Still not implemented: `outlook_delete_mail`, `outlook_delete_event`, `outlook_respond_event`, `outlook_save_attachments` and `outlook_toggle_rule`; the first four need a confirmation workflow, attachments need the path checks.

**First writes (issue #9, draft, stacked on #8 gating):** `outlook_mark_mail`, `outlook_set_category`, `outlook_create_folder` and `outlook_move_mail` are implemented (19 of 31 names). They need the `modify` capability: server mode `full` with `modify` not denied, otherwise they are refused before reaching Outlook. Deviations: `mark_mail` requires a change, `create_folder` omits the unusable folder `entry_id`, `move_mail` returns the new handle as `entry_id` and refuses Deleted Items. Send, reply, forward, delete, save-attachments, event/task writes and rule toggles are not implemented.

**Batch 2 (issue #9, draft, stacked on batch 1):** `outlook_list_events`, `outlook_get_event` and `outlook_list_tasks` are implemented experimentally (15 of 31 Python names). All read-only reads are now present in Rust. Remaining deliberate deviations: bodies/notes default to 4,000 characters (`max_body_chars` / `max_notes_chars` reach 50,000 and report `*_truncated`), contact and calendar scans are time-bounded and report `scan_truncated`, unset task dates are `null`, `entry_id` values are short session handles, and `end` must be after `start`. Offsets, `Z`, result UTC offsets, calendar window length and contact limits now match Python. Shapes and timings were checked on an interactive Outlook Classic profile on 2026-11-02; no mailbox content is recorded.

**Status (2026-11-02, draft PR #22):** five names from the existing Python inventory are implemented experimentally in Rust: `outlook_whoami`, `outlook_list_folders`, `outlook_list_mails`, `outlook_search_mails`, and `outlook_get_mail`. Rust also registers `outlook_get_thread` and `outlook_rust_status` (outside the original 31). The owner has reported selected **interactive Outlook Classic** read checks; the Python client function key-shapes were validated for **15 read-only tools via OutlookBridge STA**, but the FastMCP tool wrappers were **not** exercised. These observations do not prove complete parity or release readiness, and no Rust changes are merged to `main`.

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

The following artifacts cover all 31 Python tools. Source/CI files were produced without accessing Outlook; the separately supplied owner report covers 15 Python *client* calls against interactive Outlook Classic:

- [Python tool declarations](python-tools.declarations.json) — names, 104 declared input parameters, types, defaults, field constraints, tool descriptions, read/write annotations, resource URI associations, and return envelopes. These are an offline interpretation of the committed decorators, **not** the official FastMCP JSON Schema wire response.
- [Captured FastMCP tools/list metadata](python-tools.runtime.json) — **actual 31-tool registered input schemas**, descriptions, annotations and UI metadata captured on hosted Windows GitHub Actions, without initializing COM or opening Outlook. The runtime validator checks for drift from this committed snapshot. This proves registration metadata, **not** message-result correctness.
- [Python result field inventory and redacted fixtures](python-results.redacted.json) — code-derived field names, conditionals and synthetic placeholders. This now includes **type and shape observations from 15 live read-only Python client calls** (not wrapper calls). Fixtures remain synthetic, not copied from a mailbox, and are not strict verified result JSON Schemas.
- [Sanitized owner-reported live evidence](live-evidence-2026-11-02.md) — exact test boundary, per-call timings, Python-to-Rust differences and outstanding verification requirements. No subjects, addresses, bodies, or raw identifiers.
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

**Acceptance remains open:** Windows CI validated the 31 registered FastMCP tool schemas; the owner reports that all 15 read-only Python **client** result key shapes matched their source inventory on a real Outlook Classic STA thread. Actual FastMCP **wrapper** return shapes and conditional branches (attachment metadata, GAL, unresolved names and OOF error) are not yet live-verified. No mutation operations have been called; separate issue #8 safety gates are prerequisites. The redacted examples remain synthetic.

## Contract capture checklist

Before closing #2, reconcile the captured FastMCP input schema, defaults, descriptions, annotations and result shapes. Machine-readable input metadata is captured and tested in CI, and 15 read-only Python **client** result field sets have an owner-supplied live report. Still required: wrapper `CallToolResult` and `structuredContent` envelope, conditional/empty result shapes and an explicit parity decision for Rust differences. No mailbox content may enter fixtures or CI. Mutating tool **result shapes** can be source-reviewed without invoking any writes.

## Additional tools (NOT part of 31-tool compatibility set)

- `outlook_search`: metadata-only indexed search; ~10 results by default.
- `outlook_search_more`: request additional results by an opaque cursor.
- `outlook_get_thread`: implemented experimentally in draft PR #18; live conversation/store behavior not verified.
- `outlook_rust_status`: implemented static diagnostic tool; does not test a live Outlook connection.

Do not advertise the proposed tools as implemented before corresponding issues close.
