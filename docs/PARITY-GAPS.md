# Parity gaps: Rust Outlook MCP vs the Python baseline

Compared on branch `feat/rust-list-mails-date-filter` (HEAD `dcda2a8`), by reading code only. Python baseline: `src/outlook_mcp/tools/*.py` and `src/outlook_mcp/client/*.py`. Rust: `rust/src/contracts.rs` (params, validation), `rust/src/outlook/*.rs` (COM), `rust/src/lib.rs` (registration). Nothing was run and no mailbox was touched. "Untested-live" items come from `HANDOFF.md` and `rust/compatibility/live-evidence-2026-11-02.md`, not from new tests.

Where docs and code disagree, code wins. Stale doc statements found: `docs/STATUS.md` says MCP Apps parity (#10) is "Not implemented", but `rust/src/ui.rs` and `lib.rs:501-607` link all five `ui://outlook/*.html` views. `docs/STATUS.md` also says "Full 31-tool parity not implemented" while all 31 names are registered (38 tools).

Severity: **blocks** = blocks real daily use; **nice** = nice to have; **deliberate** = decision to keep. Size: S under half a day, M about a day, L multi-day or needs design.

## Ranked summary: top 10 for daily use (searching and pulling recent mail, calendar, tasks)

| Rank | ID | Gap | Size |
| --- | --- | --- | --- |
| 1 | G01 | `outlook_list_mails` declares `since`/`until` but rejected them. **Fixed in PR #41** (inclusive local wall-clock bounds, table and legacy paths; live read-only check passed, see the PR). | done |
| 2 | G02 | Mail timestamps were `received_local` locale text. **Fixed in PR #43**: `received` is now ISO-8601 local time with offset (threads: `sent` and `received`); live read-only check passed. | done |
| 3 | G03 | Multi-word search is a contiguous phrase match. Python matched all words anywhere. "budget forecast" misses "forecast for budget". | M |
| 4 | G06 | `from_address` was empty for Exchange (X.500 `/O=`) senders. **Fixed in PR #43**: SMTP comes from PR_SENDER_SMTP_ADDRESS (blank 2/30 before, 0/30 after in a live read-only check). | done |
| 5 | G05 | `outlook_get_mail` omits the attachments list (index, filename, size), `flagged`, `importance`, `categories`, `conversation_id`, `sent`, `bcc`. The model cannot pick an `attachment_index` for `save_attachments`. | S |
| 6 | G04 | List and search rows omit `to`, `flagged`, `has_attachments`, `importance`, `preview`. Triage needs one `get_mail` per row, which costs more tokens. | M |
| 7 | G11 | Folder specs: no `junk`/`spam`/`trash` aliases and no store-name prefix, so archive, PST or shared stores cannot be addressed. Default mailbox only. | M |
| 8 | G10 | No way to read past the 12,000-char body ceiling (no offset or continuation). Long threads and newsletters are cut with no recourse. | S |
| 9 | G22 | `outlook_list_folders` omits `unread_count` and `default_item_type`, and caps at 100 rows and depth 4. | S |
| 10 | G25 | `outlook_whoami` returns only `current_user`: no accounts, `local_time`, `timezone`, `utc_offset`. | S |

Calendar and task reads are close to Python (G26 to G33 are mostly write-side or beyond-Python). Contacts, categories, rules and out-of-office reads match Python field for field.

Counts: 31 true gaps (30 actionable plus G30, information only; IDs are stable and deliberately not contiguous), 20 deliberate deviations (D01 to D20), 12 untested-live items (U01 to U12).

## 1. True gaps

### outlook_list_mails

| ID | Tool | Gap | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| G01 | list_mails | **FIXED in PR #41.** Was: `since`/`until` declared, then rejected with "not yet implemented". Python filtered `ReceivedTime` via Jet `Restrict` (12-hour AM/PM format). | client/mail.py:126-160 | contracts.rs:19-38 (params), contracts.rs:1001-1006 (rejection), contracts.rs test `rejects_unimplemented_options_rather_than_ignoring_them`; table path takes `filter: ""` at mailbox.rs:247-257 | blocks | M |
| G02 | list/search/get_mail | **FIXED in PR #43.** Was: `received_local` is `VariantChangeType` to a locale-dependent string. Python `received` was ISO-8601 with the local offset, and `get_mail` also had `sent`. | client/mail.py:61-75, formatting.py `to_iso` | mailbox.rs:28-46 (`text_property`), mailbox.rs:110-119, table.rs:224-232. The Rust events/tasks code already has `date_iso` with offset at schedule.rs:127-143 and can be reused. | blocks | M |
| G04 | list/search_mails | Row fields missing: `to`, `flagged`, `has_attachments`, `importance`, `preview` (200 chars). Rust rows have `entry_id`, `store_id`, `subject`, `from`, `from_address`, `received_local`, `unread`. | client/mail.py:61-75 | mailbox.rs:110-119, table.rs:224-232 | nice (see rank 6) | M |
| G06 | list/search/get_mail | **FIXED in PR #43.** Was: `public_address()` blanks any address starting with `/`, so Exchange senders show an empty `from_address`. Python returned `SenderEmailAddress` as is. A safe SMTP lookup is not implemented (live evidence item 1). | client/mail.py:61-75 | table.rs:69-75, mailbox.rs:115 | blocks | M |
| G07 | search_mails | `scope=from` does not match the SMTP address of Exchange senders. Python OR-ed `PR_SENDER_SMTP_ADDRESS` (0x5D02001F) into the `from` filter. | client/mail.py:31, 213-220 | mailbox.rs:325-331 (no SMTP proptag) | nice | S |
| G03 | search_mails | Whole query is one `LIKE '%phrase%'`. Python picked the longest word as the Restrict anchor and verified the other words per item. | client/mail.py:34-48, 186-244 | mailbox.rs:312, 322-331, collect_items mailbox.rs:136-180 | blocks | M |
| G08 | list_mails | Result lacks `next_offset` (Python gave it when the page was full). Rust gives `has_more` only. | client/mail.py:177-183 | mailbox.rs:260-267, 289-298 | nice | S |
| G11 | list/search/move/create_folder | Folder resolver gaps: no `junk`, `spam`, `trash`, `calendar`, `contacts`, `tasks`, `notes` aliases; no store DisplayName prefix (cannot reach other stores, shared mailboxes, PST or archive); root is the default store only. | folders.py:25-69, constants.py:23-36 | mailbox.rs:59-107 (only inbox, sent, drafts, deleted, outbox, then default-mailbox-relative path) | nice (blocks if the user reads archives) | M |
| G12 | list_mails | `offset` is capped at 1000 (Python unbounded) and the folder table scan is capped at 10,000 rows. The cap is reported, but pages deep in a large Inbox cannot be reached. | client/mail.py:126 | contracts.rs:1000, table.rs:18 | nice | S |

### outlook_get_mail

| ID | Tool | Gap | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| G05 | get_mail | Missing: `attachments[] {index, filename, size_bytes}`, `conversation_id`, `bcc`, `sent`, `flagged`, `importance`, `categories`. Present: summary fields plus `to`, `cc`, `body`, `body_truncated`, `body_total_chars`. | client/mail.py:78-123 | mailbox.rs:379-411 | blocks | S |
| G09 | get_mail | `include_html=true` is rejected ("disabled in this read-only prototype"). Python returned `html_body`, opt-in. | client/mail.py:117-121, tools/mail.py:121-131 | contracts.rs:1046-1048, contracts.rs:55-66 | nice | M (needs a safety review) |
| G10 | get_mail | Rust body ceiling is 12,000 chars and `max_body_chars=0` (unlimited) is not offered. There is no offset or "next chunk" parameter, so the remainder is unreachable. `get_thread` is capped at 6,000 per message. | client/mail.py:104-111 (`0` means no limit, default 10000) | contracts.rs:1049-1051, mailbox.rs:392-404 | nice | S |

### Mail writes

| ID | Tool | Gap | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| G13 | send_mail | `attachments` rejected ("not supported yet"). Python attached validated paths under the user profile. | client/mail.py:264-311, utils/paths.py:38-60 | contracts.rs:199-202, 262-272, compose.rs:195-212. `safe_path.rs` already holds the path checks. | nice | M |
| G14 | send/reply/forward | `html=true` rejected on all three. Python set `HTMLBody`. | client/mail.py:285-292, 331-335, 361-366 | contracts.rs:262-268, 283, 317, 331 | nice | M |
| G15 | reply_mail | `attachments` rejected on reply. | client/mail.py:314-345 | contracts.rs:210-221, 317 | nice | M |
| G16 | move_mail | Result key renamed: Python `new_entry_id`, Rust `entry_id` (plus `store_id`). Clients written against Python read nothing. | client/mail.py:379-383 | modify.rs:121-146 | nice | S |
| G17 | save_attachments | Result shape differs. Python `{status, count, files[paths], output_dir}`; Rust `{done, saved}` (and `not_done` on failure). | client/mail.py:415-481 | attachments.rs:171-211 | nice | S |

### Folders, account

| ID | Tool | Gap | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| G22 | list_folders | Rows omit `unread_count` and `default_item_type`. Defaults differ (depth default 2, max 4, 0 allowed; Python default 4, range 1 to 10). Output is capped at 100 rows (`truncated`). `root` cannot name another store. | client/folders.py:88-113, tools/folders.py:26-37 | mailbox.rs:182-229, contracts.rs:11-17 | nice | S |
| G23 | create_folder | `parent` has the same resolver limits as G11 (no store prefix). The result omits `entry_id` (deliberate, see D09). | client/folders.py:116-123 | modify.rs:90-103 | nice | S |
| G25 | whoami | Returns only `current_user`. Python also returned `accounts[]` (display_name, smtp_address, user_name, account_type), `local_time`, `timezone`, `utc_offset`. Python's comment says agents should never guess the zone. | client/account.py:11-31 | mailbox.rs:421-424, lib.rs:485-487 | nice (needed for relative dates) | S |

### Calendar

| ID | Tool | Gap | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| G26 | create_event | `reminder_minutes` omitted always sets a 15-minute reminder. In Python `None` meant no reminder, and `0` to 10080 was accepted. Rust: no way to create without a reminder, and the max is 525,600. | client/calendar.py:127-169 (`if reminder_minutes is not None`) | schedule_write.rs:157-162, contracts.rs:491-493 | nice | S |
| G27 | create_event | Recurring meetings with attendees are rejected ("not supported yet"). Python could do both together. | client/calendar.py:150-165 | contracts.rs:609-611 | nice | M |
| G28 | create_event | `is_online_meeting=true` rejected. Python accepted and silently ignored it (no Teams link either way). Rust is more honest, but Teams meeting creation does not exist in either server. | client/calendar.py:127-139 (param never used) | contracts.rs:406-407, 474-476 | nice | L |
| G29 | create_event | Attendees must be plain SMTP addresses, 1 to 100, no display names or resolvable names. Python accepted anything `Recipients.Add` resolves. | client/calendar.py:150-156 | contracts.rs:572-596 | nice | S |
| G30 | list_events | Result `start`/`end` echo the window without an offset (Python `isoformat()` of the naive value, so also no offset: parity). Rust items carry an offset. Not a gap, noted so nobody "fixes" it. | client/calendar.py:116-121 | schedule.rs:221-228 | nice (info only) | S |
| G31 | get_event | Python's docstring promises "recurrence info", but neither server returns recurrence pattern, `MeetingStatus`, the user's own response status, or an online-meeting link. Not a Rust-only gap, but the biggest calendar hole for daily use. | tools/calendar.py:60-76 | schedule.rs:230-272 | nice | M |
| G32 | calendar (all) | Default calendar only; shared, delegate and secondary calendars are unreachable. Same in Python. | client/calendar.py:90 | schedule.rs:169-228 (`OL_FOLDER_CALENDAR`) | nice | M |

### Tasks, contacts, misc

| ID | Tool | Gap | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| G33 | tasks | No `get_task`, `update_task` or task body in `list_tasks`. Same in Python (list shows no notes). Listed because tasks are a daily need. | client/tasks.py:18-81 | schedule.rs:274-322 | nice | M |
| G34 | all tools | `response_format` (`markdown`/`json`) is not declared; it is ignored if sent, and output is always compact JSON. Python defaulted to markdown. Not a functional gap, but clients that pass it get no error. | tools/*.py `response_format` | lib.rs tool signatures; no `deny_unknown_fields` anywhere in `rust/src` | nice | S |
| G35 | search_contacts | GAL scan stops at 50,000 entries, contact folder scan at 5,000 items and 15 s per folder pass, 38 s per call. Reports `scan_truncated`. Python was unbounded (slow but complete). | client/contacts.py:103-139 | directory.rs:21-27, 189-221 | nice | S |
| G36 | tools/list | Write tools are listed in every mode even when the server mode cannot run them (known follow-up in HANDOFF.md). Model sees 38 tools, many inert in read-only mode. | n/a (Python had no modes) | lib.rs:1168-1209 (test asserts all 38) | nice | M |

## 2. Deliberate deviations (keep)

All are recorded as decisions in `HANDOFF.md`, `docs/DECISIONS.md`, `docs/CONFIRMED-ACTIONS-DESIGN.md` or `rust/compatibility/tools.md`, and verified in code.

| ID | Tool | Deviation | Python behavior | Rust location | Severity | Size |
| --- | --- | --- | --- | --- | --- | --- |
| D01 | all reads | `entry_id` is a short session handle (`m1`), table capped at 2,048, lost on server restart, "Unknown or expired" error. Raw EntryIDs are still accepted. | raw EntryID | refs.rs:12, 101-118 | deliberate | n/a |
| D02 | list/search | Limits 1..30, default 10 (Python 1..100, default 25). Token budget. | tools/mail.py:51, 86 | contracts.rs:846-852 | deliberate | n/a |
| D03 | get_mail | Body default 4,000 chars, max 12,000. Python default 10,000, `0` unlimited. | client/mail.py:78-123 | contracts.rs:1049, mailbox.rs:392 | deliberate | n/a |
| D04 | search_mails | `scope=dasl` (raw DASL) removed. Query text is always an escaped literal. Search stays metadata-only. | client/mail.py:206-209 | contracts.rs:1013-1019 | deliberate | n/a |
| D05 | get_event / get_contact | Bodies and notes default to 4,000 chars (max 50,000), with `body_truncated` / `notes_truncated`. | client/calendar.py:50-68 | schedule.rs:244-266, directory.rs:323 | deliberate | n/a |
| D06 | send/reply/forward/send_draft | Needs server mode `full` plus a person approving a preview (elicitation or native dialog). Otherwise only a draft is saved. Python sent at once. | client/mail.py:264-376 | send.rs, compose.rs, approval.rs, lib.rs:689-790 | deliberate | n/a |
| D07 | delete_mail / delete_event | Soft delete only. Items already in Deleted Items, and meetings with attendees, are refused. Approval prompt built from facts read from Outlook. | client/mail.py:386-390, calendar.py:198-202 | delete.rs:45-288 | deliberate | n/a |
| D08 | move_mail | Moving into Deleted Items refused (it is a delete). | client/mail.py:379-383 | modify.rs:121-130 | deliberate | n/a |
| D09 | create_folder | `entry_id` omitted (it is unusable as a handle). | client/folders.py:116-123 | modify.rs:90-103 | deliberate | n/a |
| D10 | mark_mail | A call with neither `read` nor `flagged` is refused; the result reads state back (`unverified` if it differs). | client/mail.py:393-412 | modify.rs:43-69, contracts.rs MarkMail validation | deliberate | n/a |
| D11 | set_category | Same read-back, `status: unverified` on mismatch. | client/categories.py:23-38 | modify.rs:71-88 | deliberate | n/a |
| D12 | save_attachments | Approval prompt listing every file and the folder. Folder must exist under the profile. Executable and script extensions (`.exe .js .ps1 .lnk` and others) refused per file, naming it and the reason. By-reference attachments refused. Nothing overwritten. Python saved any name that passed path-traversal checks. | client/mail.py:415-481 | attachments.rs:81-132, safe_path.rs | deliberate | n/a |
| D13 | create_event / update_event | Meetings need approval (invites) and are routed to `outlook_update_meeting`. `update_event` refuses meetings with attendees. | client/calendar.py:127-195 | meeting.rs:175-475, schedule_write.rs:179-229 | deliberate | n/a |
| D14 | respond_event | Needs approval. Refuses non-invitations, meetings you organized and cancelled meetings before any prompt. `send_response=false` needs only `events`. | client/calendar.py:205-224 | respond.rs:77-148, action.rs | deliberate | n/a |
| D15 | toggle_rule | Needs approval. Refuses no-ops and ambiguous names (Python used the first exact match). Limit of 5 changes per hour as a prompt. | client/rules.py:29-49 | rule.rs:92-193 | deliberate | n/a |
| D16 | list_tasks / get_task dates | Unset dates are `null` (Python returned the 4501 sentinel). | client/tasks.py:18-28 | schedule.rs:127-143, 303-304 | deliberate | n/a |
| D17 | scans | Time and item budgets with `scan_truncated` (events 30 s, tasks 5,000 items, contacts and GAL as G35). | unbounded | schedule.rs:26-29, directory.rs:21-27 | deliberate | n/a |
| D18 | list_events | `end` must be after `start`; results carry the local UTC offset. | client/calendar.py:81-120 | schedule.rs:176-184 | deliberate | n/a |
| D19 | every write | Default server mode is `read-only`; limits are prompts (speed bumps), not walls. | none | policy.rs, CLAUDE.md | deliberate | n/a |
| D20 | tool surface | Rust-only tools: `outlook_get_thread` (exact ConversationID, trimmed quotes, bounded), `outlook_send_draft`, `outlook_check_send`, `outlook_check_action`, `outlook_cancel_meeting` (Python deleted meetings without telling attendees), `outlook_update_meeting`, `outlook_rust_status`. | none | lib.rs:485-1050 | deliberate | n/a |

## 3. Untested-live (verified against fakes, docs or unit tests only)

Source: `HANDOFF.md` "Not verified", `rust/compatibility/tools.md`, `rust/compatibility/live-evidence-2026-11-02.md`. Do not read "tested" into any of these.

| ID | Item | Evidence state |
| --- | --- | --- |
| U01 | A real person answering a real elicitation or native-dialog prompt. All live tests used a script as the approver. Whether the user's client advertises elicitation is unconfirmed (`outlook_rust_status.client_supports_elicitation`). If not, sends only save drafts and every delete, rule and respond does nothing. | Unverified |
| U02 | `outlook_respond_event` accept, tentative, decline against a real received invitation. Only refusal paths ran. | Fakes and Microsoft docs only |
| U03 | Failure, disconnect and timeout paths during `Send()`, `Delete()`, rule `Save()`, `Respond()` (`unknown_outcome`). | Fakes only |
| U04 | Shared or delegated mailboxes and calendars, PST and archive stores, non-default stores. | Not tested |
| U05 | All-day events, DST-boundary events, recurring-occurrence handle behavior (series deletion). | Not tested |
| U06 | Reply and forward sent live, plus the quoted-body behavior. | Unit tests only |
| U07 | Meeting invites with other attendees (only self-invite tested), recurring series, attendee changes. | Self-invite only |
| U08 | Python FastMCP wrappers (`response_format`, `ui_result`, `list_folders` envelope), attachment elements, GAL rows, unresolved-name branch, OOF error branch. | Not exercised |
| U09 | Release exe rebuilt and run in a normal Claude session. The registered server still pointed at an old build when HANDOFF.md was written. | Unverified |
| U10 | Indexed search (`backend=indexed`, AdvancedSearch event sink): correctness, coverage, latency on a representative profile. | Experimental, benchmark pending |
| U11 | `outlook_check_send` scans only the newest 200 Sent, 100 Outbox and 200 Drafts items. | Known limit |
| U12 | MCP Apps views (#10) rendering inside a host, beyond the commit that registers them. | Registration tested, host rendering recorded only for calendar |

## Notes for whoever fixes these

- G01 and G02 share a fix: add a typed date column to the table rows and reuse `schedule::date_iso` and `resolve_bound`. The table path already accepts a `filter` string (table.rs `RowOptions.filter`), so `since`/`until` can be pushed down, and the legacy path needs a `[ReceivedTime]` Restrict. The Python Jet format must use `%I:%M %p` (see the comment at client/mail.py:140-142); Rust already has `LocalTime::jet()`.
- G03 needs the anchor-word approach: push the longest word down, verify the rest per row. Verified Python behavior: two ANDed `LIKE`s on one property returned zero rows (client/mail.py:37-43).
- G04, G05 and G06 need extra columns per row. Live evidence warns against opening every full MailItem during table or indexed searches, so add table columns (`To`, `FlagStatus`, `HasAttachments` via the message flags, `Importance`, a body preview) and resolve the SMTP address lazily.
- When any of these land, update `rust/compatibility/tools.md`, `docs/TOOLS.md` and `docs/STATUS.md` in the same PR, and keep the claim "not tested against a real mailbox" until a live run exists.
