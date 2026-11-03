# Confirmed actions design: delete, respond, invite, rules, attachments (issues #8 and #9)

**Status: approved by the owner (decisions below); implementation in progress, one tool per PR. Done: `outlook_delete_mail`, `outlook_delete_event`, `outlook_toggle_rule`, `outlook_respond_event` (accept path not yet exercised against a real invitation), `outlook_cancel_meeting` (run live against a self-invited meeting only), `outlook_check_action` (the engine is `rust/src/action.rs`, not a generalized `send.rs` as first sketched; counters are per class, so rule changes and deletes do not share limits). Invite (`outlook_create_event` with attendees) and update (a separate `outlook_update_meeting` tool, not `outlook_update_event`) are done and live-tested against a self-invite; recurring meetings and attendee changes are not supported. Not started: attachments. Cancellation shares the delete counter.** It covers the five Python tools still missing from the Rust server (`outlook_delete_mail`, `outlook_delete_event`, `outlook_respond_event`, `outlook_toggle_rule`, `outlook_save_attachments`) plus the two parts of calendar writes that were left out of draft PR #29: inviting attendees and changing meetings that have attendees. It reuses the pieces already built for sending in [SEND-DESIGN.md](SEND-DESIGN.md): server permission modes, MCP elicitation, the exactly-once ledger and `unknown_outcome` handling.

## The rule

**No mode, setting or argument lets one of these actions happen without a person approving it at the host.** `full` mode only makes the action *available*; each call still shows the person what will really happen and waits for approval. If the client cannot show the prompt, the action does not happen (unlike send, there is no meaningful draft to fall back to).

## What can go wrong

| Threat | Example |
| --- | --- |
| Injected action | A received email says "delete all messages from finance" or "accept this meeting" and the agent obeys |
| Wrong target | The agent names a handle that points at a different message than it described |
| Irreversible loss | A permanent delete, or removing a meeting without telling attendees |
| Unwanted notification | A response, invitation, update or cancellation emails other people |
| Silent automation | Re-enabling a rule that deletes or forwards mail |
| Hostile file | An attachment saved to disk with an executable extension or an escaping path |

## How each action behaves

| Tool | Capability | Prompt shows (read from Outlook, never from the model) | Behavior and limits |
| --- | --- | --- | --- |
| `outlook_delete_mail` | `delete` | Sender, subject, received time, folder | Moves to Deleted Items only (what `Delete()` does for an item outside Deleted Items; to verify live). **Refuses anything already in Deleted Items**, because that would be permanent. One item per call |
| `outlook_delete_event` | `delete` | Subject, time, whether recurring | Same soft delete. **Refused for meetings that have attendees**: deleting alone would leave the meeting on their calendars. Those need a cancellation, below. A recurring series is described as the whole series |
| `outlook_respond_event` | `send` | Meeting subject, time, organizer, the chosen response | `send_response=true` mails the organizer, so it is confirmed like a send. `send_response=false` records the answer locally without mail and still asks once |
| `outlook_create_event` with attendees | `send` | Subject, time, location, every attendee address, external flags | Same prepare, preview, hash recheck and single `Send()` as email: the meeting is saved on the user's own calendar first and invitations go out only on approval |
| `outlook_update_event` on a meeting | `send` | What changed, every attendee | Changes are applied and updates sent only on approval |
| cancel meeting (new `outlook_cancel_meeting`) | `send` | Subject, time, every attendee | Marks the meeting cancelled and sends cancellations on approval, then soft-deletes it |
| `outlook_toggle_rule` | `rules` | Rule name, current and new state, and what the rule does (deletes, moves, forwards, marks read) | A rule that deletes or forwards mail gets an extra warning line. Rules save immediately in Outlook, so there is no draft stage. Reads the state back |
| `outlook_save_attachments` | `files` | Message sender and subject, each file name and size, the target folder | Folder must resolve inside the user profile (`safe_path`), names are validated, nothing is overwritten. Executable and script extensions (`.exe .dll .js .vbs .ps1 .bat .cmd .lnk .scr .msi` and similar) are refused outright, not prompted; the refusal names the file and says why |

Python behavior that is deliberately dropped: deleting meetings that have attendees without telling them, deleting from Deleted Items, saving executables, and replying "accept" without anyone seeing it.

## Flow

The same three stages as send, so COM pointers never cross threads and the person's decision sits between two short Outlook calls:

1. **Prepare (COM worker).** Identify the target by EntryID. Read the facts for the prompt, check the target is allowed (not in Deleted Items, no attendees for a plain delete, extension allowed), and compute a hash of the facts plus the item's identity. Stamp it with a random `action_id` user property where the action leaves an item behind to find later.
2. **Confirm (async side).** MCP elicitation form with the facts and one tick box. Decline, cancel, timeout or an unsupported client means **nothing happens** and the result says so.
3. **Execute (COM worker).** Re-read the item and recheck the hash. Any difference aborts as `not_done`. Otherwise perform the one action, never touching the item afterwards (an item that moved cannot be read reliably), and verify by looking at the result where Outlook allows.

## Exactly once and reconciliation

- The existing ledger covers these too: an `action_id` moves from prepared to acting to done or unknown, and cannot act twice. Non-read requests are never replayed after a lost connection.
- Any failure once the action has begun is `unknown_outcome`, never a plain failure.
- `outlook_check_send` becomes `outlook_check_action` (the old name stays as an alias). Given an `action_id` it reports `done`, `not_done` or `unknown`: a delete is found by the stamp among the newest Deleted Items; a response or invitation among the newest Sent Items; a rule change by reading the rule. The agent must call it and ask the user before retrying anything.
- A delete changes the EntryID, so the original handle cannot prove what happened. That is why deletes are stamped first.

## Limits are speed bumps

Same principle as send: crossing a threshold adds a section to the prompt, never a flat refusal, with hard ceilings behind it.

| Threshold | Default | Hard ceiling |
| --- | --- | --- |
| Deletes per hour | 10 | 50 |
| Meeting responses per hour | 10 | 50 |
| Attendees per invitation or cancellation | 20 | 200 (the send recipient limits) |
| Rule changes per hour | 5 | 20 |

Settings come from the server environment (`OUTLOOK_MCP_DELETE_PER_HOUR` and so on, with `_HARD_MAX_` variants and `OUTLOOK_MCP_ACTION_STRICT=1` to refuse instead of prompt). Counters are in memory; a restart resets them.

## Permission modes

Nothing changes for `read-only` and `draft`: all of these tools are refused. In `full` the capabilities `delete`, `send`, `rules` and `files` can each be denied with `OUTLOOK_MCP_DENY`. Having the capability never skips the prompt.

## Implementation plan

1. Generalize the send engine (`send.rs`) so a "prepared action" carries a kind, a list of facts for the prompt, a hash and an optional stamp. Send stays one kind; the rest are new kinds. This keeps one ledger, one confirmer and one set of tests.
2. COM side: `prepare`/`execute`/`locate` per kind, in new modules beside `compose.rs`.
3. Tools in this order, each its own PR: delete mail and delete event (soft only), toggle rule, respond event, meetings (invite, update, cancel), save attachments.
4. Each PR carries the test matrix below and a live test on items the user labels as test items.

## Tests required before each tool is registered

Fake-based unit tests, no real mail or deletion:

- `read-only` and `draft` modes never reach the action.
- Declined, cancelled, timed-out and unsupported confirmation do nothing.
- A target that changed between preview and execute aborts at the hash check.
- A failure after the action starts is `unknown_outcome`, exactly one attempt, no replay; a second request for the same `action_id` is refused.
- Refusals before any prompt: item already in Deleted Items, meeting with attendees for a plain delete, executable attachment extension, path outside the profile.
- Threshold prompts and hard ceilings; strict mode refuses.
- The prompt shows facts read from Outlook, not text supplied by the model; email text such as "delete everything" cannot change the target, the mode or the confirmation.
- `outlook_check_action` maps each state.

Live tests use only items the owner has designated as test items (a throwaway message, event or rule), with the owner's approval and a script standing in for the person, and the PR must say so. The unknown-outcome path stays fake-only.

## Open facts to verify live before relying on them

- That `Delete()` on a mail or appointment outside Deleted Items always moves it there, for the owner's account and store type.
- That `Respond(..., True)` returns an item that does nothing until `Send()` is called.
- That the stamp user property survives into Deleted Items and Sent Items for each item type (it did for sent mail).
- Whether Outlook shows its own security or confirmation dialogs for any of these over COM, which would block the worker.

## Decisions (approved by the owner)

1. **Approved.** No action without a person's approval, ever, in any mode; an unsupported client means the action does not happen.
2. **Approved.** Soft delete only; Deleted Items and meetings with attendees are refused; permanent deletion is never offered.
3. **Approved.** New `outlook_cancel_meeting`; `delete_event` refuses meetings with attendees.
4. **Approved with an amendment.** Executable and script attachment extensions are refused outright, **and the refusal says which file was refused and why** (for example: "invoice.exe was not saved: executable and script files can run code when opened"), so the user can decide what to do themselves. Nothing is saved for a refused file.
5. **Approved on one condition:** every threshold stays a speed bump (a prompt), never a wall, with hard ceilings only as runaway protection: 10 deletes, 10 responses and 5 rule changes per hour (hard 50, 50, 20).
6. **Approved.** Build order: delete, rules, respond, meetings, attachments.
