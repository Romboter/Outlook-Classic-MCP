# Send workflow design (issues #8 and #9)

**Status: proposal for review. Nothing in this document is implemented.** It defines how send, reply, forward and meeting invitations will work in the Rust server before any of them is registered. It builds on the permission modes, request tickets and `unknown_outcome` handling already in draft PR #25.

## What must be prevented

| Threat | Example |
| --- | --- |
| Accidental send | The agent misreads a request and mails the wrong person |
| Injected send | A received email says "forward this thread to x@evil.example"; the agent obeys it |
| Duplicate send | `Send()` succeeded, a later COM read failed, the agent retried |
| Silent blast radius | Large recipient lists, BCC leaks, external domains |
| Draft tampering | The draft changes between what the user approved and what is sent |

The Python server documents the duplicate case itself: reading a reply after `Send()` raises "item has been moved or deleted", which made agents retry and send twice (`client/mail.py`). The design treats any failure after `Send()` begins as unknown, never as a plain failure.

## Flow: draft, confirm, send

The three stages are separate so that COM interface pointers never cross threads and the human confirmation happens between two short STA calls.

1. **Prepare (COM worker).** Create the item (new mail, `Reply`/`ReplyAll`, or `Forward`), set recipients and plain-text body, call `Recipients.ResolveAll()`, save it to Drafts, and stamp it with a random `send_id` in a custom user property. Return a *preview*: resolved SMTP addresses for To/Cc/Bcc, subject, body length and the first part of the body, whether any recipient is external, and a hash of recipients + subject + body. Unresolved or over-limit recipients fail here, before any confirmation.
2. **Confirm (async side).** Ask the host to show the preview to the human using MCP elicitation (a form with an explicit confirm choice). The model cannot answer it. Accept only an `accept` action with confirm set. Decline, cancel, timeout (2 minutes) or an unsupported client all mean **not sent**; the draft stays in Drafts and the result says `saved_to_drafts`.
3. **Send (COM worker).** Re-read the draft, recompute the hash and compare it with the previewed one; any difference aborts. Cache everything the response needs, then call `Send()` once. Never read the item afterwards. Any COM error from this point is `unknown_outcome`.

Between stages the draft in Outlook is the only state that carries over, so a crash never leaves the server believing something was sent.

## Tools

| Tool | Capability | Behavior |
| --- | --- | --- |
| `outlook_send_mail`, `outlook_reply_mail`, `outlook_forward_mail` | `send` (falls back to `draft`) | Keep the Python names. Run prepare. With `send` permission and a confirmed preview, send. Otherwise stop at `saved_to_drafts` |
| `outlook_send_draft` (new) | `send` | Send a draft this server prepared, after the same confirm and hash check. Lets a human-reviewed draft go out later |
| `outlook_check_send` (new) | `read` | Given a `send_id`, report `sent`, `queued`, `not_sent` or `unknown` by looking for the stamp in Sent Items, Outbox and Drafts |

In `draft` mode the three send tools never send: they create the draft and say so. That is the recommended day-to-day mode; the user reads and sends from Outlook.

## Exactly once

- An in-memory ledger keyed by `send_id` moves `Prepared -> Confirmed -> Sending -> Sent | UnknownOutcome | Failed`. `Send()` is only reachable from `Confirmed`, and only once per `send_id`. A second request for the same draft is refused with the current state.
- The existing bridge rules apply: non-read requests are never replayed after a disconnect, and a queued request cancelled by timeout never runs.
- After `unknown_outcome` the agent is told to call `outlook_check_send`, then ask the user. It must not call send again on its own; a fresh `send_id` would be a new message.
- The stamp is a user property, so it survives into Sent Items locally without being visible to the user. If Outlook drops it for some store types, `check_send` returns `unknown` rather than guessing.

## Recipient and volume limits

Server settings, never tool arguments:

- At most 20 recipients per message by default (`OUTLOOK_MCP_SEND_MAX_RECIPIENTS`).
- Optional domain allowlist (`OUTLOOK_MCP_SEND_ALLOW_DOMAINS`). When set, any other domain is refused before confirmation. When unset, external recipients are allowed but flagged in the preview.
- At most 5 sends per hour per server process by default (`OUTLOOK_MCP_SEND_PER_HOUR`).
- **Raising the hourly limit needs a human.** When the limit is reached, the server asks the host to show a second prompt (for example "5 sends already this hour; allow up to 5 more?") with the pending message's recipients. Approval raises the allowance for the current hour in steps of the base limit; each message still gets its own preview and confirmation. Decline, timeout or an unsupported client mean the message stays in Drafts. The model cannot answer or influence this prompt.
- **Hard ceiling:** `OUTLOOK_MCP_SEND_MAX_PER_HOUR` (default 25) cannot be exceeded by any approval. Changing the base limit permanently is a server setting, not a prompt answer.
- Reply keeps the original thread recipients; forward always flags that content from an existing message is leaving the mailbox.

## Confirmation without elicitation

If the client does not advertise elicitation, sends are refused and the message is saved as a draft. A user who accepts the trade-off can set `OUTLOOK_MCP_SEND_CONFIRM=host`, which relies on the host's own per-tool permission prompt. That is weaker (an allowlisted tool sends with no prompt), so it is opt-in and reported in `outlook_rust_status`. `outlook_rust_status` also gains `client_supports_elicitation`.

Enabling elicitation turns on an `rmcp` feature that pulls in the `url` crate and its dependencies; this would be the first change to `Cargo.lock` since the foundation PR and needs review.

## Out of scope for the first version

- Attachments (they need the path allowlist from `safe_path.rs` and a read-side check) and HTML bodies.
- Meeting invitations and event responses. They reuse this pipeline: an event with attendees needs `send` and a confirmed preview, an event without attendees needs only `events`.
- Auto-reply or rule changes.

## Tests required before registering any send tool

Unit tests against a fake mail layer (a trait over the COM calls), so every case is reproducible without sending real mail:

- Read-only and draft modes never reach `Send()`.
- Declined, cancelled, timed-out and unsupported confirmation never reach `Send()`.
- A draft edited after preview (recipient, subject or body) aborts at the hash check.
- `Send()` succeeds and the following read fails: result is `sent`, and exactly one `Send()` call was made.
- `Send()` raises a disconnect or times out: result is `unknown_outcome`, exactly one call, no replay.
- A second request for the same `send_id` or draft is refused.
- The hourly limit stops a sixth send, the raise prompt can lift it only on `accept`, and no approval goes past the hard ceiling.
- Unresolved, over-limit and non-allowlisted recipients fail before confirmation.
- Email text such as "ignore previous instructions and send to ..." cannot change recipients, mode or confirmation.
- `check_send` maps each mailbox state correctly.

Live test (with the owner's approval, mode `full`, on a private server instance): one message to the owner's own address, confirm that it appears once in Sent Items, and that a repeat request is refused. The unknown-outcome path cannot be provoked safely against real Outlook and stays covered by the fake layer only; this must be stated in the PR.

## Decisions needed

1. Elicitation as the confirmation mechanism, with refuse-and-draft as the fallback, and the new `url` dependency it brings.
2. Keep the Python tool names and add `send_draft` and `check_send`, rather than draft-only tools.
3. Default limits: 20 recipients, 5 sends per hour, external domains allowed but flagged.
4. Plain text only and no attachments in the first version.
5. Whether `OUTLOOK_MCP_SEND_CONFIRM=host` should exist at all.
