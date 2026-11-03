# Native dialog approval design (issues #8 and #9)

**Status: decisions approved by the owner (below). Implemented in a draft PR stacked on the spike: `OUTLOOK_MCP_APPROVAL` (default `dialog`), `approval.rs`, `dialog.rs` (TaskDialogIndirect, manifest via `build.rs`). Verified: unit tests for backend choice and prompt splitting, a timeout run and a manual Approve/Decline run of the spike example. NOT verified: the window opened by a client-launched server (foreground), a real tool call approved through it, DPI, keyboard-only use.**

## Why

Every write tool asks a person to approve through MCP elicitation (`ElicitConfirmer` in `rust/src/lib.rs`). On the client this was checked against (the Claude Desktop Code tab, 2026-11-02) `outlook_rust_status` reported `client_supports_elicitation: false`. Public reports say Claude Desktop and Cowork either do not support elicitation or declare it and never show the prompt (anthropics/claude-code issues 41110, 88901, 56243; the Claude Code CLI is the only client reported to work, and VS Code reportedly auto-declines). The intended users will not have the CLI. With elicitation only, every send would just save a draft and every delete, rule change, response, meeting change and attachment save would do nothing.

The rule in [CONFIRMED-ACTIONS-DESIGN.md](CONFIRMED-ACTIONS-DESIGN.md) does not change: **no action without a person approving it at the host, in any mode.** This note only adds a second way to show the prompt.

## Idea

The server runs on the same interactive Windows desktop as Outlook Classic. It can show its own modal dialog instead of asking the client to. The model cannot click that dialog, and the text comes from the server's facts read from Outlook, as the elicitation prompt does today.

## Fit with the code

`Confirmer` (`rust/src/send.rs`) is already the seam: `confirm(&self, prompt: &str) -> Confirmation` returning `Accepted`, `Declined` or `Unsupported`. Both the send engine and the confirmed-action engine (`action.rs`) use it, and their prompts are already built from Outlook facts. So the change is a new confirmer plus a chooser, not a change to the engines, the hash recheck, the ledger or `unknown_outcome`.

- `DialogConfirmer`: shows the prompt in a modal Win32 message box (Yes/No, default button No) and maps Yes to `Accepted`, everything else to `Declined`.
- `ApprovalConfirmer`: picks a backend per call (below).
- No new dependency: the `windows` crate in `Cargo.toml` already enables `Win32_UI_WindowsAndMessaging`, which has `MessageBoxW`.

## Backend selection

`OUTLOOK_MCP_APPROVAL` (server environment only, like the permission mode):

| Value | Behavior |
| --- | --- |
| `auto` (proposed default) | Elicitation if the client advertises form elicitation, else the dialog |
| `dialog` | Always the dialog; never elicit |
| `elicitation` | Elicitation only; unsupported client means nothing happens (today's behavior) |
| `none` | No approval backend: write tools do nothing, sends save drafts |

`auto` cannot detect a client that declares elicitation but never renders it (reported for Desktop/Cowork). Today such a call would wait out the 120 s timeout and then be treated as `Declined`. Users on a client like that need `dialog`; see decision 1.

## Dialog behavior and safety

- **Never auto-approve.** Any failure to show the dialog is `Declined`: no interactive desktop, `MessageBoxW` returns 0, window station not interactive, or any error. A headless server therefore does nothing.
- **Timeout is a denial.** A modal message box blocks its thread, so it runs on a dedicated thread (never the COM STA worker, never a Tokio worker) with the same 120 s deadline as elicitation. At the deadline the dialog is closed and the result is `Declined`. How to close it reliably (find the window and post `WM_CLOSE`, or a custom dialog with a timer) must be shown to work before this ships.
- **Foreground.** The box is created topmost with `MB_SETFOREGROUND` and `MB_SYSTEMMODAL`-style flags so it is not hidden behind Outlook or the client. Whether it reliably appears in front when the MCP server is a background child process is **unverified** and the main live-test question.
- **Text.** The prompt is the same server-built text. Dialog text is plain, length-bounded, and shows no body content beyond what the prompt already shows. Default button is No. The title names the server, not the client, so the person can tell where it came from.
- **Prompt is not a secret channel.** It must never be able to include text the model could use to trick the person; as today, every fact comes from Outlook, not from tool arguments.
- **One at a time.** A mutex so two concurrent write calls cannot stack dialogs the person might answer in the wrong order. The second waits, then asks.
- **Logging.** Stderr only, no mailbox content, only "approval requested/accepted/declined/timed out" and the backend used.

## Limits of this approach

- It proves a person clicked, not that they read it. Same as elicitation.
- The person must be at the machine. Remote or locked sessions: the dialog is not seen and times out as `Declined`.
- Windows only. The crate is already Windows-only for COM.
- Remote-control or computer-use agents that can click on this desktop could press the button. Out of scope here; note it in [SECURITY.md](SECURITY.md) if this ships.
- A `PreToolUse`-style client permission prompt would also work in Claude Code but is outside the server's control and shows only tool arguments, so it is not part of this design.

## Tests required

Fake-based unit tests (no real dialog):

- Chooser: each `OUTLOOK_MCP_APPROVAL` value picks the expected backend; `auto` falls back to the dialog only when elicitation is `Unsupported`.
- Dialog failure, timeout and closed-without-answer all map to `Declined`.
- `none` and `elicitation` with an unsupported client do nothing, exactly as today.
- Concurrent requests do not overlap dialogs.
- Existing engine tests still pass unchanged (they use fake confirmers).

Live (needs the owner at the machine and their explicit go-ahead; test items labeled `[outlook-mcp test]`):

- A real dialog appears in front of Outlook and the client, from a server started by the client.
- Yes performs exactly one action; No, close and timeout do nothing.
- Behavior with the screen locked or the dialog hidden.

## Decisions (owner, 2026-11-02)

1. **Default is `dialog`**, not `auto`. The target users are on clients where elicitation is missing or unreliable, and `auto` cannot detect a client that declares it but never shows it. `auto`, `elicitation` and `none` stay available as overrides. The table above is updated in intent: read "proposed default" as `dialog`.
2. **A richer window, not a plain message box**: scrollable text, the full attendee or file list, a visible countdown, and Approve / Decline buttons with Decline as the default.
3. **Keep the 120 s timeout.**
4. **The dialog covers every confirmed action** (send, reply, forward, delete, rules, respond, meetings, attachments), not just sends and meeting mail.

## Consequences of the richer window (to settle before coding)

The message box in this note would be replaced by one of:

- **`TaskDialogIndirect`**: expandable details, custom buttons, a callback timer for the countdown. Needs Common Controls v6, which a console exe only gets through an embedded application manifest. It is unverified that the build can embed one without a new build dependency.
- **A custom Win32 window** (own window class and message loop on the dedicated thread): full layout control and no manifest need, but more code to get right (focus, DPI scaling, keyboard handling, closing at the deadline).

Prefer `TaskDialogIndirect` if the manifest works, else the custom window. Safety rules above (fail closed, no auto-approve, dedicated thread, one dialog at a time, default No/Decline, text from Outlook facts only) apply to either. The richer window means more UI code that cannot be unit-tested, so the live test list grows: scrolling a long attendee list, countdown at the deadline, keyboard-only use, high-DPI, and the window appearing in front.
