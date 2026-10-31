# Runbook 03 — Validate Outlook Classic COM attachment

**Use when:** Testing the Rust STA worker on an actual Windows machine with Outlook Classic; especially after COM/reconnect changes.

**Risk:** This is read-only, but the probe reports active profile data locally. Use a trusted terminal and do not paste results into public logs. Do not run against an unattended service session.

## Preconditions

1. Build the latest experimental Rust branch (see [runbook 02](02-build-test.md)).
2. Sign in to Outlook **Classic** manually. Wait for Inbox and folders to load.
3. Ensure the Rust terminal and Outlook are in the **same interactive user's session**.
4. Close any account setup, security prompt or modal Outlook dialog. Confirm you are not using New Outlook (`olk.exe`).

## Procedure

1. Verify process without dumping sensitive state:
   ```powershell
   Get-Process OUTLOOK -ErrorAction SilentlyContinue | Select-Object Id,SessionId
   ```
2. Execute the **safe account/count probe** from the repo root:
   ```powershell
   .\rust\target\debug\outlook-mcp-rs.exe --probe-outlook
   ```
   This prints JSON including `current_user`, Inbox name, `inbox_items` and `inbox_subfolders`. View locally; do **not** share raw output.
3. Confirm `current_user` is the expected Outlook profile, and counts are plausible. Record only pass/fail and anonymized counts if permitted.
4. Optional, only with explicit approval and an appropriate test profile:
   ```powershell
   .\rust\target\debug\outlook-mcp-rs.exe --probe-first-mail
   ```
   **Caution:** This includes a subject line. Never paste it into GitHub or chat unless explicitly authorized.
5. In an MCP Inspector session, invoke `outlook_whoami` and `outlook_list_folders` as described in [runbook 04](04-mcp-smoke-test.md).
6. To test reconnect behavior, restart **the Rust MCP process first** and repeat. For a separate supervised test, close/restart Outlook and repeat after the profile is ready; expect session handles to reset.

## Success criteria

- Probe attaches without auto-launching Outlook; the expected profile is queried.
- `outlook_whoami` and folder listing return real, plausible values with no COM errors.
- Rust restart, and separately Outlook restart, are handled without a crash or false claim of connection.
- No mail writes, SMTP sends or attachment saves occur.

## Failure handling

- "Outlook Classic is not running / ROT": verify process and user/session; do **not** invoke COM from SYSTEM or a detached service.
- Timed out: check for modal dialogs or busy Outlook and stop running additional requests; a 45-second caller timeout does **not** abort the COM call.
- COM RPC disconnected or unavailable: restart the Rust MCP process; check Outlook state before repeating. Avoid blind repeated calls.
- Wrong user/profile: stop testing immediately; close the Rust process and correct the Windows/Outlook session.

## Evidence

Use [validation report](../templates/validation-report.md): OS/Outlook build, Rust commit, safe pass/fail, failure category and whether reconnect worked. **No profile names, sender data, raw JSON or message identifiers.**
