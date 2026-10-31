# Outlook Classic MCP — Rust implementation

> **Status: experimental read-only MCP implementation.** Five COM-backed mailbox tools are now exposed, but they still require live Outlook Classic testing on Windows. The existing Python implementation is unchanged.

This directory will contain a Windows-native rewrite optimized for **fast, low-token searches** against an already-running Outlook Classic profile. The goal is local stdio MCP, indexed Outlook search, concise search-result metadata, and full-message reads only when specifically requested.

## Current behavior

The Rust executable exposes `outlook_rust_status` and five read-only MCP mailbox tools. On a Windows workstation with Outlook Classic running, you can also deliberately invoke the CLI probes:

```powershell
cargo run -- --probe-outlook
cargo run -- --probe-first-mail
```

The first probe reads the bound profile's display name and Inbox counts. The second additionally reads only the first Inbox item's subject for local diagnostics. Both return a compact JSON object **to your terminal**, not to the MCP client. Avoid sharing that output publicly. The probe attaches through the COM Running Object Table and does not auto-launch Outlook. No sends, writes, moves, deletes, or direct OST parsing occur.

**Important:** Windows CI cannot verify actual Outlook behavior because the hosted runner has no interactive Outlook profile. Live Windows/Outlook integration testing remains required for issue #3; do not claim this is production ready.

## Experimental read-only MCP tools (issue #4)

The Rust stdio MCP now registers **five Outlook Classic tools**, in addition to the diagnostic `outlook_rust_status`. This is an **experimental implementation that has not been validated against an actual Outlook Classic profile**. Do not replace the existing Python MCP configuration in production yet.

| Tool | Behavior |
| --- | --- |
| `outlook_whoami` | Current Outlook display name |
| `outlook_list_folders` | Default mailbox root (or a supplied folder), maximum depth 4 / 100 entries |
| `outlook_list_mails` | One folder, sorted latest first, metadata only; 10 items by default / max 30 |
| `outlook_search_mails` | Outlook `Items.Restrict` baseline using escaped DASL; indexed search is issue #5 |
| `outlook_get_mail` | Read one message by Outlook EntryID (optionally StoreID), with bounded plain-text body |

The MCP server itself **cannot send, move, edit, or delete email**. The tools return minimal JSON strings. Search/list results intentionally do not include email bodies. `get_mail` defaults to 4,000 characters and is capped at 12,000. HTML body retrieval is disabled. Invalid or unsupported arguments return explicit errors.

Known limitations before replacing Python:
- Outlook COM calls have **not** been tested on a real Outlook Classic installation.
- The current search uses `Items.Restrict`, not indexed `AdvancedSearch`; benchmarking and faster indexed search belong to #5.
- The `since` and `until` parameters deliberately return unsupported errors rather than ignoring filters.
- `ReceivedTime` is returned as **Windows-locale `received_local` text**, not an ISO timestamp. Do not interpret it as UTC.
- Folder resolution currently supports default-store folders, not fully qualified shared/delegated stores.
- Long `EntryID` values are currently included in results; shorter agent-facing IDs are tracked in #6.
- A timeout is not a cancellation: a queued **read** may finish later. No writes are exposed.

For a protocol smoke test, run with an MCP client and call `outlook_rust_status` without Outlook. The actual mailbox tools require the running interactive Outlook Classic process and the Windows user profile hosting the MCP process.

## Build

Requires Rust **1.88+** (edition 2024) on Windows for eventual COM integration.

```powershell
cd rust
cargo build
cargo test
cargo fmt --all -- --check
cargo clippy --all-targets -- -D warnings
cargo run -- --version
```

Run the MCP stdio server:

```powershell
cargo run
```

To inspect the protocol, use an MCP client or the MCP Inspector. **Do not replace the working Python MCP in production until real Outlook tests pass.**

### Dependency and lockfile policy

The MCP SDK is pinned at `rmcp =3.5.1`. The checked-in `Cargo.lock` pins transitive dependencies; CI uses `--locked` for builds, linting, and tests. Update the lockfile intentionally via `cargo update` or `cargo generate-lockfile` on a review branch, and include it in the same PR as dependency changes. Live Outlook COM behavior is not tested by hosted CI.

## Design constraints

- **Fork-owned development only:** work in `Romboter/Outlook-Classic-MCP`, through feature branches and PRs targeted to that repository's `main`. Do not sync with or PR to the original author's repository.
- **No OST parser or second mail index in v1:** leverage Outlook Classic's indexed search first, benchmarking `AdvancedSearch` + `Search.GetTable` and `Folder.GetTable`.
- **Single STA COM worker:** Outlook interfaces stay on their dedicated Windows apartment thread; tool handlers communicate using messages, never cross-thread COM interface sharing.
- **Read-only first:** avoid exposing send/move/delete tools until permission gating and operation-aware error recovery are verified.
- **Context budgets:** default search to at most 10 metadata hits; use opaque short-lived identifiers internally mapped to both EntryID and StoreID, then selectively fetch bodies.
- **Data safety:** no email bodies, recipients, or file paths in routine logs. MCP stdout contains protocol messages only. Email bodies are untrusted input.
- **Interop:** preserve all 31 existing Python tool schemas where possible. The contract inventory is in `compatibility/tools.md`.

## Milestones

- [#2](https://github.com/Romboter/Outlook-Classic-MCP/issues/2) Foundation, dependencies, contracts, CI.
- [#3](https://github.com/Romboter/Outlook-Classic-MCP/issues/3) Windows Outlook COM bridge and STA recovery.
- [#4](https://github.com/Romboter/Outlook-Classic-MCP/issues/4) Five functioning read-only tools.
- [#5](https://github.com/Romboter/Outlook-Classic-MCP/issues/5) Indexed Outlook search benchmarking and implementation.
- [#6](https://github.com/Romboter/Outlook-Classic-MCP/issues/6) Token-minimal result design.

## Layout

```text
rust/
  Cargo.toml
  src/
    main.rs     # stdio entry point
    lib.rs      # Read-only MCP tool router
    contracts.rs # Schema + parameter validation
    outlook.rs  # STA COM worker and explicit CLI probe
    outlook/mailbox.rs # Bounded read-only COM operations
  compatibility/
    tools.md    # existing Python tool inventory
```

The Python server stays at `src/outlook_mcp/` and is unaffected. Any upstream MIT-covered material reused in the rewrite must keep the required license and attribution.
