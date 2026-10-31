# Outlook Classic MCP — Rust implementation

> **Status: foundation only.** This is a standalone Rust MCP **scaffold**, not an Outlook reader yet. The original Python implementation remains fully functional and unchanged.

This directory will contain a Windows-native rewrite optimized for **fast, low-token searches** against an already-running Outlook Classic profile. The goal is local stdio MCP, indexed Outlook search, concise search-result metadata, and full-message reads only when specifically requested.

## Current behavior

The Rust executable exposes only `outlook_rust_status`. It reports that the COM bridge and mailbox tools are **not implemented**; it does **not** read messages or authenticate to Outlook.

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

To inspect the protocol, use an MCP client or the MCP Inspector. **Do not register the scaffold in place of the working Python server.**

### Dependency and lockfile policy

The MCP SDK is pinned at `rmcp =3.5.1`. The checked-in `Cargo.lock` pins transitive dependencies; CI uses `--locked` for builds, linting, and tests. Update the lockfile intentionally via `cargo update` or `cargo generate-lockfile` on a review branch, and include it in the same PR as dependency changes. Live Outlook COM behavior is not tested by this scaffold.

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
    lib.rs      # MCP tool router (foundation only)
  compatibility/
    tools.md    # existing Python tool inventory
```

The Python server stays at `src/outlook_mcp/` and is unaffected. Any upstream MIT-covered material reused in the rewrite must keep the required license and attribution.
