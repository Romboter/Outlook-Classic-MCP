# Outlook Classic MCP — Rust implementation

> **Status: experimental COM probe.** The Rust MCP server still has **no mail-reading tools**. A separate, explicit Windows CLI command now attempts to attach to the already-running Outlook Classic process. The original Python implementation remains unchanged.

This directory will contain a Windows-native rewrite optimized for **fast, low-token searches** against an already-running Outlook Classic profile. The goal is local stdio MCP, indexed Outlook search, concise search-result metadata, and full-message reads only when specifically requested.

## Current behavior

The Rust MCP executable exposes only `outlook_rust_status`; it **does not** read or return messages to an LLM. On a Windows workstation with Outlook Classic running, you can deliberately invoke the CLI probes:

```powershell
cargo run -- --probe-outlook
cargo run -- --probe-first-mail
```

The first probe reads the bound profile's display name and Inbox counts. The second additionally reads only the first Inbox item's subject for local diagnostics. Both return a compact JSON object **to your terminal**, not to the MCP client. Avoid sharing that output publicly. The probe attaches through the COM Running Object Table and does not auto-launch Outlook. No sends, writes, moves, deletes, or direct OST parsing occur.

**Important:** Windows CI cannot verify actual Outlook behavior because the hosted runner has no interactive Outlook profile. Live Windows/Outlook integration testing remains required for issue #3; do not claim this is production ready.

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
    lib.rs      # MCP tool router (diagnostic only)
    outlook.rs  # STA COM worker, explicit CLI probe
  compatibility/
    tools.md    # existing Python tool inventory
```

The Python server stays at `src/outlook_mcp/` and is unaffected. Any upstream MIT-covered material reused in the rewrite must keep the required license and attribution.
