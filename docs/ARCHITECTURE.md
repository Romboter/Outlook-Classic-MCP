# Architecture

## Runtime path

```text
Claude Code / MCP client
    | local JSON-RPC messages on stdin/stdout
    v
rust/src/main.rs -> rmcp server -> rust/src/lib.rs
    | validated request from rust/src/contracts.rs
    | tokio::spawn_blocking + request/reply channel
    v
rust/src/outlook.rs — one dedicated STA worker
    | CoInitializeEx(STA), idle message pump
    | GetActiveObject("Outlook.Application")
    v
Outlook Classic COM / MAPI, running in user's session
    | reads mailbox through existing Outlook profile
    v
COM mail / Table / AdvancedSearch / Conversation APIs
    | metadata, dates, bodies only when explicitly requested
    v
bounded JSON result -> short-ID sanitization -> MCP tool response
```

**Data trust boundaries:** MCP requests are untrusted input; `contracts.rs` validates ranges and unsupported flags. Outlook email subjects and bodies are **untrusted retrieved content** and must not be treated as developer/system instructions. COM interfaces remain STA-local. Raw EntryID and StoreID should remain inside the worker's in-memory cache; returned messages use short session handles.

## Transport and lifetime

- The executable is `rust/target/debug/outlook-mcp-rs.exe` for local builds; release path is `rust/target/release/outlook-mcp-rs.exe` **only after building release**.
- No arguments starts MCP stdio. Standard output must contain **only** MCP JSON-RPC. Diagnostics go to stderr.
- `--help`, `--version`, and `--probe-outlook` are separate CLI invocations, **not** MCP mode.
- The worker attaches to already-running Outlook using the COM Running Object Table. It deliberately does **not** start Outlook.
- Requests have an approximately **45-second caller wait**; a timeout is *not* COM cancellation. Async index search uses an internal **18-second** event timeout and best-effort `Stop`.
- On recognized disconnect errors the worker attempts one read-only reconnect. Session-scoped short IDs may be invalid after reconnect or restart.

## Read and search backends

| Operation | Code | Notes |
| --- | --- | --- |
| List folders | `mailbox.rs` | Depth capped at 4; result capped at 100 folders |
| List mail / basic search | `mailbox.rs`, `table.rs` | `auto`, `table`, `legacy` modes; returns metadata only |
| Index-optimized search | `indexed.rs` | Explicit `backend:"indexed"` only; requires `IsInstantSearchEnabled`, waits for `AdvancedSearchComplete`, then `Search.GetTable` |
| Get one email | `mailbox.rs` | Full plain-text body only by explicit request; HTML disabled |
| Get thread | `thread.rs`, `thread_text.rs` | Exact ConversationID, bounded messages and optional conservative quote truncation |
| Session handle cache | `refs.rs` | In-memory maximum 2,048 IDs; mapping includes store identity |

The default search path has **not** been proven faster on real Outlook. The `auto` backend does not use `indexed` until validation demonstrates correct and reliable behavior. `Folder.GetTable` is a lightweight rowset, **not by itself proof of indexed search**. Details: [benchmark guide](../rust/docs/search-benchmark.md).

## Failure model

- Missing or unsupported Outlook Classic: explicit error rather than simulated mailbox output.
- Store cannot satisfy Table columns: automatic list/subject/sender search can fall back to legacy; explicit `table` mode reports an error. `indexed` never silently falls back.
- Missing/stale short handle: repeat search/list to populate session handle cache.
- Missing/deleted or inaccessible messages in a conversation: count and report omissions; do not synthesize correspondence.
- Long-running COM operation: caller may time out while the STA remains blocked. No blind retry of future mutating operations.

## Why not parse OST?

Outlook itself owns mailbox/profile identity, caching, providers and search. An independent OST parser or secondary index would add complexity, synchronization, security and data-retention risks. Current design uses the existing Outlook profile and index where it proves correct; we will only reconsider with benchmark evidence and a separate decision.

See [security](SECURITY.md), [tool contracts](TOOLS.md), and [architectural risks](DECISIONS.md).
