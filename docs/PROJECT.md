# Project overview

## Why this fork exists

**Problem:** Claude Code should be able to find and selectively read mail from an already-running **Outlook Classic** on Windows with low latency and a small context footprint. The user's Outlook profile caches roughly one year of mail locally. The working Python server uses `pywin32`; the fork-specific rewrite evaluates a native Rust MCP server and Outlook's existing COM/search capabilities.

**Success looks like:** Selective, correct email search/read/thread retrieval with bounded metadata, small server-side message handles, safe read-only defaults and measured performance on a real Outlook Classic profile. Expand toward the Python tool surface only after safety and real-world testing.

## In scope

- Windows Outlook **Classic**, using the interactive user's existing authenticated COM session.
- Local MCP over **stdio**, not a remotely exposed web service.
- One dedicated **STA COM worker**; the async MCP handlers send requests, not COM objects.
- Inbox/folder discovery, compact mail search/list/read, conversation retrieval.
- `Items.Restrict`, `Folder.GetTable`, and opt-in `Application.AdvancedSearch`/completion events, with honest backend reporting.
- Windows build CI and reproducible Rust dependencies.
- Later: permission-gated writes, parity with the Python server, MCP Apps UI and packaging.

## Out of scope for the current prototype

- New Outlook (`olk.exe`): no compatible Outlook Classic COM surface.
- Raw OST decoding or a duplicate full-text database/index.
- Microsoft Graph / OAuth / cloud mailbox registration for baseline reads.
- Background Windows service or headless session.
- Production Rust deployment, mutating tools, calendar/contact/task/rule parity, a signed release.
- Any upstream synchronization. This is an independent development line within the user's fork.

## Code layout

```text
src/outlook_mcp/              Existing Python implementation — keep working
rust/Cargo.toml               Rust crate (edition 2024; Rust >= 1.88)
rust/Cargo.lock               Pinned transitive dependencies
rust/src/main.rs              stdio MCP entrypoint and explicit local probes
rust/src/lib.rs               MCP router and tool discovery
rust/src/contracts.rs         Parameters and server-side validation
rust/src/outlook.rs           Dedicated COM STA and request channel
rust/src/outlook/mailbox.rs   Mail/folder read operations
rust/src/outlook/table.rs     Lightweight Table row access
rust/src/outlook/indexed.rs   Opt-in AdvancedSearch + completion handling
rust/src/outlook/refs.rs      Bounded session-scoped short message IDs
rust/src/outlook/thread.rs    ConversationID/Conversation.GetTable retrieval
rust/src/thread_text.rs       Conservative quoted-text trimming and budgets
rust/docs/                    Detailed bench and thread testing guides
docs/                         Project, operational runbooks and handoff docs
.github/workflows/rust-ci.yml Windows fmt, Clippy, tests, build, CLI smoke
```

## Backlog and branches

[Roadmap issue #1](https://github.com/Romboter/Outlook-Classic-MCP/issues/1) is the work ledger. Issues [#2–#7](https://github.com/Romboter/Outlook-Classic-MCP/issues) have experimental implementations in seven stacked draft PRs:

| Draft PR | Topic |
| --- | --- |
| [#12](https://github.com/Romboter/Outlook-Classic-MCP/pull/12) | Rust foundation and CI |
| [#13](https://github.com/Romboter/Outlook-Classic-MCP/pull/13) | STA COM bridge |
| [#14](https://github.com/Romboter/Outlook-Classic-MCP/pull/14) | Five read-only MCP tools |
| [#15](https://github.com/Romboter/Outlook-Classic-MCP/pull/15) | Folder.GetTable and benchmarking modes |
| [#16](https://github.com/Romboter/Outlook-Classic-MCP/pull/16) | Short session-scoped mail handles |
| [#17](https://github.com/Romboter/Outlook-Classic-MCP/pull/17) | Opt-in asynchronous indexed search |
| [#18](https://github.com/Romboter/Outlook-Classic-MCP/pull/18) | Conversation/thread retrieval |

Issues [#8](https://github.com/Romboter/Outlook-Classic-MCP/issues/8) (write-operation safety), [#9](https://github.com/Romboter/Outlook-Classic-MCP/issues/9) (full parity), [#10](https://github.com/Romboter/Outlook-Classic-MCP/issues/10) (MCP Apps), and [#11](https://github.com/Romboter/Outlook-Classic-MCP/issues/11) (release) are incomplete. **The 31-tool inventory is not 31 implemented Rust tools.**

## Acceptance gates

1. CI builds/tests pass (already achieved for the experimental branches as of this baseline).
2. Live Outlook Classic under the same interactive Windows account: COM attachment, MCP reads and reconnects verified.
3. Search correctness and performance benchmarked on the cached mailbox; opt-in indexed events and their timeout behavior verified.
4. Conversation completeness/quote trimming compared with originals.
5. Security/privacy review; only then decide on broader permissions or release packaging.

For exact status at a later date, inspect the linked issues, PR heads and CI results. This document is a dated snapshot, **not proof of a working production deployment**.
