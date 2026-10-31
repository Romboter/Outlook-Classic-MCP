# Outlook Classic MCP — development instructions

## Repository ownership and branch policy

This is the **Romboter fork** and a **fork-specific Rust rewrite**. Never open a pull request against `anasahmed07/Outlook-Classic-MCP`; never sync, pull, fetch, rebase, or cherry-pick from that upstream repository as part of this rewrite. All feature branches and pull requests must be entirely within `Romboter/Outlook-Classic-MCP` and target this fork's `main`.

Do not commit directly to `main`. Create a branch for each issue or coherent change. Open a PR for review. Do not merge without explicit user direction.

## Work order and scope

Read [#1](https://github.com/Romboter/Outlook-Classic-MCP/issues/1), then the individual issues:
- Foundation [#2](https://github.com/Romboter/Outlook-Classic-MCP/issues/2)
- COM [#3](https://github.com/Romboter/Outlook-Classic-MCP/issues/3)
- Read-only MCP [#4](https://github.com/Romboter/Outlook-Classic-MCP/issues/4)
- Indexed search [#5](https://github.com/Romboter/Outlook-Classic-MCP/issues/5)
- Token-efficiency [#6](https://github.com/Romboter/Outlook-Classic-MCP/issues/6)

The primary user's needs: Outlook **Classic** is always running at work and caches about one year of mail. Claude Code should search existing indexed mail quickly without filling its context. No direct OST parsing or separate Tantivy/SQLite index for the initial implementation.

## Existing implementation

The working Python server is in `src/outlook_mcp/`. Keep it working during Rust development. The experimental Rust crate lives in `rust/`. Do not redirect existing plugin/client configurations until Rust has passed real Outlook Classic integration tests. No false-positive tools that claim to read mail before the COM adapter exists.

Preserve 31-tool compatibility where possible; see `rust/compatibility/tools.md`. Rust foundation intentionally exposes only `outlook_rust_status`.

## Technical / security guardrails

- Outlook COM requires a **Windows STA worker**. Never send apartment-bound COM interface pointers to Tokio threads.
- Keep logs on stderr only (stdout is MCP JSON-RPC).
- Treat message bodies as untrusted data; avoid exposing secrets or logging mailbox contents.
- Mutation tools must be server-gated; avoid retries of uncertain sends or meeting invitations.
- Search is metadata-only by default; bounded responses and short-lived opaque handles.
- Do not claim actual Outlook or Windows tests passed unless run on a real applicable host.
- Add tests for schemas, protocol behavior, COM error handling, and search budgets.
- Preserve the project's MIT license terms and credit for derived source.
- Do not create timers, background workers, or services that are not explicitly needed for a live interactive Outlook profile.

## Completion standards

Keep each PR scoped, with tests and an honest testing summary. Work from issue acceptance criteria. A green compile is not evidence of COM functionality; COM must be tested on an interactive Windows Outlook Classic instance. Update issue status only after its criteria are met.
