# Outlook Classic MCP — development instructions

## Repository ownership and branch policy

This is the **Romboter fork** and a **fork-specific Rust rewrite**. Never open a pull request against `anasahmed07/Outlook-Classic-MCP`; never sync, pull, fetch, rebase, or cherry-pick from that upstream repository as part of this rewrite. All feature branches and pull requests must be entirely within `Romboter/Outlook-Classic-MCP`. Target this fork's `main` for independent changes or an appropriate **fork-owned feature branch** for stacked PRs while prerequisites are still unmerged. Never target upstream.

Do not commit to `main`. **No pull requests** (owner decision, 2026-11-03): do the work on `rust/development`, run the gates, and commit straight to `rust/development` **only when the owner says to commit**. Never touch `main` or upstream.

## Work order and scope

Read [#1](https://github.com/Romboter/Outlook-Classic-MCP/issues/1), then the individual issues:
- Foundation [#2](https://github.com/Romboter/Outlook-Classic-MCP/issues/2)
- COM [#3](https://github.com/Romboter/Outlook-Classic-MCP/issues/3)
- Read-only MCP [#4](https://github.com/Romboter/Outlook-Classic-MCP/issues/4)
- Indexed search [#5](https://github.com/Romboter/Outlook-Classic-MCP/issues/5)
- Token-efficiency [#6](https://github.com/Romboter/Outlook-Classic-MCP/issues/6)
- Conversation retrieval [#7](https://github.com/Romboter/Outlook-Classic-MCP/issues/7)

The primary user's needs: Outlook **Classic** is always running at work and caches about one year of mail. Claude Code should search existing indexed mail quickly without filling its context. No direct OST parsing or separate Tantivy/SQLite index for the initial implementation.

## Existing implementation

The working Python server is in `src/outlook_mcp/`. Keep it working during Rust development. The experimental Rust crate lives in `rust/`. Do not redirect existing plugin/client configurations until Rust has passed real Outlook Classic integration tests. No false-positive tools that claim to read mail before the COM adapter exists.

Preserve 31-tool compatibility where possible; see `rust/compatibility/tools.md`. The Rust branch stack through draft PR #18 exposes six experimental read-only Outlook tools plus `outlook_rust_status`; do **not** claim this code is merged to `main` or tested against an actual Outlook Classic mailbox.

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

## Project documentation

Start at [`docs/README.md`](docs/README.md) for architecture, tool contracts, development rules, risk notes and task-specific runbooks under `docs/runbooks/`. Maintain runbooks as behavior changes, link test evidence to the relevant issue and never publish private mailbox content. The detailed original search/thread testing notes also remain at `rust/docs/`.

## Completion standards

Keep each PR scoped, with tests and an honest testing summary. Work from issue acceptance criteria. A green compile is not evidence of COM functionality; COM must be tested on an interactive Windows Outlook Classic instance. Update issue status only after its criteria are met.
