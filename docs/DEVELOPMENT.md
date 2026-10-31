# Development workflow

## Repository rules

**Only `Romboter/Outlook-Classic-MCP` is in scope.** Do not fetch, pull, sync, merge, rebase or submit a PR to `anasahmed07/Outlook-Classic-MCP`. Do not commit to this fork's `main`. All work uses scoped branches and draft PRs; merge only with the owner's explicit direction.

The current Rust prototype is a **stacked draft PR chain (#12–#18)**. New changes requiring current Rust features can branch from `feat/rust-thread-retrieval-issue-7`. To create documentation-only work that includes all current features, use a docs branch from that same feature tip and target that branch in a stacked PR. This avoids pretending the Rust code already exists on `main`.

## Windows prerequisites

- Windows 10/11, Git, Rust stable **1.88 or newer** with cargo/rustfmt/Clippy, and Visual Studio Build Tools with an appropriate MSVC C++ toolchain when using the MSVC Rust target.
- Outlook **Classic** only for live integration. It must be running, profile ready, under the same interactive account as the Rust process.
- VS Code / PowerShell are optional; commands below work from PowerShell.

## Check out your fork only

```powershell
git clone https://github.com/Romboter/Outlook-Classic-MCP.git
cd Outlook-Classic-MCP
git remote -v
git fetch origin
git switch feat/rust-thread-retrieval-issue-7
git pull --ff-only origin feat/rust-thread-retrieval-issue-7
git switch -c feat/your-scoped-change
```

The `origin` remote must point at **Romboter**. The commands deliberately never touch an upstream remote. If your local clone has an `upstream` remote, do not use it for this rewrite.

## Local verification from repo root

```powershell
rustc --version
cargo --version
cargo fmt --manifest-path rust/Cargo.toml --all -- --check
cargo clippy --locked --manifest-path rust/Cargo.toml --all-targets -- -D warnings
cargo test --locked --manifest-path rust/Cargo.toml --all-targets
cargo build --locked --manifest-path rust/Cargo.toml
.\rust\target\debug\outlook-mcp-rs.exe --version
.\rust\target\debug\outlook-mcp-rs.exe --help
```

For a first-time build, Windows dependency downloads may require network connectivity. Do not remove `--locked` or regenerate `Cargo.lock` without a deliberate dependency-update PR.

## Code-change conventions

- Update `rust/src/contracts.rs` when adding or changing tool arguments; reject unimplemented options instead of silently ignoring them.
- Register tools in `rust/src/lib.rs`; COM access must be dispatched via the dedicated STA in `rust/src/outlook.rs`. Never move `IDispatch` across threads.
- Keep mailbox data out of logs. MCP stdio stdout is protocol-only; diagnostics use stderr.
- Maintain short ID mapping and output truncation. Every new returned list, including `messages`, must pass through the server-side sanitizer in `rust/src/outlook/refs.rs`.
- Add unit tests to cover bounds, unicode, missing items, stale handles, and failure paths. Never add live mailbox data as test fixtures.
- Document any new risk in [decisions](DECISIONS.md), and update the relevant [runbook](runbooks/README.md).
- Do not claim live Outlook behavior based only on GitHub's `windows-latest` runner. It does not run the user's interactive mailbox.

## CI behavior

`.github/workflows/rust-ci.yml` runs on relevant Rust, `CLAUDE.md` or workflow file changes for branch/PR events; it checks formatting, Clippy, tests, build, help/version. **Docs-only changes under `docs/**` are not currently a trigger for this workflow.** PRs still require manual link review. The repository's `publish.yml` targets the original **Python/PyPI** workflow; it does not publish Rust binaries.

Use [GitHub workflow runbook](runbooks/08-github-workflow.md) for branch/PR/issue operations.
