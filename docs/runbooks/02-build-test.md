# Runbook 02 — Build, lint and test the Rust MCP

**Use when:** Changing Rust code, dependencies, tool contracts or preparing a draft PR.

## Prerequisites

Windows toolchain from [runbook 01](01-development-setup.md); checkout of a fork-owned feature branch. Run commands from the **repository root**.

## Procedure

```powershell
git status --short --branch
cargo fmt --manifest-path rust/Cargo.toml --all -- --check
cargo clippy --locked --manifest-path rust/Cargo.toml --all-targets -- -D warnings
cargo test --locked --manifest-path rust/Cargo.toml --all-targets
cargo build --locked --manifest-path rust/Cargo.toml
.\rust\target\debug\outlook-mcp-rs.exe --version
.\rust\target\debug\outlook-mcp-rs.exe --help
```

If formatting fails and you intend to reformat source, run `cargo fmt --manifest-path rust/Cargo.toml --all`, inspect the diff, then repeat the checks. **Do not** automatically commit formatting that changes unrelated code.

Optional release-optimization build (not a release artifact or packaging test):

```powershell
cargo build --release --locked --manifest-path rust/Cargo.toml
```

## Expected result

All commands exit successfully. The CLI prints version/help; **it does not start the MCP transport in `--help` or `--version` mode**. Unit tests use synthetic data, not live Outlook.

## CI

`.github/workflows/rust-ci.yml` runs the same core checks on hosted `windows-latest`. It triggers for relevant `rust/**`, `CLAUDE.md`, or workflow changes in feature branches and PRs. A docs-only PR may not run this CI due to workflow path filters. If a source PR is red, open the failed job and fix the first real error; preserve `--locked`.

The unrelated `publish.yml` workflow is the **Python/PyPI** publishing pipeline, not a Rust binary release.

## Common failures

- Formatting: apply rustfmt locally and review the diff.
- Clippy: treat warnings as errors and avoid broad suppression without explanation.
- Lockfile mismatch: review `Cargo.toml` / `Cargo.lock` changes together. Only regenerate dependencies intentionally.
- COM behavior failing on workstation with green CI: build checks are not COM integration; use [runbook 06](06-troubleshooting.md).

## Evidence

Save command, build commit SHA, Rust version and pass/fail. Keep test output sanitized. Continue to [MCP smoke testing](04-mcp-smoke-test.md).
