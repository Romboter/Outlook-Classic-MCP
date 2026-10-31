# Runbook 10 — Rust release preflight and go/no-go

**Use when:** Evaluating a distributable Windows build after the experimental implementation is complete. This is a **future checklist**, not a claim that a release already exists.

**Current decision (2026-10-31): NO-GO.** Build CI passed, but actual Outlook Classic and full MCP integration have not been validated; safety/parity/release tasks are incomplete.

## Technical gates (all required)

- [ ] Source branch/PR provenance verified wholly within **Romboter fork**; explicit owner approval for merge and release.
- [ ] `Cargo.toml` / committed `Cargo.lock` reviewed; format/Clippy/tests/build green on intended release SHA.
- [ ] Interactive Outlook Classic COM attach and reconnect runbook passed on representative Windows versions.
- [ ] MCP Inspector initializes, lists tools and calls all read tools on approved real-profile test cases.
- [ ] `Items.Restrict`, `Folder.GetTable`, and opt-in `AdvancedSearch` results checked for correctness and performance against representative folders/stores; indexed callback reliably completes.
- [ ] ConversationIDs, cross-folder/cross-store thread resolution, quote trimming, and raw recovery tested against real conversations.
- [ ] Input validation, body/result budgets, short-ID eviction and privacy review verified with negative cases.
- [ ] Complete Python contract comparison completed, with explicit user-approved parity gaps; do **not** represent 31 Python tool names as 31 Rust implementations.
- [ ] Permission modes and operation-aware retry gates implemented **before any future mutation tools are enabled**.
- [ ] Reproducible Windows x64 release binary generated and tested, published with checksum, license/attribution, change log and dependency provenance; signing decision documented.
- [ ] Installation, version check, upgrade, rollback, no-admin path and coexistence with existing Python server demonstrated; test `claude mcp remove` does not remove the Python server.
- [ ] No confidential mailbox data/credentials in artifacts, CI logs, test fixtures or benchmark reports.
- [ ] Issue #11 acceptance criteria checked, release tag and GitHub release created **only after explicit approval**.

## Candidate build (not a publish step)

```powershell
cargo fmt --manifest-path rust/Cargo.toml --all -- --check
cargo clippy --locked --manifest-path rust/Cargo.toml --all-targets -- -D warnings
cargo test --locked --manifest-path rust/Cargo.toml --all-targets
cargo build --release --locked --manifest-path rust/Cargo.toml
.\rust\target\release\outlook-mcp-rs.exe --version
Get-FileHash .\rust\target\release\outlook-mcp-rs.exe -Algorithm SHA256
```

Hash only identifies the candidate build; it does not validate Outlook integration or certify a release.

## Go/no-go

**GO** only when every required gate is evidenced by an accountable reviewer and owner approves. **NO-GO** on broken COM behavior, false search results, lost new email content, unexpected mailbox mutation, missing privacy controls, uncertain rollback or unresolved compatibility risks.

The repository's existing `.github/workflows/publish.yml` is a Python package publishing workflow. **Do not trigger or repurpose it as a Rust release mechanism** without a separate reviewed change.
