# Operational runbooks

These instructions cover the **experimental Rust rewrite**, not production support of the existing Python server. All commands assume a checkout of the user's **Romboter fork** and a trusted workstation. Runbooks describe procedures to perform; they are **not evidence** that a live Outlook validation has already occurred.

| Task | Runbook | Completion evidence |
| --- | --- | --- |
| Configure Windows for Rust and Outlook development | [01 — Development setup](01-development-setup.md) | Toolchain versions and fork remote verified |
| Confirm a change builds and passes tests | [02 — Build and test](02-build-test.md) | Format, Clippy, tests and build pass |
| Prove Rust attaches to real Outlook Classic | [03 — Outlook COM validation](03-outlook-com-validation.md) | Local probe succeeds; redacted evidence |
| Test MCP handshake, schemas, and read-only calls | [04 — MCP smoke test](04-mcp-smoke-test.md) | Correct tools, bounded responses, explicit failures |
| Compare real search implementations | [05 — Search benchmark](05-search-benchmark.md) | Same-query cold/warm results and correctness |
| Diagnose COM/Outlook/indexed-search failures | [06 — Troubleshooting](06-troubleshooting.md) | Root-cause category, safe recovery |
| Verify conversation retrieval | [07 — Thread validation](07-thread-validation.md) | Message association and quote accuracy checked |
| Manage branches, PRs, issues and CI | [08 — GitHub workflow](08-github-workflow.md) | Fork-only PR, issue and test evidence |
| Test installation, rollback and coexistence | [09 — Trial deployment/rollback](09-test-deployment-rollback.md) | Existing Python unaffected; removal verified |
| Decide if a Rust release is ready | [10 — Release preflight](10-release-preflight.md) | All release gates documented, no assumptions |

**Common practices:** Always use synthetic or permission-approved test email. Do not attach real mail JSON, sender addresses or MAPI identifiers to tickets. Treat `cargo`/CI success as build evidence only. Keep `main` and upstream unchanged; put improvements on new feature branches and draft PRs.

See [project status](../STATUS.md), [security](../SECURITY.md), and the [validation-report template](../templates/validation-report.md).
