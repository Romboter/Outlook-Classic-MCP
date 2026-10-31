# Outlook Classic MCP — project documentation

**Repository:** [Romboter/Outlook-Classic-MCP](https://github.com/Romboter/Outlook-Classic-MCP)  
**Documentation baseline:** 2026-10-31, stacked Rust work through draft PR [#18](https://github.com/Romboter/Outlook-Classic-MCP/pull/18).

This fork contains two implementations: the **existing Python Outlook Classic MCP server**, which must remain functional, and an **experimental Windows-native Rust rewrite** in `rust/`. The Rust implementation passes hosted Windows build/tests but **has not been validated on a live Outlook Classic mailbox**. Treat every installation or search example as *test-only* until the live-validation gates pass.

## Start here

| Need | Document |
| --- | --- |
| Understand what this project does and what is complete | [Project overview](PROJECT.md) |
| Understand processes, boundaries and data flow | [Architecture](ARCHITECTURE.md) |
| See all Rust tools, bounds and current limitations | [Tool reference](TOOLS.md) |
| Contribute or hand off development | [Development guide](DEVELOPMENT.md) |
| Understand design decisions and unverified assumptions | [Decisions and risks](DECISIONS.md) |
| Protect company mailbox data | [Security and privacy](SECURITY.md) |
| Check outstanding work and release gates | [Project status](STATUS.md) |
| Perform a task step by step | [Runbook index](runbooks/README.md) |

## Task runbooks

- [Prepare a Windows development workstation](runbooks/01-development-setup.md)
- [Build, format, lint and test the Rust crate](runbooks/02-build-test.md)
- [Validate Outlook Classic COM attachment on a real profile](runbooks/03-outlook-com-validation.md)
- [Smoke-test the MCP protocol and read-only tools](runbooks/04-mcp-smoke-test.md)
- [Benchmark Table, legacy and indexed email search](runbooks/05-search-benchmark.md)
- [Troubleshoot Outlook COM and indexed-search errors](runbooks/06-troubleshooting.md)
- [Validate conversation/thread retrieval and quote trimming](runbooks/07-thread-validation.md)
- [Work on issues and maintain the fork-only branch/PR stack](runbooks/08-github-workflow.md)
- [Trial deployment, configuration and rollback](runbooks/09-test-deployment-rollback.md)
- [Release-readiness checklist — do not publish prematurely](runbooks/10-release-preflight.md)
- [Copyable test evidence template](templates/validation-report.md)

## Canonical code and existing detailed guides

- Rust implementation and build entrypoint: [`../rust/README.md`](../rust/README.md)
- Python-to-Rust 31-tool compatibility inventory: [`../rust/compatibility/tools.md`](../rust/compatibility/tools.md)
- Detailed search measurements: [`../rust/docs/search-benchmark.md`](../rust/docs/search-benchmark.md)
- Detailed conversation testing: [`../rust/docs/thread-retrieval.md`](../rust/docs/thread-retrieval.md)
- Repository agent/development rules: [`../CLAUDE.md`](../CLAUDE.md)

**Maintenance rule:** Update a runbook when the code or a tested procedure changes. Record real execution evidence separately from examples. Never paste real emails, MAPI IDs, credentials or personally identifiable data into a GitHub issue, PR or CI log.
