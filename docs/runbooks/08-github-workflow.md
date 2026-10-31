# Runbook 08 — Develop in the fork, update issues, and review CI

**Use when:** Taking on a new issue, creating a branch, opening or reviewing a PR, or updating the roadmap.

## Non-negotiable repository policy

- Repository: `Romboter/Outlook-Classic-MCP` only.
- **Never** fetch, pull, rebase, cherry-pick or sync from `anasahmed07/Outlook-Classic-MCP` for this rewrite; never submit upstream PRs.
- Do not commit to `main` or merge without owner approval.
- Feature work remains in dedicated branches. Because earlier Rust PRs are stacked and unmerged, base dependent work on the latest corresponding **fork-owned feature branch**, not on `main` if the required code is absent.

## Procedure

1. Read [roadmap issue #1](https://github.com/Romboter/Outlook-Classic-MCP/issues/1), the target issue, related PRs, and [status](../STATUS.md).
2. Verify branch and remote:
   ```powershell
   git remote -v
   git status --short --branch
   git fetch origin
   ```
   Proceed only when `origin` is your fork; do not use any upstream remote.
3. Example new branch based on the active Rust feature stack:
   ```powershell
   git switch feat/rust-thread-retrieval-issue-7
   git pull --ff-only origin feat/rust-thread-retrieval-issue-7
   git switch -c feat/issue-XX-short-task
   ```
   For a documentation-only task, use a `docs/` branch.
4. Implement one logical scope. Add/adjust tests and docs. Run [runbook 02](02-build-test.md) for Rust changes.
5. Commit; publish to **origin** only:
   ```powershell
   git status --short
   git add -- <paths-you-intend-to-commit>
   git diff --cached --stat
   git commit -m "docs: describe task runbooks"
   git push -u origin HEAD
   ```
6. Open a **draft PR in your fork**. For stacked changes, select the branch that contains the dependency as the PR **base**. Include linked issue, what changed, limitations, testing evidence, rollback, and a clear statement that real Outlook remains untested if applicable.
7. Inspect GitHub Actions from the actual PR head. A docs-only PR might not trigger `rust-ci.yml` because of path filters; do not imply green CI was performed on docs-only changes if it did not run.
8. Update the original issue's status section **without closing it** until every acceptance criterion is met. Update roadmap #1 for new features or material status changes.
9. Keep a concise handoff: branch, base, PR number, last CI status, next technical action, open risks.

## PR description template

```text
Summary:
Linked issue:
Branch / base (both in Romboter fork):
Changes:
Validation performed (specific commit and test):
Not tested / blockers:
Data safety / rollback:
Merge status: draft; no merge until owner approves.
```

## Review and merges

Never auto-merge the stacked Rust PRs just because CI passes. Verify that each lower dependency PR is ready and that merging/rebasing stays **within this fork**. A PR currently targeting a feature branch may need its base adjusted during intentional integration; do not target upstream.

A `README` or `docs` update is not proof of live Outlook validation. Document exactly what someone actually tested.
