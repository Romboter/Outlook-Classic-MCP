# Runbook 01 — Prepare a Windows development workstation

**Use when:** Onboarding a developer or setting up a clean machine to build the experimental Rust MCP.

## Prerequisites

- Windows 10/11 and permission to install dev tools.
- Git, Rust **1.88+** and cargo; rustfmt, Clippy, and a compatible MSVC build toolchain (Visual Studio Build Tools, C++ workload).
- For **live** testing only: installed Outlook **Classic**, running with a ready profile in the *same interactive Windows session*. New Outlook is unsupported.
- Network access for initial Rust dependency downloads and GitHub access.

## Steps (PowerShell)

1. Verify the toolchain:
   ```powershell
   git --version
   rustc --version
   cargo --version
   rustup show active-toolchain
   ```
2. Clone **only your fork** and check the remote:
   ```powershell
   git clone https://github.com/Romboter/Outlook-Classic-MCP.git
   cd Outlook-Classic-MCP
   git remote -v
   ```
   Expected `origin` contains `Romboter/Outlook-Classic-MCP`. Do **not** fetch/pull/rebase from the original author's upstream.
3. Select the active unmerged Rust feature branch:
   ```powershell
   git fetch origin
   git switch feat/rust-thread-retrieval-issue-7
   git status --short --branch
   ```
   If the branch is already merged or superseded, consult [status](../STATUS.md) and the PR chain before choosing a replacement. Avoid assuming Rust is on `main`.
4. Build without modifying the lockfile:
   ```powershell
   cargo build --locked --manifest-path rust/Cargo.toml
   .\rust\target\debug\outlook-mcp-rs.exe --version
   ```
5. On a trusted Outlook Classic workstation, check that `OUTLOOK.EXE` is already running:
   ```powershell
   Get-Process OUTLOOK -ErrorAction SilentlyContinue |
     Select-Object ProcessName,Id,SessionId
   ```

## Expected result

- Rust/Cargo versions meet the crate's minimum, the build succeeds, and the binary prints its version.
- `origin` targets your fork.
- For live testing, Outlook Classic is running under the same user/session; this **does not** itself validate COM.

## Failure / recovery

- `cargo: command not found`: install the Rust toolchain; reopen PowerShell and verify PATH.
- MSVC linker/tool error: install the required Visual Studio C++ Build Tools. Do not change the Rust target to work around COM compatibility without reviewing impact.
- No `OUTLOOK.EXE`: start Outlook Classic normally, sign in, wait for the profile and folders to load. Avoid launching it through a headless service.
- Wrong remote or unexpected branch: stop before fetching/merging; compare `git remote -v` with the required fork URL.

## Record

Only toolchain and OS versions, build pass/fail and whether Outlook Classic was running. Do not capture profile identity or mailbox content in shared logs. Continue with [build/test](02-build-test.md) or [COM validation](03-outlook-com-validation.md).
