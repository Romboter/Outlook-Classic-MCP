# Design decisions and unresolved risks

These are the decisions embodied in current code, not a claim that every design has been verified live.

| Decision | Reason | Consequence / verification |
| --- | --- | --- |
| Fork-owned rewrite with no upstream synchronization | Independent private direction | PR bases and remotes must stay inside Romboter fork |
| Preserve Python while building Rust | No regression to working server | Rust remains separate in `rust/`; don't switch production client |
| Local MCP stdio over an HTTP service | Outlook is interactive/local; smaller exposed surface | stdout reserved for JSON-RPC, stderr logs |
| COM on one dedicated STA with message pumping | Outlook COM apartment constraints | Live event callback, busy/modal and restart tests pending |
| Attach to existing Outlook Classic instance | The profile is already open and authenticated | Opening New Outlook or running unattended won't satisfy COM |
| Outlook-owned search; no raw OST parsing/index copy | Lower storage, synchronization and security cost | Must benchmark actual `Items.Restrict`, Table and indexed correctness |
| `AdvancedSearch` is explicit opt-in | Event completion and provider support not live-tested | `auto` stays on existing paths until validated |
| Compact session handles | Avoid repeatedly sending large MAPI IDs to the LLM | Handles expire; cache 2,048; not authorization tokens |
| Read-only Rust surface | Avoid accidental sends/deletions during migration | Full permissions and operation-aware retry work deferred to issue #8 |
| Conservative quoted-history trimming | Reduce duplicate conversation text without inventing summaries | Expose omissions and raw message recovery; compare with originals |
| Keep dates as locale strings initially | Typed UTC conversion requires live provider testing | Never label `received_local` as UTC |
| Branch+draft PR for every change | Reviewable, reversible work | No direct `main` writes or merges without owner approval |

## Risks still open

1. **COM runtime validity:** Running/compiled Windows code does not establish successful attachment, event subscription or store row conversions.
2. **Search accuracy and latency:** No representative user-profile cold/warm measurements. Indexed response timing, coverage, locale and provider behavior unproven.
3. **Conversation content loss:** Recognized forwarded/original separators may appear within new content. Quote trimming is a heuristic, not a perfect semantic parser.
4. **Sensitive content exposure:** Mail subjects and bodies are potentially proprietary. Local MCP access still requires trusted clients, endpoint permissions and secure logs.
5. **Schema migration:** Python's 31 named tools lack a complete machine-readable result/input parity snapshot in the Rust repo.
6. **Release process:** Windows packaging/signing and client cutover/rollback are not yet validated.

Before changing an architectural choice, add a dated section to this file containing the decision, alternatives, evidence, impact and PR/issue link.
