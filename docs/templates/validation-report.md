# Validation report — template

> Copy this template into a private, approved location. Do not commit sensitive mailbox content or identifiers to this repository.

**Task / runbook:**  
**Issue / draft PR:**  
**Git commit SHA:**  
**Tester / review role:**  
**Date / test environment:**  
**Windows version:**  
**Outlook Classic build / store category:**  
**Rust version / binary build:**  

## Preconditions

- [ ] Work performed in the Romboter fork / private test environment
- [ ] Explicit test permission / permitted mailbox data
- [ ] Existing Python MCP remains operational
- [ ] Required branch and toolchain verified

## Actions performed

| Step | Command/tool (sanitized) | Expected | Actual (sanitized) | Pass/fail |
| --- | --- | --- | --- | --- |
| 1 |  |  |  |  |
| 2 |  |  |  |  |
| 3 |  |  |  |  |

## Results and observations

**Backend used (if search):**  
**Anonymized result count:**  
**Cold / median / p95 timings (ms, if measured):**  
**Has more / truncated / omitted counts:**  
**COM reconnect behavior:**  
**Quote trimming comparison with original:**  
**Expected versus actual behavior:**  
**Known limitations:**  

## Failure and rollback

**Error category (no raw mailbox output):**  
**Reproduction steps without private identifiers:**  
**Recovery attempted:**  
**Rollback performed / result:**  
**Issue link / next action:**  

## Approval

- [ ] CI verified (record workflow ID and target SHA) — *CI is not live Outlook testing*
- [ ] Live Outlook validation completed — *check only if actually run*
- [ ] Sensitive details removed from the shared report
- [ ] Reviewer approval recorded for the next stage

**Decision:** PASS / FAIL / BLOCKED / NOT RUN

**Evidence rules:** Never include real email subjects/bodies, personal addresses, customer/employee names, EntryIDs, StoreIDs, credentials, or private absolute paths.
