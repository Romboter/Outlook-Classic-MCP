# Runbook 07 — Validate conversation retrieval and text trimming

**Use when:** Testing `outlook_get_thread` from issue #7. This is **not** satisfied by successful Rust compilation.

**Detailed procedure:** [`rust/docs/thread-retrieval.md`](../../rust/docs/thread-retrieval.md).

## Prerequisites

A trusted Outlook Classic profile and an approved test conversation with at least two replies. For cross-folder and cross-store verification, obtain test messages in Inbox/Sent Items and (if supported) two stores. Use Inspector from [runbook 04](04-mcp-smoke-test.md).

## Steps

1. Find an approved test message via `outlook_search_mails` (start with `backend:legacy`).
2. Copy its *returned* short `entry_id` within the same MCP session.
3. Invoke:
   ```json
   {"entry_id":"m1","max_messages":10,"include_bodies":true,"trim_quoted":true,"per_message_chars":1500,"total_body_chars":8000}
   ```
   Replace `m1` with the actual returned handle.
4. Confirm `association` is `verified_conversation_id`, count/ordering plausible, and no unrelated messages merely sharing the subject appear.
5. Compare `trim_quoted:true` against `trim_quoted:false` using **the same test conversation**. For each message, retrieve the original via `outlook_get_mail` when verifying that new content was not suppressed.
6. Check omission/truncation metadata: `older_messages_omitted`, `scan_truncated`, `inaccessible_messages`, `missing_store_identity`, `identity_mismatch`, `anchor_included`, `quote_history_omitted`, and `body_truncated`.
7. Repeat with `include_bodies:false` and with lower budgets. Confirm no body exceeds per-message/total limits, including Unicode characters.
8. Check a single-message conversation, missing/deleted message, duplicate subject across unrelated threads, and (where available) cross-store ConversationID. These cases must not crash or invent missing messages.
9. Close/restart the Rust MCP client, repeat search/list to get new handles, and ensure old handles return a clear error.

## Expected result

- Exact native ConversationID association; no subject-only joining.
- Newest bounded subset ordered chronologically; omission metadata truthful.
- Short message handles resolve to the correct originals in the same session.
- For quote trimming, **no newly written content lost** in approved samples. Preserve raw mode for uncertain patterns.

## Failure / safety

If messages are missing, StoreID conversion is wrong, or trimmed output loses new text, record the test as **failed**. Do not promote thread retrieval as reliable or merge as complete. Compare directly in Outlook locally; never publish message contents/identifiers in issues.

**Known Outlook behavior:** `Conversation.GetTable` omits Deleted Items. Some stores do not enable conversations. Hard scan cap in current implementation is 350 rows; the default response does not imply a complete lifelong conversation.

## Evidence

Use [test report template](../templates/validation-report.md) with sanitized thread counts and pass/fail. No real subjects, addresses, bodies, EntryIDs or StoreIDs.
