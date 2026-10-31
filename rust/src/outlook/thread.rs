//! Read-only, bounded thread retrieval using Outlook's native Conversation.
//! No subject-based joins, writes, or COM access outside the worker STA.
//! Deleted Items are intentionally absent from Conversation.GetTable.

use super::win;
use crate::{
    contracts::{GetMailParams, ThreadParams},
    thread_text::{limit_chars, trim_quote_tail},
};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::cmp::Ordering;
use windows::{
    Win32::System::{
        Com::{DISPATCH_METHOD, DISPATCH_PROPERTYGET, IDispatch},
        Variant::{VAR_CHANGE_FLAGS, VARIANT, VT_BSTR, VT_R8, VariantChangeType},
    },
    core::BSTR,
};

// Documented property: PR_STORE_ENTRYID (binary) in Conversation.GetTable.
// Row.BinaryToString converts it to the precise StoreID that GetItemFromID
// needs when messages span stores.
const STORE_ENTRY_ID: &str = "https://schemas.microsoft.com/mapi/proptag/0x0FFB0102";
const MAX_ROWS: usize = 350;

struct Candidate {
    entry_id: String,
    store_id: String,
    sort_time: f64,
    ordinal: usize,
}

fn cell(row: &IDispatch, name: &str) -> Option<VARIANT> {
    win::invoke(row, "Item", DISPATCH_PROPERTYGET, vec![VARIANT::from(name)])
        .or_else(|_| win::invoke(row, "Item", DISPATCH_METHOD, vec![VARIANT::from(name)]))
        .ok()
}

fn text_variant(value: &VARIANT) -> Option<String> {
    if let Ok(v) = BSTR::try_from(value) {
        return Some(v.to_string());
    }
    let mut coerced = VARIANT::default();
    unsafe { VariantChangeType(&mut coerced, value, VAR_CHANGE_FLAGS(0), VT_BSTR) }.ok()?;
    BSTR::try_from(&coerced).ok().map(|v| v.to_string())
}

fn row_text(row: &IDispatch, name: &str) -> String {
    cell(row, name)
        .as_ref()
        .and_then(text_variant)
        .unwrap_or_default()
}

fn date_number(row: &IDispatch, name: &str) -> Option<f64> {
    let value = cell(row, name)?;
    let mut converted = VARIANT::default();
    unsafe { VariantChangeType(&mut converted, &value, VAR_CHANGE_FLAGS(0), VT_R8) }.ok()?;
    let date = f64::try_from(&converted).ok()?;
    date.is_finite().then_some(date)
}

fn mail_text(item: &IDispatch, name: &str) -> String {
    win::property(item, name)
        .ok()
        .as_ref()
        .and_then(text_variant)
        .unwrap_or_default()
}

fn get_by_id(namespace: &IDispatch, entry: &str, store: &str) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(entry)];
    if !store.is_empty() {
        args.push(VARIANT::from(store));
    }
    win::method_object(namespace, "GetItemFromID", args)
}

fn is_conversation_supported(item: &IDispatch) -> Result<bool> {
    let folder = win::object_property(item, "Parent")
        .context("Cannot access parent folder of selected message")?;
    let store = win::object_property(&folder, "Store")
        .context("Cannot access store of selected message")?;
    let enabled = win::property(&store, "IsConversationEnabled")
        .context("Store does not report conversation support")?;
    bool::try_from(&enabled).context("Store returned invalid IsConversationEnabled")
}

fn collect_candidates(table: &IDispatch) -> Result<(Vec<Candidate>, bool, usize)> {
    // Conversation.GetTable supplies default EntryID, MessageClass, CreationTime
    // and maintains conversation index order. Do not RemoveAll.
    let columns = win::object_property(table, "Columns")?;
    win::invoke(
        &columns,
        "Add",
        DISPATCH_METHOD,
        vec![VARIANT::from(STORE_ENTRY_ID)],
    )
    .context("Conversation table cannot provide cross-store IDs safely")?;
    // Optional date columns help chronological ordering; no heuristic sorting
    // of localized date display strings.
    for col in ["SentOn", "ReceivedTime"] {
        let _ = win::invoke(&columns, "Add", DISPATCH_METHOD, vec![VARIANT::from(col)]);
    }
    let mut candidates = Vec::new();
    let mut examined = 0usize;
    let mut missing_identity = 0usize;
    loop {
        let end = bool::try_from(&win::property(table, "EndOfTable")?)
            .context("Invalid Conversation.Table EndOfTable value")?;
        if end {
            return Ok((candidates, false, missing_identity));
        }
        if examined >= MAX_ROWS {
            return Ok((candidates, true, missing_identity));
        }
        examined += 1;
        let row = win::method_object(table, "GetNextRow", vec![])?;
        let class = row_text(&row, "MessageClass");
        if !class.starts_with("IPM.Note") {
            continue;
        }
        let entry = row_text(&row, "EntryID");
        let store = win::invoke(
            &row,
            "BinaryToString",
            DISPATCH_METHOD,
            vec![VARIANT::from(STORE_ENTRY_ID)],
        )
        .ok()
        .as_ref()
        .and_then(text_variant)
        .unwrap_or_default();
        if entry.is_empty() || store.is_empty() {
            missing_identity += 1;
            continue;
        }
        let sort_time = date_number(&row, "SentOn")
            .or_else(|| date_number(&row, "ReceivedTime"))
            .or_else(|| date_number(&row, "CreationTime"))
            .unwrap_or(0.0);
        candidates.push(Candidate {
            entry_id: entry,
            store_id: store,
            sort_time,
            ordinal: examined,
        });
    }
}

impl win::Session {
    pub fn read_thread(&mut self, params: ThreadParams) -> Result<Value> {
        let max_messages = params.max_messages.unwrap_or(10) as usize;
        let per_message = params.per_message_chars.unwrap_or(1500) as usize;
        let total_budget = params.total_body_chars.unwrap_or(8000) as usize;
        let trim = params.trim_quoted.unwrap_or(true);
        let include_bodies = params.include_bodies.unwrap_or(true);
        let mut anchor = GetMailParams {
            entry_id: params.entry_id,
            store_id: params.store_id,
            include_body: Some(false),
            include_html: Some(false),
            max_body_chars: None,
        };
        self.refs.resolve_input(&mut anchor)?;
        let selected = get_by_id(
            &self.namespace,
            &anchor.entry_id,
            anchor.store_id.as_deref().unwrap_or(""),
        )
        .context("Selected mail is unavailable; repeat your search/list")?;
        if win::int_property(&selected, "Class")? != 43 {
            bail!("Selected entry is not an Outlook MailItem");
        }
        if !is_conversation_supported(&selected)? {
            bail!(
                "This Outlook store does not support Conversation view; read individual messages instead"
            );
        }
        let conv_id = win::string_property(&selected, "ConversationID")
            .context("Selected mail does not provide ConversationID")?;
        if conv_id.is_empty() {
            bail!("Selected mail has no ConversationID");
        }
        let conversation = win::method_object(&selected, "GetConversation", vec![])
            .context("Outlook GetConversation returned no available conversation")?;
        let actual_id = win::string_property(&conversation, "ConversationID")?;
        if actual_id != conv_id {
            bail!("Conversation identity changed; cannot safely associate messages");
        }
        let table = win::method_object(&conversation, "GetTable", vec![])
            .context("Conversation.GetTable unavailable for this Outlook store")?;
        let (mut candidates, scan_truncated, missing_identity) = collect_candidates(&table)?;
        // Sort using typed Outlook DATE values, not localized strings. On ties,
        // preserve the native ConversationIndex row order.
        candidates.sort_by(|a, b| {
            a.sort_time
                .partial_cmp(&b.sort_time)
                .unwrap_or(Ordering::Equal)
                .then(a.ordinal.cmp(&b.ordinal))
        });
        let discovered = candidates.len();
        let older_omitted = discovered.saturating_sub(max_messages);
        let recent = candidates.into_iter().skip(older_omitted);

        let mut output = Vec::new();
        let mut total_retained = 0usize;
        let mut inaccessible = 0usize;
        let mut identity_mismatch = 0usize;
        let mut suppressed_quotes = 0usize;
        let mut body_truncated_messages = 0usize;
        let mut anchor_included = false;

        for candidate in recent {
            let item = match get_by_id(&self.namespace, &candidate.entry_id, &candidate.store_id) {
                Ok(item) => item,
                Err(_) => {
                    inaccessible += 1;
                    continue;
                }
            };
            if win::int_property(&item, "Class").unwrap_or_default() != 43 {
                identity_mismatch += 1;
                continue;
            }
            // Fail closed on unavailable IDs; never join by matching Subject.
            if win::string_property(&item, "ConversationID")
                .ok()
                .as_deref()
                != Some(conv_id.as_str())
            {
                identity_mismatch += 1;
                continue;
            }
            if candidate.entry_id == anchor.entry_id
                && candidate.store_id == anchor.store_id.as_deref().unwrap_or("")
            {
                anchor_included = true;
            }
            let mut msg = json!({
                "entry_id":candidate.entry_id,"store_id":candidate.store_id,
                "subject":mail_text(&item,"Subject"),
                "from":mail_text(&item,"SenderName"),
                "from_address":mail_text(&item,"SenderEmailAddress"),
                "sent_local":mail_text(&item,"SentOn"),
                "received_local":mail_text(&item,"ReceivedTime")
            });
            if include_bodies {
                let raw = mail_text(&item, "Body");
                let raw_chars = raw.chars().count();
                let trimmed = if trim {
                    trim_quote_tail(&raw)
                } else {
                    crate::thread_text::Trimmed {
                        text: raw,
                        quote_removed: false,
                        chars_removed: 0,
                    }
                };
                let budget = per_message.min(total_budget.saturating_sub(total_retained));
                let (body, truncated) = limit_chars(&trimmed.text, budget);
                total_retained += body.chars().count();
                if trimmed.quote_removed {
                    suppressed_quotes += 1;
                }
                if truncated {
                    body_truncated_messages += 1;
                }
                msg["body"] = json!(body);
                msg["body_original_chars"] = json!(raw_chars);
                msg["quote_history_omitted"] = json!(trimmed.quote_removed);
                msg["quote_chars_omitted"] = json!(trimmed.chars_removed);
                msg["body_truncated"] = json!(truncated);
            }
            output.push(msg);
        }
        Ok(json!({
            "association":"verified_conversation_id",
            "order":"chronological_by_outlook_date_with_conversation_index_ties",
            "messages":output,
            "count":output.len(),
            "matched_candidates":discovered,
            "older_messages_omitted":older_omitted,
            "scan_truncated":scan_truncated,
            "max_scanned_rows":MAX_ROWS,
            "missing_store_identity":missing_identity,
            "inaccessible_messages":inaccessible,
            "identity_mismatch":identity_mismatch,
            "anchor_included":anchor_included,
            "quote_trimming_enabled":trim,
            "messages_with_quote_omissions":suppressed_quotes,
            "messages_with_body_truncation":body_truncated_messages,
            "total_body_chars_returned":total_retained,
            "total_body_chars_limit":total_budget,
            "note":"Conversation.GetTable omits Deleted Items. Quoted history trimming is conservative, but may remove meaningful text; retrieve any message raw with outlook_get_mail when accuracy matters."
        }))
    }
}
