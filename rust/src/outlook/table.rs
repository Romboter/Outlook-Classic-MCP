//! Lightweight Outlook Table row enumeration. It deliberately never opens a
//! MailItem per result; full message access remains a separate explicit tool.
//!
//! Folder.GetTable is a provider-backed rowset, NOT necessarily an indexed
//! full-text search. Cross-folder indexed AdvancedSearch is a later step.
use super::win;
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use windows::{
    Win32::System::{
        Com::{DISPATCH_METHOD, DISPATCH_PROPERTYGET, IDispatch},
        Variant::{VAR_CHANGE_FLAGS, VARIANT, VT_BSTR, VariantChangeType},
    },
    core::BSTR,
};

const MESSAGE_FLAGS: &str = "http://schemas.microsoft.com/mapi/proptag/0x0E070003";
const MAX_SCANNED: usize = 10_000;

#[derive(Debug, Clone)]
pub struct RowOptions<'a> {
    pub filter: &'a str,
    pub limit: usize,
    pub offset: usize,
    pub unread_only: bool,
    pub sender_substring: Option<&'a str>,
}

pub struct RowPage {
    pub items: Vec<Value>,
    pub has_more: bool,
    pub scanned: usize,
    pub scan_truncated: bool,
}

fn cell(row: &IDispatch, name: &str) -> Option<VARIANT> {
    // Row.Item(name) is the default indexed property. Some COM versions
    // expose it as a method instead of PROPERTYGET.
    let args = || vec![VARIANT::from(name)];
    win::invoke(row, "Item", DISPATCH_PROPERTYGET, args())
        .or_else(|_| win::invoke(row, "Item", DISPATCH_METHOD, args()))
        .ok()
}

fn text_from_variant(value: &VARIANT) -> Option<String> {
    if let Ok(value) = BSTR::try_from(value) {
        return Some(value.to_string());
    }
    let mut coerced = VARIANT::default();
    unsafe { VariantChangeType(&mut coerced, value, VAR_CHANGE_FLAGS(0), VT_BSTR) }.ok()?;
    BSTR::try_from(&coerced).ok().map(|v| v.to_string())
}

fn text(row: &IDispatch, name: &str) -> String {
    cell(row, name)
        .and_then(|v| text_from_variant(&v))
        .unwrap_or_default()
}

fn unread(row: &IDispatch) -> Option<bool> {
    let flags = i32::try_from(&cell(row, MESSAGE_FLAGS)?).ok()?;
    // PR_MESSAGE_FLAGS MSGFLAG_READ is bit zero.
    Some(flags & 1 == 0)
}

pub fn store_instant_search(folder: &IDispatch) -> Option<bool> {
    let store = win::object_property(folder, "Store").ok()?;
    let value = win::property(&store, "IsInstantSearchEnabled").ok()?;
    bool::try_from(&value).ok()
}

pub fn enumerate(folder: &IDispatch, store_id: &str, options: RowOptions<'_>) -> Result<RowPage> {
    // Unlike Items.Restrict, this obtains lightweight table rows without
    // instantiating each Outlook MailItem.
    let table = win::method_object(folder, "GetTable", vec![VARIANT::from(options.filter)])
        .context("Outlook Folder.GetTable unavailable for this folder/filter")?;
    enumerate_existing(&table, store_id, options)
}

/// Enumerate an already-completed Search.GetTable, not a new search.
pub fn enumerate_existing(table: &IDispatch, store_id: &str, options: RowOptions<'_>) -> Result<RowPage> {
    let columns = win::object_property(&table, "Columns")?;
    win::invoke(&columns, "RemoveAll", DISPATCH_METHOD, vec![])?;

    for column in ["EntryID", "Subject", "ReceivedTime", "MessageClass"] {
        win::invoke(
            &columns,
            "Add",
            DISPATCH_METHOD,
            vec![VARIANT::from(column)],
        )
        .with_context(|| format!("Outlook Table lacks required column {column}"))?;
    }
    // Optional columns vary with store and message class. Missing values are
    // returned empty/null, not silently fetched via a separate MailItem.
    let mut sender_name_available = false;
    let mut sender_address_available = false;
    if win::invoke(
        &columns,
        "Add",
        DISPATCH_METHOD,
        vec![VARIANT::from("SenderName")],
    )
    .is_ok()
    {
        sender_name_available = true;
    }
    if win::invoke(
        &columns,
        "Add",
        DISPATCH_METHOD,
        vec![VARIANT::from("SenderEmailAddress")],
    )
    .is_ok()
    {
        sender_address_available = true;
    }
    let flags_available = win::invoke(
        &columns,
        "Add",
        DISPATCH_METHOD,
        vec![VARIANT::from(MESSAGE_FLAGS)],
    )
    .is_ok();
    if options.unread_only && !flags_available {
        bail!("provider does not expose PR_MESSAGE_FLAGS for unread filtering");
    }
    if options.sender_substring.is_some() && !sender_name_available && !sender_address_available {
        bail!("provider does not expose sender columns for sender filtering");
    }

    win::invoke(
        &table,
        "Sort",
        DISPATCH_METHOD,
        vec![VARIANT::from("ReceivedTime"), VARIANT::from(true)],
    )?;
    let sender_filter = options.sender_substring.map(str::to_lowercase);
    let mut collected = Vec::with_capacity(options.limit);
    let mut skipped = 0;
    let mut scanned = 0;

    // EndOfTable is an explicit bool; never default to false on COM errors.
    loop {
        if scanned >= MAX_SCANNED {
            return Ok(RowPage {
                items: collected,
                has_more: true,
                scanned,
                scan_truncated: true,
            });
        }
        let end = bool::try_from(&win::property(&table, "EndOfTable")?)
            .context("Outlook Table.EndOfTable returned an invalid value")?;
        if end {
            break;
        }
        scanned += 1;
        let row = win::method_object(&table, "GetNextRow", vec![])?;
        let class = text(&row, "MessageClass");
        if !class.starts_with("IPM.Note") && !class.starts_with("IPM.Post") {
            continue;
        }
        let entry_id = text(&row, "EntryID");
        if entry_id.is_empty() {
            continue;
        }

        let unread_state = if flags_available { unread(&row) } else { None };
        if options.unread_only {
            // Never include a message with unknown unread state.
            if unread_state != Some(true) {
                continue;
            }
        }
        let sender_name = if sender_name_available {
            text(&row, "SenderName")
        } else {
            String::new()
        };
        let sender_address = if sender_address_available {
            text(&row, "SenderEmailAddress")
        } else {
            String::new()
        };
        if let Some(ref filter) = sender_filter
            && !sender_name.to_lowercase().contains(filter)
            && !sender_address.to_lowercase().contains(filter)
        {
            continue;
        }
        if skipped < options.offset {
            skipped += 1;
            continue;
        }
        if collected.len() >= options.limit {
            return Ok(RowPage {
                items: collected,
                has_more: true,
                scanned,
                scan_truncated: false,
            });
        }
        collected.push(json!({
            "entry_id": entry_id,
            "store_id": store_id,
            "subject": text(&row, "Subject"),
            "from": sender_name,
            "from_address": sender_address,
            "received_local": text(&row, "ReceivedTime"),
            "unread": unread_state
        }));
    }

    Ok(RowPage {
        items: collected,
        has_more: false,
        scanned,
        scan_truncated: false,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn table_scan_cap_is_bounded() {
        assert_eq!(MAX_SCANNED, 10_000);
    }
}
