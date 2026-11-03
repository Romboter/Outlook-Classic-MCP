//! Lightweight Outlook Table row enumeration. It deliberately never opens a
//! MailItem per result; full message access remains a separate explicit tool.
//!
//! Folder.GetTable is a provider-backed rowset, NOT necessarily an indexed
//! full-text search. Cross-folder indexed AdvancedSearch is a later step.
use super::win;
use crate::contracts::has_all_words;
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
// Table columns for triage fields (PR_DISPLAY_TO, PR_FLAG_STATUS, PR_IMPORTANCE, body text).
// A table cell holds at most 255 characters of body, enough for the preview.
const DISPLAY_TO: &str = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F";
const FLAG_STATUS: &str = "http://schemas.microsoft.com/mapi/proptag/0x10900003";
// PR_HASATTACH disagreed with Attachments.Count in live checks; the message flag matched it.
const MSGFLAG_HASATTACH: i32 = 0x10;
const IMPORTANCE: &str = "http://schemas.microsoft.com/mapi/proptag/0x00170003";
const BODY_TEXT: &str = "urn:schemas:httpmail:textdescription";
const PREVIEW_CHARS: usize = 200;

/// First 200 characters of a body on one line, like the Python `preview`.
pub fn preview_text(raw: &str) -> String {
    raw.split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .chars()
        .take(PREVIEW_CHARS)
        .collect()
}

fn add_column(columns: &IDispatch, name: &str) -> bool {
    win::invoke(columns, "Add", DISPATCH_METHOD, vec![VARIANT::from(name)]).is_ok()
}

#[derive(Debug, Clone)]
pub struct RowOptions<'a> {
    pub filter: &'a str,
    pub limit: usize,
    pub offset: usize,
    pub unread_only: bool,
    pub sender_substring: Option<&'a str>,
    /// Words that must all appear in the row's subject or sender (lowercase).
    pub words: Option<(&'a [String], WordScope)>,
}

/// Which row text extra search words must appear in.
#[derive(Debug, Clone, Copy)]
pub enum WordScope {
    Subject,
    Sender,
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

fn message_flags(row: &IDispatch) -> Option<i32> {
    i32::try_from(&cell(row, MESSAGE_FLAGS)?).ok()
}

fn unread(row: &IDispatch) -> Option<bool> {
    // PR_MESSAGE_FLAGS MSGFLAG_READ is bit zero.
    Some(message_flags(row)? & 1 == 0)
}

/// Exchange-internal senders report an X.500 DN (`/O=...`), not an SMTP address.
/// Withhold it: it is long, opaque and useless to a client.
// ponytail: blank instead of resolving SMTP via Sender.GetExchangeUser(), which needs a MailItem.
pub fn public_address(address: &str) -> String {
    if address.starts_with('/') {
        String::new()
    } else {
        address.to_string()
    }
}

/// PR_SENDER_SMTP_ADDRESS: the SMTP address even when `SenderEmailAddress` is an X.500 DN.
pub const SENDER_SMTP: &str = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F";

/// The SMTP address for an Exchange (X.500) sender when one was found, else the raw address.
pub fn resolve_sender_address(raw: &str, smtp: Option<String>) -> String {
    match smtp {
        Some(smtp) if raw.starts_with('/') && smtp.contains('@') => smtp,
        _ => raw.to_string(),
    }
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
pub fn enumerate_existing(
    table: &IDispatch,
    store_id: &str,
    options: RowOptions<'_>,
) -> Result<RowPage> {
    let columns = win::object_property(table, "Columns")?;
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
    let sender_smtp_available = win::invoke(
        &columns,
        "Add",
        DISPATCH_METHOD,
        vec![VARIANT::from(SENDER_SMTP)],
    )
    .is_ok();
    let [
        to_available,
        flag_available,
        importance_available,
        preview_available,
    ] = [DISPLAY_TO, FLAG_STATUS, IMPORTANCE, BODY_TEXT].map(|column| add_column(&columns, column));
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
        table,
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
        let end = bool::try_from(&win::property(table, "EndOfTable")?)
            .context("Outlook Table.EndOfTable returned an invalid value")?;
        if end {
            break;
        }
        scanned += 1;
        let row = win::method_object(table, "GetNextRow", vec![])?;
        let class = text(&row, "MessageClass");
        // Meeting items (Class 53) are readable via get_mail and the legacy
        // backend, so the table backend must list them too.
        if !class.starts_with("IPM.Note")
            && !class.starts_with("IPM.Post")
            && !class.starts_with("IPM.Schedule.Meeting")
        {
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
            let raw = text(&row, "SenderEmailAddress");
            let smtp =
                (sender_smtp_available && raw.starts_with('/')).then(|| text(&row, SENDER_SMTP));
            resolve_sender_address(&raw, smtp)
        } else {
            String::new()
        };
        if let Some(ref filter) = sender_filter
            && !sender_name.to_lowercase().contains(filter)
            && !sender_address.to_lowercase().contains(filter)
        {
            continue;
        }
        let subject = text(&row, "Subject");
        if let Some((words, scope)) = options.words {
            let haystack = match scope {
                WordScope::Subject => subject.to_lowercase(),
                WordScope::Sender => format!("{sender_name} {sender_address}").to_lowercase(),
            };
            if !has_all_words(&haystack, words) {
                continue;
            }
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
            "subject": subject,
            "from": sender_name,
            "from_address": public_address(&sender_address),
            "received": cell(&row, "ReceivedTime").as_ref().and_then(super::schedule::variant_iso),
            "unread": unread_state,
            "to": to_available.then(|| text(&row, DISPLAY_TO)),
            "flagged": flag_available
                .then(|| cell(&row, FLAG_STATUS).and_then(|v| i32::try_from(&v).ok()))
                .flatten()
                .map(|status| status == 2),
            "has_attachments": flags_available
                .then(|| message_flags(&row))
                .flatten()
                .map(|flags| flags & MSGFLAG_HASATTACH != 0),
            "importance": importance_available
                .then(|| cell(&row, IMPORTANCE).and_then(|v| i32::try_from(&v).ok()))
                .flatten(),
            "preview": preview_available.then(|| preview_text(&text(&row, BODY_TEXT)))
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
    fn x500_addresses_are_withheld() {
        assert_eq!(public_address("/O=EXCHANGELABS/CN=X"), "");
        assert_eq!(public_address("a@b.com"), "a@b.com");
    }
    #[test]
    fn exchange_senders_use_the_smtp_address_when_found() {
        let dn = "/O=EXCHANGELABS/CN=X";
        assert_eq!(
            resolve_sender_address(dn, Some("a@b.com".into())),
            "a@b.com"
        );
        // No SMTP, or a value that is not an address: keep the DN, which public_address withholds.
        assert_eq!(resolve_sender_address(dn, None), dn);
        assert_eq!(resolve_sender_address(dn, Some(String::new())), dn);
        // Plain SMTP senders are never replaced.
        assert_eq!(
            resolve_sender_address("a@b.com", Some("c@d.com".into())),
            "a@b.com"
        );
    }
    #[test]
    fn preview_is_one_line_and_capped() {
        assert_eq!(preview_text("  Hi\r\n\r\nthere \t you  "), "Hi there you");
        assert_eq!(
            preview_text(&"ab ".repeat(200)).chars().count(),
            PREVIEW_CHARS
        );
        assert_eq!(preview_text(""), "");
    }
    #[test]
    fn table_scan_cap_is_bounded() {
        assert_eq!(MAX_SCANNED, 10_000);
    }
}
