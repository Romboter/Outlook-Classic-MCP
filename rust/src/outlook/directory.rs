//! Read-only contacts, directory, categories, rules and out-of-office reads.
//!
//! Every scan is bounded (items visited, folders visited, wall-clock for the
//! Global Address List) so one call cannot outlast the bridge timeout.
use super::mailbox::{bool_property, optional_text};
use super::win;
use crate::contracts::{
    GetContactParams, ListContactsParams, ResolveNameParams, SearchContactsParams, checked_limit,
    matches_all, query_words,
};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::time::{Duration, Instant};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const OL_CLASS_CONTACT: i32 = 40;
const OL_CONTACT_ITEM: i32 = 2;
const MAX_FOLDERS: usize = 200;
const MAX_ITEMS_SCANNED: usize = 5000;
const MAX_GAL_ENTRIES: i32 = 50_000;
/// Folder discovery plus item scan; the bridge itself times out at 45 s.
const CONTACT_BUDGET: Duration = Duration::from_secs(15);
/// Overall ceiling including the Global Address List scan.
const CALL_BUDGET: Duration = Duration::from_secs(38);
const NOTES_CHARS: usize = 4000;
const OOF_PROPTAG: &str = "http://schemas.microsoft.com/mapi/proptag/0x661D000B";

fn nth(collection: &IDispatch, index: i32) -> Result<IDispatch> {
    win::method_object(collection, "Item", vec![VARIANT::from(index)])
}

/// Stores with the default mailbox first, so a budget cut-off loses the
/// least likely places (shared mailboxes, archives) rather than the user's own.
fn ordered_stores(namespace: &IDispatch) -> Result<Vec<IDispatch>> {
    let default_id = win::object_property(namespace, "DefaultStore")
        .map(|s| optional_text(&s, "StoreID"))
        .unwrap_or_default();
    let stores = win::object_property(namespace, "Stores")?;
    let count = win::int_property(&stores, "Count")?.clamp(0, 50);
    let (mut first, mut rest) = (Vec::new(), Vec::new());
    for i in 1..=count {
        let Ok(store) = nth(&stores, i) else { continue };
        if !default_id.is_empty() && optional_text(&store, "StoreID") == default_id {
            first.push(store);
        } else {
            rest.push(store);
        }
    }
    first.extend(rest);
    Ok(first)
}

/// Contact folders of one store as `(folder, display path)`. The second value
/// is true when discovery stopped at the deadline.
fn store_contact_folders(store: &IDispatch, deadline: Instant) -> (Vec<(IDispatch, String)>, bool) {
    fn walk(
        folder: IDispatch,
        path: String,
        depth: u32,
        deadline: Instant,
        out: &mut Vec<(IDispatch, String)>,
    ) -> bool {
        if Instant::now() > deadline || out.len() >= MAX_FOLDERS {
            return true;
        }
        if win::int_property(&folder, "DefaultItemType").ok() == Some(OL_CONTACT_ITEM) {
            out.push((folder.clone(), path.clone()));
        }
        if depth >= 3 {
            return false;
        }
        let Ok(subs) = win::object_property(&folder, "Folders") else {
            return false;
        };
        let count = win::int_property(&subs, "Count").unwrap_or(0).clamp(0, 100);
        for i in 1..=count {
            if let Ok(sub) = nth(&subs, i) {
                let name = optional_text(&sub, "Name");
                if walk(sub, format!("{path}/{name}"), depth + 1, deadline, out) {
                    return true;
                }
            }
        }
        false
    }

    let mut out = Vec::new();
    let Ok(root) = win::method_object(store, "GetRootFolder", vec![]) else {
        return (out, false);
    };
    let cut = walk(
        root,
        optional_text(store, "DisplayName"),
        0,
        deadline,
        &mut out,
    );
    (out, cut)
}

fn contact_summary(contact: &IDispatch, store_id: &str) -> Value {
    json!({
        "entry_id": optional_text(contact, "EntryID"),
        "store_id": store_id,
        "full_name": optional_text(contact, "FullName"),
        "email": optional_text(contact, "Email1Address"),
        "company": optional_text(contact, "CompanyName"),
        "job_title": optional_text(contact, "JobTitle"),
        "mobile": optional_text(contact, "MobileTelephoneNumber"),
    })
}

fn is_contact(item: &IDispatch) -> bool {
    win::int_property(item, "Class").ok() == Some(OL_CLASS_CONTACT)
}

pub(super) fn list_contacts(namespace: &IDispatch, params: ListContactsParams) -> Result<Value> {
    let limit = checked_limit(params.limit).map_err(anyhow::Error::msg)?;
    let offset = params.offset.unwrap_or(0) as usize;
    let deadline = Instant::now() + CONTACT_BUDGET;
    let (mut rows, mut skipped, mut scanned) = (Vec::new(), 0usize, 0usize);
    let (mut has_more, mut scan_truncated) = (false, false);

    'stores: for store in ordered_stores(namespace)? {
        let (folders, cut) = store_contact_folders(&store, deadline);
        scan_truncated |= cut;
        for (folder, path) in folders {
            let Ok(items) = win::object_property(&folder, "Items") else {
                continue;
            };
            let _ = win::invoke(
                &items,
                "Sort",
                DISPATCH_METHOD,
                vec![VARIANT::from("[FullName]")],
            );
            let store_id = optional_text(&folder, "StoreID");
            let count = win::int_property(&items, "Count").unwrap_or(0).max(0);
            for i in 1..=count {
                if scanned >= MAX_ITEMS_SCANNED || Instant::now() > deadline {
                    scan_truncated = true;
                    break 'stores;
                }
                scanned += 1;
                let Ok(contact) = nth(&items, i) else {
                    continue;
                };
                if !is_contact(&contact) {
                    continue;
                }
                if skipped < offset {
                    skipped += 1;
                    continue;
                }
                if rows.len() >= limit {
                    has_more = true;
                    break 'stores;
                }
                let mut row = contact_summary(&contact, &store_id);
                row["folder"] = json!(path);
                rows.push(row);
            }
        }
        if Instant::now() > deadline {
            scan_truncated = true;
            break;
        }
    }
    Ok(
        json!({"count": rows.len(), "offset": offset, "has_more": has_more,
        "scan_truncated": scan_truncated, "items": rows}),
    )
}

fn search_directory(
    namespace: &IDispatch,
    words: &[String],
    max: usize,
    deadline: Instant,
) -> Result<(Vec<Value>, bool)> {
    // No Exchange directory on the profile is normal, not an error.
    let Ok(gal) = win::method_object(namespace, "GetGlobalAddressList", vec![]) else {
        return Ok((vec![], false));
    };
    let entries = win::object_property(&gal, "AddressEntries")?;
    let count = win::int_property(&entries, "Count")?.clamp(0, MAX_GAL_ENTRIES);
    let mut rows = Vec::new();
    for i in 1..=count {
        if Instant::now() > deadline {
            return Ok((rows, true));
        }
        let Ok(entry) = nth(&entries, i) else {
            continue;
        };
        let name = optional_text(&entry, "Name");
        if !matches_all(&name, words) {
            continue;
        }
        let user = win::method_object(&entry, "GetExchangeUser", vec![]).ok();
        let field = |p: &str| {
            user.as_ref()
                .map(|u| optional_text(u, p))
                .unwrap_or_default()
        };
        rows.push(json!({
            "full_name": name,
            "email": field("PrimarySmtpAddress"),
            "company": field("CompanyName"),
            "job_title": field("JobTitle"),
            "mobile": field("MobileTelephoneNumber"),
            "source": "directory",
        }));
        if rows.len() >= max {
            break;
        }
    }
    Ok((rows, count == MAX_GAL_ENTRIES))
}

pub(super) fn search_contacts(
    namespace: &IDispatch,
    params: SearchContactsParams,
) -> Result<Value> {
    let limit = checked_limit(params.limit).map_err(anyhow::Error::msg)?;
    let words = query_words(&params.query);
    let started = Instant::now();
    let deadline = started + CONTACT_BUDGET;
    let (mut rows, mut scanned, mut scan_truncated) = (Vec::<Value>::new(), 0usize, false);

    'stores: for store in ordered_stores(namespace)? {
        let (folders, cut) = store_contact_folders(&store, deadline);
        scan_truncated |= cut;
        for (folder, path) in folders {
            let Ok(items) = win::object_property(&folder, "Items") else {
                continue;
            };
            let store_id = optional_text(&folder, "StoreID");
            let count = win::int_property(&items, "Count").unwrap_or(0).max(0);
            for i in 1..=count {
                if scanned >= MAX_ITEMS_SCANNED || Instant::now() > deadline {
                    scan_truncated = true;
                    break 'stores;
                }
                scanned += 1;
                let Ok(contact) = nth(&items, i) else {
                    continue;
                };
                if !is_contact(&contact) {
                    continue;
                }
                let haystack = ["FullName", "Email1Address", "CompanyName", "JobTitle"]
                    .iter()
                    .map(|p| optional_text(&contact, p))
                    .collect::<Vec<_>>()
                    .join(" ");
                if matches_all(&haystack, &words) {
                    let mut row = contact_summary(&contact, &store_id);
                    row["folder"] = json!(path);
                    row["source"] = json!("contacts");
                    rows.push(row);
                    if rows.len() >= limit {
                        break 'stores;
                    }
                }
            }
        }
        if Instant::now() > deadline {
            scan_truncated = true;
            break;
        }
    }

    let searched_directory = params.include_directory.unwrap_or(true) && rows.len() < limit;
    let mut directory_truncated = false;
    if searched_directory {
        let (found, truncated) =
            search_directory(namespace, &words, limit - rows.len(), started + CALL_BUDGET)?;
        directory_truncated = truncated;
        let seen: Vec<String> = rows
            .iter()
            .filter_map(|r| r["email"].as_str())
            .map(str::to_lowercase)
            .collect();
        rows.extend(found.into_iter().filter(|r| {
            r["email"]
                .as_str()
                .is_none_or(|e| e.is_empty() || !seen.contains(&e.to_lowercase()))
        }));
    }
    Ok(
        json!({"query": params.query, "count": rows.len(), "items": rows,
        "searched_directory": searched_directory,
        "scan_truncated": scan_truncated || directory_truncated}),
    )
}

pub(super) fn get_contact(namespace: &IDispatch, params: GetContactParams) -> Result<Value> {
    let mut args = vec![VARIANT::from(params.entry_id.as_str())];
    if let Some(ref store) = params.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    let contact = win::method_object(namespace, "GetItemFromID", args)
        .context("Contact no longer available; EntryID may be stale")?;
    if !is_contact(&contact) {
        bail!("EntryID does not identify a contact");
    }
    let mut data = contact_summary(&contact, params.store_id.as_deref().unwrap_or(""));
    let notes = optional_text(&contact, "Body");
    let object = data.as_object_mut().context("invalid contact result")?;
    for (key, prop) in [
        ("business_phone", "BusinessTelephoneNumber"),
        ("home_phone", "HomeTelephoneNumber"),
        ("address", "BusinessAddress"),
    ] {
        object.insert(key.into(), json!(optional_text(&contact, prop)));
    }
    object.insert(
        "notes_truncated".into(),
        json!(notes.chars().count() > NOTES_CHARS),
    );
    object.insert(
        "notes".into(),
        json!(notes.chars().take(NOTES_CHARS).collect::<String>()),
    );
    Ok(data)
}

pub(super) fn resolve_name(namespace: &IDispatch, params: ResolveNameParams) -> Result<Value> {
    let recipient = win::method_object(
        namespace,
        "CreateRecipient",
        vec![VARIANT::from(params.name.as_str())],
    )?;
    win::invoke(&recipient, "Resolve", DISPATCH_METHOD, vec![])?;
    if !bool_property(&recipient, "Resolved") {
        return Ok(json!({"resolved": false, "query": params.name,
            "note": "No unique match. Try the full display name or exact address, or use outlook_search_contacts."}));
    }
    let entry = win::object_property(&recipient, "AddressEntry")?;
    let mut smtp = optional_text(&entry, "Address");
    if let Ok(user) = win::method_object(&entry, "GetExchangeUser", vec![]) {
        let primary = optional_text(&user, "PrimarySmtpAddress");
        if !primary.is_empty() {
            smtp = primary;
        }
    }
    Ok(json!({"resolved": true, "query": params.name,
        "display_name": optional_text(&entry, "Name"), "smtp_address": smtp}))
}

pub(super) fn list_categories(namespace: &IDispatch) -> Result<Value> {
    let categories = win::object_property(namespace, "Categories")?;
    let count = win::int_property(&categories, "Count")?.clamp(0, 500);
    let mut rows = Vec::new();
    for i in 1..=count {
        let category = nth(&categories, i)?;
        rows.push(json!({"name": optional_text(&category, "Name"),
            "color": win::int_property(&category, "Color").unwrap_or(0)}));
    }
    Ok(json!({"count": rows.len(), "items": rows}))
}

pub(super) fn list_rules(namespace: &IDispatch) -> Result<Value> {
    let store = win::object_property(namespace, "DefaultStore")?;
    let rules = win::method_object(&store, "GetRules", vec![])?;
    let count = win::int_property(&rules, "Count")?.clamp(0, 500);
    let mut rows = Vec::new();
    for i in 1..=count {
        let rule = nth(&rules, i)?;
        rows.push(json!({"index": i, "name": optional_text(&rule, "Name"),
            "enabled": bool_property(&rule, "Enabled")}));
    }
    Ok(json!({"count": rows.len(), "items": rows}))
}

pub(super) fn get_out_of_office(namespace: &IDispatch) -> Result<Value> {
    let state = win::object_property(namespace, "DefaultStore")
        .and_then(|store| win::object_property(&store, "PropertyAccessor"))
        .and_then(|accessor| {
            win::invoke(
                &accessor,
                "GetProperty",
                DISPATCH_METHOD,
                vec![VARIANT::from(OOF_PROPTAG)],
            )
        })
        .ok()
        .and_then(|v| {
            bool::try_from(&v)
                .ok()
                .or_else(|| i32::try_from(&v).ok().map(|n| n != 0))
        });
    Ok(match state {
        Some(on) => json!({"out_of_office": on, "status": if on { "on" } else { "off" }}),
        None => json!({"out_of_office": null, "status": "unknown",
            "note": "Could not read OOF property. Check Outlook File > Automatic Replies directly."}),
    })
}
