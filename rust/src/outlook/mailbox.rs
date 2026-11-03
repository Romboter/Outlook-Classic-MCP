//! Late-bound, read-only Outlook operations on the existing STA session.
//!
//! No COM interface leaves this thread. List/search materialize bounded,
//! metadata-only JSON results; bodies are read only by GetMail.
use super::{MailCommand, directory, indexed, schedule, table, win};
use crate::contracts::{
    FolderParams, GetMailParams, ListMailsParams, SearchMailsParams, checked_limit, escape_dasl,
};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::time::Instant;
use windows::{
    Win32::System::{
        Com::{DISPATCH_METHOD, IDispatch},
        Variant::{VAR_CHANGE_FLAGS, VARIANT, VT_BSTR, VariantChangeType},
    },
    core::BSTR,
};

pub(super) fn optional_text(object: &IDispatch, name: &str) -> String {
    win::string_property(object, name).unwrap_or_default()
}

fn text_property(object: &IDispatch, name: &str) -> String {
    // Outlook DATE properties come back as VT_DATE rather than VT_BSTR.
    // Coercion uses the current Windows locale. We label dates as local text,
    // not ISO-8601, until typed conversion has been tested on real Outlook.
    let Ok(value) = win::property(object, name) else {
        return String::new();
    };
    if let Ok(text) = BSTR::try_from(&value) {
        return text.to_string();
    }
    let mut converted = VARIANT::default();
    let result = unsafe { VariantChangeType(&mut converted, &value, VAR_CHANGE_FLAGS(0), VT_BSTR) };
    if result.is_err() {
        return String::new();
    }
    BSTR::try_from(&converted)
        .map(|s| s.to_string())
        .unwrap_or_default()
}

pub(super) fn bool_property(object: &IDispatch, name: &str) -> bool {
    let Ok(value) = win::property(object, name) else {
        return false;
    };
    bool::try_from(&value).unwrap_or(false)
}

pub(super) fn default_folder(namespace: &IDispatch, number: i32) -> Result<IDispatch> {
    win::method_object(namespace, "GetDefaultFolder", vec![VARIANT::from(number)])
}

fn folder_path(namespace: &IDispatch, path: Option<&str>) -> Result<IDispatch> {
    let path = path.unwrap_or("inbox").trim();
    let key = path.to_lowercase();
    let kind = match key.as_str() {
        "" | "inbox" => Some(6),
        "sent" | "sent items" => Some(5),
        "drafts" => Some(16),
        "deleted" | "deleted items" => Some(3),
        "outbox" => Some(4),
        _ => None,
    };
    if let Some(kind) = kind {
        return default_folder(namespace, kind);
    }

    // Relative folders in the default mailbox: Inbox/Projects/...
    // or a mailbox-root path like Projects/Customers.
    let segments: Vec<&str> = path
        .split(['/', '\\'])
        .filter(|s| !s.trim().is_empty())
        .collect();
    if segments.len() > 8 {
        bail!("folder path is too deep");
    }
    let mut current = if segments
        .first()
        .is_some_and(|s| s.eq_ignore_ascii_case("inbox"))
    {
        default_folder(namespace, 6)?
    } else {
        win::object_property(&default_folder(namespace, 6)?, "Parent")?
    };
    let rest = if segments
        .first()
        .is_some_and(|s| s.eq_ignore_ascii_case("inbox"))
    {
        &segments[1..]
    } else {
        &segments[..]
    };
    for part in rest {
        if part.len() > 120 || *part == ".." || *part == "." {
            bail!("invalid folder segment");
        }
        let folders = win::object_property(&current, "Folders")?;
        current = win::method_object(&folders, "Item", vec![VARIANT::from(*part)])
            .with_context(|| format!("folder segment not found: {part}"))?;
    }
    Ok(current)
}

fn item_summary(item: &IDispatch, store_id: &str) -> Value {
    json!({
        "entry_id": optional_text(item, "EntryID"),
        "store_id": store_id,
        "subject": optional_text(item, "Subject"),
        "from": optional_text(item, "SenderName"),
        "from_address": table::public_address(&optional_text(item, "SenderEmailAddress")),
        "received_local": text_property(item, "ReceivedTime"),
        "unread": bool_property(item, "UnRead")
    })
}

fn collect_items(
    items: &IDispatch,
    store_id: &str,
    limit: usize,
    offset: usize,
    unread_only: bool,
    sender_filter: Option<&str>,
) -> Result<(Vec<Value>, bool)> {
    let count = win::int_property(items, "Count")?.max(0) as usize;
    let mut hits = Vec::with_capacity(limit);
    let mut matched = 0usize;
    let sender_filter = sender_filter.map(str::to_lowercase);

    // Stop after one result beyond limit to report has_more accurately.
    // Each visited Item is a COM call; indexed table search is tracked in #5.
    for i in 1..=count {
        let item = match win::method_object(items, "Item", vec![VARIANT::from(i as i32)]) {
            Ok(item) => item,
            Err(_) => continue,
        };
        let class = win::int_property(&item, "Class").unwrap_or_default();
        if class != 43 && class != 53 {
            continue;
        }
        if unread_only && !bool_property(&item, "UnRead") {
            continue;
        }
        if let Some(ref filter) = sender_filter {
            let sender = format!(
                "{} {}",
                optional_text(&item, "SenderName"),
                optional_text(&item, "SenderEmailAddress")
            )
            .to_lowercase();
            if !sender.contains(filter) {
                continue;
            }
        }
        if matched < offset {
            matched += 1;
            continue;
        }
        if hits.len() >= limit {
            return Ok((hits, true));
        }
        hits.push(item_summary(&item, store_id));
    }
    Ok((hits, false))
}

fn sort_received(items: &IDispatch) -> Result<()> {
    win::invoke(
        items,
        "Sort",
        DISPATCH_METHOD,
        vec![VARIANT::from("[ReceivedTime]"), VARIANT::from(true)],
    )?;
    Ok(())
}

fn list_folders(namespace: &IDispatch, params: FolderParams) -> Result<Value> {
    let start = if params.root.as_deref().unwrap_or("").trim().is_empty() {
        win::object_property(&default_folder(namespace, 6)?, "Parent")?
    } else {
        folder_path(namespace, params.root.as_deref())?
    };
    let max_depth = params.max_depth.unwrap_or(2);
    let mut rows = Vec::new();
    let root_name = optional_text(&start, "Name");
    fn walk(
        folder: &IDispatch,
        name: String,
        depth: u32,
        max_depth: u32,
        rows: &mut Vec<Value>,
    ) -> Result<()> {
        if rows.len() >= 100 {
            return Ok(());
        }
        let count = win::object_property(folder, "Items")
            .ok()
            .and_then(|items| win::int_property(&items, "Count").ok());
        rows.push(json!({"name": optional_text(folder, "Name"),
            "path": name, "item_count": count}));
        if depth >= max_depth {
            return Ok(());
        }
        let folders = win::object_property(folder, "Folders")?;
        let count = win::int_property(&folders, "Count")?.clamp(0, 100);
        for i in 1..=count {
            if rows.len() >= 100 {
                break;
            }
            let child = win::method_object(&folders, "Item", vec![VARIANT::from(i)])?;
            let child_name = optional_text(&child, "Name");
            walk(
                &child,
                format!("{}/{}", name, child_name),
                depth + 1,
                max_depth,
                rows,
            )?;
        }
        Ok(())
    }
    walk(&start, root_name, 0, max_depth, &mut rows)?;
    Ok(json!({"count":rows.len(), "max_depth":max_depth,
        "truncated":rows.len() >= 100, "items":rows}))
}

fn list_mails(namespace: &IDispatch, params: ListMailsParams) -> Result<Value> {
    let started = Instant::now();
    let limit = checked_limit(params.limit).map_err(anyhow::Error::msg)?;
    let folder = folder_path(namespace, params.folder.as_deref())?;
    let store_id = optional_text(&folder, "StoreID");
    let folder_name = optional_text(&folder, "Name");
    let offset = params.offset.unwrap_or(0) as usize;
    let instant_search_enabled = table::store_instant_search(&folder);

    let backend = params.backend.as_deref().unwrap_or("auto");
    let queried = if backend == "legacy" {
        None
    } else {
        Some(table::enumerate(
            &folder,
            &store_id,
            table::RowOptions {
                filter: "",
                limit,
                offset,
                unread_only: params.unread_only.unwrap_or(false),
                sender_substring: params.from_address.as_deref(),
            },
        ))
    };

    match queried {
        Some(Ok(page)) => Ok(json!({
            "folder": folder_name,
            "count": page.items.len(), "limit": limit, "offset": offset,
            "has_more": page.has_more, "items": page.items,
            "backend": "folder_table", "scanned": page.scanned,
            "scan_truncated": page.scan_truncated,
            "instant_search_enabled": instant_search_enabled,
            "elapsed_ms": started.elapsed().as_millis()
        })),
        Some(Err(error)) if backend == "table" => Err(error),
        _ => {
            // Preserve function on stores that reject Table columns/filters.
            // Never imply that the fallback was fast or indexed.
            let items = win::object_property(&folder, "Items")?;
            sort_received(&items)?;
            let (rows, has_more) = collect_items(
                &items,
                &store_id,
                limit,
                offset,
                params.unread_only.unwrap_or(false),
                params.from_address.as_deref(),
            )?;
            Ok(json!({
                "folder": folder_name, "count": rows.len(), "limit": limit,
                "offset": offset, "has_more": has_more, "items": rows,
                "backend": if backend == "legacy" { "items_legacy" } else { "items_legacy_fallback" },
                "table_fallback": backend == "auto",
                "instant_search_enabled": instant_search_enabled,
                "elapsed_ms": started.elapsed().as_millis()
            }))
        }
    }
}

fn search_mails(namespace: &IDispatch, params: SearchMailsParams) -> Result<Value> {
    let started = Instant::now();
    let limit = checked_limit(params.limit).map_err(anyhow::Error::msg)?;
    let folder = folder_path(namespace, params.folder.as_deref())?;
    let store_id = optional_text(&folder, "StoreID");
    let folder_name = optional_text(&folder, "Name");
    let instant_search_enabled = table::store_instant_search(&folder);
    let safe_query = escape_dasl(params.query.trim());
    let scope = params.scope.as_deref().unwrap_or("subject_body");
    let backend = params.backend.as_deref().unwrap_or("auto");
    if backend == "indexed" {
        bail!("indexed search must be dispatched through the session application event sink");
    }
    if scope == "subject_body" && backend == "table" {
        bail!(
            "Folder.GetTable cannot reliably filter body text; use auto or legacy until indexed AdvancedSearch is implemented"
        );
    }

    // Only select subject/sender table filters whose properties are supported
    // for folder tables. Body-inclusive search stays on the compatibility
    // fallback until AdvancedSearch + Search.GetTable are implemented.
    let filter = match scope {
        "subject" => format!("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%'"),
        "from" => format!(
            "@SQL=(\"urn:schemas:httpmail:fromname\" LIKE '%{safe_query}%' OR \"urn:schemas:httpmail:fromemail\" LIKE '%{safe_query}%')"
        ),
        _ => format!(
            "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{safe_query}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{safe_query}%')"
        ),
    };

    if scope != "subject_body" && backend != "legacy" {
        // Folder.GetTable evaluates the filter before row enumeration.
        // It does not guarantee an indexed query even when Instant Search
        // is enabled. Report only observed elapsed time, never speed claims.
        match table::enumerate(
            &folder,
            &store_id,
            table::RowOptions {
                filter: &filter,
                limit,
                offset: 0,
                unread_only: false,
                sender_substring: None,
            },
        ) {
            Ok(page) => {
                return Ok(json!({
                    "query": params.query, "scope": scope, "folder": folder_name,
                    "count": page.items.len(), "has_more": page.has_more,
                    "items": page.items, "backend": "folder_table",
                    "scanned": page.scanned, "scan_truncated": page.scan_truncated,
                    "instant_search_enabled": instant_search_enabled,
                    "elapsed_ms": started.elapsed().as_millis()
                }));
            }
            Err(error) if backend == "table" => return Err(error),
            Err(_) => {}
        }
    }
    let items = win::object_property(&folder, "Items")?;
    let restricted = win::method_object(&items, "Restrict", vec![VARIANT::from(filter.as_str())])?;
    sort_received(&restricted)?;
    let (rows, has_more) = collect_items(&restricted, &store_id, limit, 0, false, None)?;
    Ok(json!({
        "query": params.query, "scope": scope, "folder": folder_name,
        "count": rows.len(), "has_more": has_more, "items": rows,
        "backend": "items_restrict_legacy",
        "table_fallback": backend == "auto" && scope != "subject_body",
        "instant_search_enabled": instant_search_enabled,
        "elapsed_ms": started.elapsed().as_millis()
    }))
}

fn get_mail(namespace: &IDispatch, params: GetMailParams) -> Result<Value> {
    let mut args = vec![VARIANT::from(params.entry_id.as_str())];
    if let Some(ref store) = params.store_id {
        if store.is_empty() || store.len() > 4096 {
            bail!("invalid store_id");
        }
        args.push(VARIANT::from(store.as_str()));
    }
    let item = win::method_object(namespace, "GetItemFromID", args)
        .context("Message no longer available; EntryID may be stale")?;
    let class = win::int_property(&item, "Class")?;
    if class != 43 && class != 53 {
        bail!("EntryID does not identify a mail item");
    }
    let limit = params.max_body_chars.unwrap_or(4000) as usize;
    let mut data = item_summary(&item, params.store_id.as_deref().unwrap_or(""));
    let object = data.as_object_mut().context("invalid mail result")?;
    object.insert("to".into(), json!(optional_text(&item, "To")));
    object.insert("cc".into(), json!(optional_text(&item, "CC")));
    if params.include_body.unwrap_or(true) {
        let body = optional_text(&item, "Body");
        let total_chars = body.chars().count();
        object.insert(
            "body".into(),
            json!(body.chars().take(limit).collect::<String>()),
        );
        object.insert("body_truncated".into(), json!(total_chars > limit));
        object.insert("body_total_chars".into(), json!(total_chars));
    }
    Ok(data)
}

impl win::Session {
    pub fn execute(&mut self, command: MailCommand) -> Result<Value> {
        let mut response = match command {
            MailCommand::WhoAmI => {
                let recipient = win::object_property(&self.namespace, "CurrentUser")?;
                Ok(json!({"current_user": win::string_property(&recipient, "Name")?}))
            }
            MailCommand::ListEvents(params) => schedule::list_events(&self.namespace, params),
            MailCommand::GetEvent(mut params) => {
                let mut shim = GetMailParams {
                    entry_id: params.entry_id.clone(),
                    store_id: params.store_id.clone(),
                    include_body: None,
                    include_html: None,
                    max_body_chars: None,
                };
                self.refs.resolve_input(&mut shim)?;
                params.entry_id = shim.entry_id;
                params.store_id = shim.store_id;
                schedule::get_event(&self.namespace, params)
            }
            MailCommand::ListTasks(params) => schedule::list_tasks(&self.namespace, params),
            MailCommand::ListContacts(params) => directory::list_contacts(&self.namespace, params),
            MailCommand::SearchContacts(params) => {
                directory::search_contacts(&self.namespace, params)
            }
            MailCommand::GetContact(mut params) => {
                // Contact handles share the mail handle table.
                let mut shim = GetMailParams {
                    entry_id: params.entry_id.clone(),
                    store_id: params.store_id.clone(),
                    include_body: None,
                    include_html: None,
                    max_body_chars: None,
                };
                self.refs.resolve_input(&mut shim)?;
                params.entry_id = shim.entry_id;
                params.store_id = shim.store_id;
                directory::get_contact(&self.namespace, params)
            }
            MailCommand::ResolveName(params) => directory::resolve_name(&self.namespace, params),
            MailCommand::ListCategories => directory::list_categories(&self.namespace),
            MailCommand::ListRules => directory::list_rules(&self.namespace),
            MailCommand::GetOutOfOffice => directory::get_out_of_office(&self.namespace),
            MailCommand::ListFolders(params) => list_folders(&self.namespace, params),
            MailCommand::ListMails(params) => list_mails(&self.namespace, params),
            MailCommand::SearchMails(params) if params.backend.as_deref() == Some("indexed") => {
                let folder = folder_path(&self.namespace, params.folder.as_deref())?;
                let store_id = optional_text(&folder, "StoreID");
                indexed::search(&self._application, &folder, &params, &store_id)
            }
            MailCommand::SearchMails(params) => search_mails(&self.namespace, params),
            MailCommand::GetThread(params) => self.read_thread(params),
            MailCommand::GetMail(mut params) => {
                self.refs.resolve_input(&mut params)?;
                get_mail(&self.namespace, params)
            }
        }?;
        self.refs.compact_response(&mut response);
        Ok(response)
    }
}
