//! Late-bound, read-only Outlook operations on the existing STA session.
//!
//! No COM interface leaves this thread. List/search materialize bounded,
//! metadata-only JSON results; bodies are read only by GetMail.
use super::{
    MailCommand, attachments, cancel, compose, delete, directory, indexed, meeting, modify,
    respond, rule, schedule, schedule_write, table, win,
};
use crate::action::ActionKind;
use crate::contracts::{
    FolderParams, GetMailParams, ListMailsParams, SearchMailsParams, checked_limit, escape_dasl,
    has_all_words, split_search_words,
};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::time::Instant;
use windows::{
    Win32::System::{
        Com::{DISPATCH_METHOD, IDispatch},
        Variant::VARIANT,
    },
    core::BSTR,
};

pub(super) fn optional_text(object: &IDispatch, name: &str) -> String {
    win::string_property(object, name).unwrap_or_default()
}

/// Sender address, as SMTP where Outlook can give one. Exchange-internal senders
/// report an X.500 DN in `SenderEmailAddress`, so ask for PR_SENDER_SMTP_ADDRESS then.
fn sender_address(item: &IDispatch) -> String {
    let raw = optional_text(item, "SenderEmailAddress");
    let smtp = raw
        .starts_with('/')
        .then(|| {
            let accessor = win::object_property(item, "PropertyAccessor").ok()?;
            let args = vec![VARIANT::from(table::SENDER_SMTP)];
            let value = win::invoke(&accessor, "GetProperty", DISPATCH_METHOD, args).ok()?;
            BSTR::try_from(&value).ok().map(|s| s.to_string())
        })
        .flatten();
    table::public_address(&table::resolve_sender_address(&raw, smtp))
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

/// Well-known folder names (default store) to `OlDefaultFolders` numbers.
fn default_folder_kind(path: &str) -> Option<i32> {
    match path.trim().to_lowercase().as_str() {
        "" | "inbox" => Some(6),
        "sent" | "sent items" => Some(5),
        "drafts" => Some(16),
        "deleted" | "deleted items" | "trash" => Some(3),
        "outbox" => Some(4),
        "junk" | "spam" | "junk email" | "junk e-mail" => Some(23),
        "calendar" => Some(9),
        "contacts" => Some(10),
        "tasks" => Some(13),
        "notes" => Some(12),
        _ => None,
    }
}

/// Root folder of the store whose display name equals `name` (any store, case-insensitive).
fn store_root(namespace: &IDispatch, name: &str) -> Option<IDispatch> {
    let roots = win::object_property(namespace, "Folders").ok()?;
    let count = win::property(&roots, "Count")
        .ok()
        .and_then(|v| i32::try_from(&v).ok())?;
    (1..=count.min(50)).find_map(|i| {
        let root = win::method_object(&roots, "Item", vec![VARIANT::from(i)]).ok()?;
        optional_text(&root, "Name")
            .eq_ignore_ascii_case(name)
            .then_some(root)
    })
}

fn walk_from_store(root: IDispatch, rest: &[&str]) -> Result<IDispatch> {
    let mut current = root.clone();
    for (i, part) in rest.iter().enumerate() {
        if part.len() > 120 || *part == ".." || *part == "." {
            bail!("invalid folder segment");
        }
        // "Store/Inbox" and friends: the store's own default folder.
        if i == 0
            && let Some(kind) = default_folder_kind(part).filter(|_| !part.is_empty())
            && let Ok(store) = win::object_property(&root, "Store")
            && let Ok(folder) =
                win::method_object(&store, "GetDefaultFolder", vec![VARIANT::from(kind)])
        {
            current = folder;
            continue;
        }
        let folders = win::object_property(&current, "Folders")?;
        current = win::method_object(&folders, "Item", vec![VARIANT::from(*part)])
            .with_context(|| format!("folder segment not found: {part}"))?;
    }
    Ok(current)
}

pub(super) fn folder_path(namespace: &IDispatch, path: Option<&str>) -> Result<IDispatch> {
    let path = path.unwrap_or("inbox").trim();
    if let Some(kind) = default_folder_kind(path) {
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
    // A leading segment naming another store (archive, PST, shared mailbox)
    // starts the walk at that store's root.
    if let Some((first, rest)) = segments.split_first()
        && let Some(root) = store_root(namespace, first)
    {
        return walk_from_store(root, rest);
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
        "from_address": sender_address(item),
        "received": schedule::date_iso(item, "ReceivedTime"),
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
    words: Option<(&[String], &str)>,
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
                sender_address(&item)
            )
            .to_lowercase();
            if !sender.contains(filter) {
                continue;
            }
        }
        if let Some((words, scope)) = words
            && !has_all_words(&item_haystack(&item, scope), words)
        {
            continue;
        }
        if matched < offset {
            matched += 1;
            continue;
        }
        if hits.len() >= limit {
            return Ok((hits, true));
        }
        let mut row = item_summary(&item, store_id);
        add_triage_fields(&item, &mut row);
        hits.push(row);
    }
    Ok((hits, false))
}

/// Lowercased text the search `scope` looks in, like Python's `_search_haystack`.
fn item_haystack(item: &IDispatch, scope: &str) -> String {
    let text = match scope {
        "subject" => optional_text(item, "Subject"),
        "from" => format!(
            "{} {}",
            optional_text(item, "SenderName"),
            sender_address(item)
        ),
        _ => format!(
            "{} {}",
            optional_text(item, "Subject"),
            optional_text(item, "Body")
        ),
    };
    text.to_lowercase()
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
            "path": name, "item_count": count,
            "unread_count": win::int_property(folder, "UnReadItemCount").unwrap_or(0),
            "default_item_type": win::int_property(folder, "DefaultItemType").unwrap_or(-1)}));
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
    let date_filter =
        crate::contracts::received_filter(params.since.as_deref(), params.until.as_deref())
            .map_err(anyhow::Error::msg)?;
    let queried = if backend == "legacy" {
        None
    } else {
        Some(table::enumerate(
            &folder,
            &store_id,
            table::RowOptions {
                filter: &date_filter,
                limit,
                offset,
                unread_only: params.unread_only.unwrap_or(false),
                sender_substring: params.from_address.as_deref(),
                words: None,
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
            let mut items = win::object_property(&folder, "Items")?;
            if !date_filter.is_empty() {
                items = win::method_object(
                    &items,
                    "Restrict",
                    vec![VARIANT::from(date_filter.as_str())],
                )?;
            }
            sort_received(&items)?;
            let (rows, has_more) = collect_items(
                &items,
                &store_id,
                limit,
                offset,
                params.unread_only.unwrap_or(false),
                params.from_address.as_deref(),
                None,
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
    // Several words: filter on the most selective one, then require the rest per row/item.
    let (anchor, remaining) = split_search_words(params.query.trim());
    let safe_query = escape_dasl(&anchor);
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
                words: (!remaining.is_empty()).then_some((
                    remaining.as_slice(),
                    if scope == "from" {
                        table::WordScope::Sender
                    } else {
                        table::WordScope::Subject
                    },
                )),
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
    let (rows, has_more) = collect_items(
        &restricted,
        &store_id,
        limit,
        0,
        false,
        None,
        (!remaining.is_empty()).then_some((remaining.as_slice(), scope)),
    )?;
    Ok(json!({
        "query": params.query, "scope": scope, "folder": folder_name,
        "count": rows.len(), "has_more": has_more, "items": rows,
        "backend": "items_restrict_legacy",
        "table_fallback": backend == "auto" && scope != "subject_body",
        "instant_search_enabled": instant_search_enabled,
        "elapsed_ms": started.elapsed().as_millis()
    }))
}

/// Triage fields list and search rows carry besides the summary: `to`, `flagged`,
/// `has_attachments`, `importance` and a one-line `preview`.
fn add_triage_fields(item: &IDispatch, row: &mut Value) {
    let attachments = win::object_property(item, "Attachments")
        .ok()
        .and_then(|a| win::int_property(&a, "Count").ok());
    row["to"] = json!(optional_text(item, "To"));
    row["flagged"] = json!(is_flagged(item));
    row["has_attachments"] = json!(attachments.map(|n| n > 0));
    row["importance"] = json!(win::int_property(item, "Importance").ok());
    row["preview"] = json!(table::preview_text(&optional_text(item, "Body")));
}

/// `OlFlagStatus.olFlagMarked`.
fn is_flagged(item: &IDispatch) -> bool {
    win::int_property(item, "FlagStatus").ok() == Some(2)
}

/// 1-indexed attachment list, the index `outlook_save_attachments` takes. Bounded like it.
fn attachment_list(item: &IDispatch) -> Vec<Value> {
    let Ok(collection) = win::object_property(item, "Attachments") else {
        return Vec::new();
    };
    let count = win::int_property(&collection, "Count")
        .unwrap_or(0)
        .clamp(0, 100);
    (1..=count)
        .filter_map(|i| {
            let attachment =
                win::method_object(&collection, "Item", vec![VARIANT::from(i)]).ok()?;
            Some(json!({
                "index": i,
                "filename": optional_text(&attachment, "FileName"),
                "size_bytes": win::int_property(&attachment, "Size").ok(),
            }))
        })
        .collect()
}

fn account_list(application: &IDispatch) -> Vec<Value> {
    let Ok(accounts) = win::object_property(application, "Session")
        .and_then(|session| win::object_property(&session, "Accounts"))
    else {
        return Vec::new();
    };
    let count = win::int_property(&accounts, "Count")
        .unwrap_or(0)
        .clamp(0, 50);
    (1..=count)
        .filter_map(|i| {
            let account = win::method_object(&accounts, "Item", vec![VARIANT::from(i)]).ok()?;
            Some(json!({
                "display_name": optional_text(&account, "DisplayName"),
                "smtp_address": optional_text(&account, "SmtpAddress"),
                "user_name": optional_text(&account, "UserName"),
                "account_type": win::int_property(&account, "AccountType").ok(),
            }))
        })
        .collect()
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
    object.insert("bcc".into(), json!(optional_text(&item, "BCC")));
    object.insert("sent".into(), json!(schedule::date_iso(&item, "SentOn")));
    object.insert("flagged".into(), json!(is_flagged(&item)));
    object.insert(
        "importance".into(),
        json!(win::int_property(&item, "Importance").ok()),
    );
    object.insert(
        "categories".into(),
        json!(optional_text(&item, "Categories")),
    );
    object.insert(
        "conversation_id".into(),
        json!(optional_text(&item, "ConversationID")),
    );
    object.insert("attachments".into(), json!(attachment_list(&item)));
    if params.include_body.unwrap_or(true) {
        let body = optional_text(&item, "Body");
        let total_chars = body.chars().count();
        let offset = params.body_offset.unwrap_or(0) as usize;
        object.insert(
            "body".into(),
            json!(body.chars().skip(offset).take(limit).collect::<String>()),
        );
        let end = offset.saturating_add(limit);
        object.insert("body_truncated".into(), json!(total_chars > end));
        if total_chars > end {
            object.insert("body_next_offset".into(), json!(end));
        }
        object.insert("body_total_chars".into(), json!(total_chars));
    }
    Ok(data)
}

impl win::Session {
    /// Turn a short session handle into the raw EntryID/StoreID pair.
    fn resolve_ids(
        &self,
        entry_id: &str,
        store_id: &Option<String>,
    ) -> Result<(String, Option<String>)> {
        let mut shim = GetMailParams {
            entry_id: entry_id.to_string(),
            store_id: store_id.clone(),
            include_body: None,
            include_html: None,
            max_body_chars: None,
            body_offset: None,
        };
        self.refs.resolve_input(&mut shim)?;
        Ok((shim.entry_id, shim.store_id))
    }

    pub fn execute(&mut self, command: MailCommand) -> Result<Value> {
        let mut response = match command {
            MailCommand::WhoAmI => {
                let recipient = win::object_property(&self.namespace, "CurrentUser")?;
                let (local_time, timezone, utc_offset) = schedule::clock_info();
                Ok(json!({
                    "current_user": win::string_property(&recipient, "Name")?,
                    "accounts": account_list(&self._application),
                    "local_time": local_time,
                    "timezone": timezone,
                    "utc_offset": utc_offset
                }))
            }
            MailCommand::CreateTask(params) => {
                schedule_write::create_task(&self._application, &self.namespace, params)
            }
            MailCommand::CompleteTask(mut params) => {
                (params.entry_id, params.store_id) =
                    self.resolve_ids(&params.entry_id, &params.store_id)?;
                schedule_write::complete_task(&self.namespace, params)
            }
            MailCommand::CreateEvent(params) => {
                schedule_write::create_event(&self._application, &self.namespace, params)
            }
            MailCommand::UpdateEvent(mut params) => {
                (params.entry_id, params.store_id) =
                    self.resolve_ids(&params.entry_id, &params.store_id)?;
                schedule_write::update_event(&self.namespace, params)
            }
            MailCommand::ActionPrepare(mut target)
                if target.kind == ActionKind::SaveAttachments =>
            {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                attachments::prepare(&self.namespace, &target)
            }
            MailCommand::ActionExecute {
                mut target,
                preview,
            } if target.kind == ActionKind::SaveAttachments => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                attachments::execute(&self.namespace, &target, &preview)
            }
            // Saving a file leaves no mark in Outlook to look for afterwards.
            MailCommand::ActionLocate { target, .. }
                if target
                    .as_ref()
                    .is_some_and(|t| t.kind == ActionKind::SaveAttachments) =>
            {
                Ok(serde_json::json!({"located": "unknown"}))
            }
            MailCommand::ActionPrepare(target) if target.kind == ActionKind::InviteMeeting => {
                meeting::prepare_invite(&self._application, &self.namespace, &target)
            }
            MailCommand::ActionExecute { target, preview }
                if target.kind == ActionKind::InviteMeeting =>
            {
                meeting::execute_invite(&self._application, &target, &preview)
            }
            MailCommand::ActionLocate { action_id, target }
                if target
                    .as_ref()
                    .is_some_and(|t| t.kind == ActionKind::InviteMeeting) =>
            {
                meeting::locate_invite(&self.namespace, &action_id)
            }
            MailCommand::ActionPrepare(mut target) if target.kind == ActionKind::UpdateMeeting => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                meeting::prepare_update(&self.namespace, &target)
            }
            MailCommand::ActionExecute {
                mut target,
                preview,
            } if target.kind == ActionKind::UpdateMeeting => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                meeting::execute_update(&self.namespace, &target, &preview)
            }
            MailCommand::ActionLocate { mut target, .. }
                if target
                    .as_ref()
                    .is_some_and(|t| t.kind == ActionKind::UpdateMeeting) =>
            {
                if let Some(t) = target.as_mut() {
                    (t.entry_id, t.store_id) = self.resolve_ids(&t.entry_id, &t.store_id)?;
                }
                meeting::locate_update(&self.namespace, target.as_ref())
            }
            MailCommand::ActionPrepare(mut target) if target.kind == ActionKind::CancelMeeting => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                cancel::prepare(&self.namespace, &target)
            }
            MailCommand::ActionExecute {
                mut target,
                preview,
            } if target.kind == ActionKind::CancelMeeting => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                cancel::execute(&self.namespace, &target, &preview)
            }
            MailCommand::ActionLocate {
                action_id,
                mut target,
            } if target
                .as_ref()
                .is_some_and(|t| t.kind == ActionKind::CancelMeeting) =>
            {
                if let Some(t) = target.as_mut() {
                    (t.entry_id, t.store_id) = self.resolve_ids(&t.entry_id, &t.store_id)?;
                }
                cancel::locate(&self.namespace, &action_id, target.as_ref())
            }
            MailCommand::ActionPrepare(mut target) if target.kind == ActionKind::RespondEvent => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                respond::prepare(&self.namespace, &target)
            }
            MailCommand::ActionExecute {
                mut target,
                preview,
            } if target.kind == ActionKind::RespondEvent => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                respond::execute(&self.namespace, &target, &preview)
            }
            MailCommand::ActionLocate {
                action_id,
                mut target,
            } if target
                .as_ref()
                .is_some_and(|t| t.kind == ActionKind::RespondEvent) =>
            {
                if let Some(t) = target.as_mut() {
                    (t.entry_id, t.store_id) = self.resolve_ids(&t.entry_id, &t.store_id)?;
                }
                respond::locate(&self.namespace, &action_id, target.as_ref())
            }
            // Rules are named, not addressed by item id, so nothing is resolved.
            MailCommand::ActionPrepare(target) if target.kind == ActionKind::ToggleRule => {
                rule::prepare(&self.namespace, &target)
            }
            MailCommand::ActionExecute { target, preview }
                if target.kind == ActionKind::ToggleRule =>
            {
                rule::execute(&self.namespace, &target, &preview)
            }
            MailCommand::ActionLocate { target, .. }
                if target
                    .as_ref()
                    .is_some_and(|t| t.kind == ActionKind::ToggleRule) =>
            {
                rule::locate(&self.namespace, target.as_ref())
            }
            MailCommand::ActionPrepare(mut target) => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                delete::prepare(&self.namespace, &target)
            }
            MailCommand::ActionExecute {
                mut target,
                preview,
            } => {
                (target.entry_id, target.store_id) =
                    self.resolve_ids(&target.entry_id, &target.store_id)?;
                delete::execute(&self.namespace, &target, &preview)
            }
            MailCommand::ActionLocate {
                action_id,
                mut target,
            } => {
                if let Some(t) = target.as_mut() {
                    (t.entry_id, t.store_id) = self.resolve_ids(&t.entry_id, &t.store_id)?;
                }
                delete::locate(&self.namespace, &action_id, target.as_ref())
            }
            MailCommand::ComposePrepare(request) => {
                let this = &*self;
                compose::prepare(&this._application, &this.namespace, request, |d| {
                    let (entry_id, store_id) = this.resolve_ids(&d.entry_id, &d.store_id)?;
                    Ok(crate::send::DraftRef { entry_id, store_id })
                })
            }
            MailCommand::ComposeInspect(draft) => {
                let (entry_id, store_id) = self.resolve_ids(&draft.entry_id, &draft.store_id)?;
                compose::inspect(
                    &self.namespace,
                    &crate::send::DraftRef { entry_id, store_id },
                )
            }
            MailCommand::ComposeSend { draft, preview } => {
                let (entry_id, store_id) = self.resolve_ids(&draft.entry_id, &draft.store_id)?;
                compose::send(
                    &self.namespace,
                    &crate::send::DraftRef { entry_id, store_id },
                    &preview,
                )
            }
            MailCommand::ComposeLocate(send_id) => compose::locate(&self.namespace, &send_id),
            MailCommand::MarkMail(mut params) => {
                (params.entry_id, params.store_id) =
                    self.resolve_ids(&params.entry_id, &params.store_id)?;
                modify::mark_mail(&self.namespace, params)
            }
            MailCommand::SetCategory(mut params) => {
                (params.entry_id, params.store_id) =
                    self.resolve_ids(&params.entry_id, &params.store_id)?;
                modify::set_category(&self.namespace, params)
            }
            MailCommand::CreateFolder(params) => modify::create_folder(&self.namespace, params),
            MailCommand::MoveMail(mut params) => {
                (params.entry_id, params.store_id) =
                    self.resolve_ids(&params.entry_id, &params.store_id)?;
                modify::move_mail(&self.namespace, params)
            }
            MailCommand::ListEvents(params) => schedule::list_events(&self.namespace, params),
            MailCommand::GetEvent(mut params) => {
                let mut shim = GetMailParams {
                    entry_id: params.entry_id.clone(),
                    store_id: params.store_id.clone(),
                    include_body: None,
                    include_html: None,
                    max_body_chars: None,
                    body_offset: None,
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
                    body_offset: None,
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

#[cfg(test)]
mod tests {
    use super::default_folder_kind;

    #[test]
    fn folder_aliases_map_to_default_folders() {
        assert_eq!(default_folder_kind(" Trash "), Some(3));
        assert_eq!(default_folder_kind("Spam"), Some(23));
        assert_eq!(default_folder_kind("junk email"), Some(23));
        assert_eq!(default_folder_kind("Calendar"), Some(9));
        assert_eq!(default_folder_kind(""), Some(6));
        assert_eq!(default_folder_kind("Inbox/Projects"), None);
    }
}
