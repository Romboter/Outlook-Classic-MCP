//! In-place mailbox changes (`Capability::Modify`): flag/read state,
//! categories, folder creation and moving mail.
//!
//! Each operation reads the result back from Outlook and reports what is
//! actually there. Nothing here sends or deletes mail; moving into Deleted
//! Items is refused because that is a delete and needs its own permission.
use super::mailbox::{bool_property, default_folder, folder_path, optional_text};
use super::win;
use crate::contracts::{CreateFolderParams, MarkMailParams, MoveMailParams, SetCategoryParams};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const OL_FOLDER_DELETED: i32 = 3;
const OL_FLAG_MARKED: i32 = 2;
const OL_FLAG_NONE: i32 = 0;

fn item_by_id(namespace: &IDispatch, entry_id: &str, store_id: Option<&str>) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(entry_id)];
    if let Some(store) = store_id {
        args.push(VARIANT::from(store));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Item no longer available; EntryID may be stale")
}

fn mail_item(namespace: &IDispatch, entry_id: &str, store_id: Option<&str>) -> Result<IDispatch> {
    let item = item_by_id(namespace, entry_id, store_id)?;
    let class = win::int_property(&item, "Class")?;
    if class != 43 && class != 53 {
        bail!("EntryID does not identify a mail item");
    }
    Ok(item)
}

fn save(item: &IDispatch) -> Result<()> {
    win::invoke(item, "Save", DISPATCH_METHOD, vec![]).map(|_| ())
}

pub(super) fn mark_mail(namespace: &IDispatch, params: MarkMailParams) -> Result<Value> {
    let item = mail_item(namespace, &params.entry_id, params.store_id.as_deref())?;
    if let Some(read) = params.read {
        win::put_property(&item, "UnRead", VARIANT::from(!read))?;
    }
    if let Some(flagged) = params.flagged {
        let status = if flagged {
            OL_FLAG_MARKED
        } else {
            OL_FLAG_NONE
        };
        win::put_property(&item, "FlagStatus", VARIANT::from(status))?;
    }
    save(&item)?;
    // Report what Outlook holds now, not what was asked for.
    let unread = bool_property(&item, "UnRead");
    let flagged = win::int_property(&item, "FlagStatus").ok() == Some(OL_FLAG_MARKED);
    let matches = params.read.is_none_or(|read| read != unread)
        && params.flagged.is_none_or(|want| want == flagged);
    Ok(json!({
        "status": if matches { "updated" } else { "unverified" },
        "entry_id": optional_text(&item, "EntryID"),
        "store_id": params.store_id.clone().unwrap_or_default(),
        "unread": unread,
        "flagged": flagged,
    }))
}

pub(super) fn set_category(namespace: &IDispatch, params: SetCategoryParams) -> Result<Value> {
    let item = item_by_id(namespace, &params.entry_id, params.store_id.as_deref())?;
    win::put_property(
        &item,
        "Categories",
        VARIANT::from(params.categories.as_str()),
    )
    .context("this item does not support categories")?;
    save(&item)?;
    let actual = optional_text(&item, "Categories");
    Ok(json!({
        "status": if actual == params.categories { "updated" } else { "unverified" },
        "entry_id": optional_text(&item, "EntryID"),
        "store_id": params.store_id.clone().unwrap_or_default(),
        "subject": optional_text(&item, "Subject"),
        "categories": actual,
    }))
}

pub(super) fn create_folder(namespace: &IDispatch, params: CreateFolderParams) -> Result<Value> {
    let parent = folder_path(namespace, params.parent.as_deref())?;
    let folders = win::object_property(&parent, "Folders")?;
    let created = win::method_object(&folders, "Add", vec![VARIANT::from(params.name.as_str())])
        .context("Outlook refused to create the folder (a folder with this name may exist)")?;
    Ok(json!({
        "status": "created",
        "name": optional_text(&created, "Name"),
        "path": format!("{}/{}", optional_text(&parent, "Name"), optional_text(&created, "Name")),
    }))
}

/// True when `folder` is the Deleted Items folder or sits inside it.
fn in_deleted_items(namespace: &IDispatch, folder: &IDispatch) -> bool {
    let Ok(deleted) = default_folder(namespace, OL_FOLDER_DELETED) else {
        return true; // cannot prove it is safe
    };
    let deleted_id = optional_text(&deleted, "EntryID");
    let mut current = folder.clone();
    for _ in 0..16 {
        if optional_text(&current, "EntryID") == deleted_id {
            return true;
        }
        match win::object_property(&current, "Parent") {
            Ok(parent) => current = parent,
            Err(_) => return false,
        }
    }
    true // implausibly deep: treat as unsafe
}

pub(super) fn move_mail(namespace: &IDispatch, params: MoveMailParams) -> Result<Value> {
    let item = mail_item(namespace, &params.entry_id, params.store_id.as_deref())?;
    let target = folder_path(namespace, Some(&params.target_folder))?;
    if in_deleted_items(namespace, &target) {
        bail!(
            "moving mail into Deleted Items is a delete; it is refused here and needs delete permission"
        );
    }
    let moved = win::method_object(&item, "Move", vec![VARIANT::from(target.clone())])
        .context("Outlook could not move the item")?;
    let landed = win::object_property(&moved, "Parent")
        .map(|p| optional_text(&p, "EntryID") == optional_text(&target, "EntryID"))
        .unwrap_or(false);
    Ok(json!({
        "status": if landed { "moved" } else { "unverified" },
        // The EntryID changes when an item moves; this is the new one.
        "entry_id": optional_text(&moved, "EntryID"),
        "store_id": optional_text(&target, "StoreID"),
        "folder": optional_text(&target, "Name"),
    }))
}
