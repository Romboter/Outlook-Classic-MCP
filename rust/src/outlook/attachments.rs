//! Outlook side of confirmed attachment saving (`crate::action`).
//!
//! The person approves a prompt listing the message, every file name and size
//! and the resolved folder. Files that could run code are refused one by one,
//! with the file named and the reason given; the rest are saved. Names come
//! from the sender, so each one is validated, and `create_unique` never
//! overwrites an existing file.
use super::compose::{fnv1a, new_id};
use super::mailbox::optional_text;
use super::schedule::date_iso;
use super::table::public_address;
use super::win;
use crate::action::{ActionPreview, ActionTarget, Fact};
use crate::safe_path::{
    create_unique, default_roots, plain_path, refusal_reason, resolve_output_dir,
};
use crate::ticket::NOT_SENT;
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::path::{Path, PathBuf};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const MAX_ATTACHMENTS: i32 = 100;
/// `OlAttachmentType`: a shortcut to a file, with nothing stored in the message.
const OL_BY_REFERENCE: i32 = 4;

fn not_done(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

struct Attachment {
    object: IDispatch,
    name: String,
    size: i64,
    /// Set when the file must not be saved: names the file and says why.
    refusal: Option<String>,
}

fn item_by_id(namespace: &IDispatch, target: &ActionTarget) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(target.entry_id.as_str())];
    if let Some(store) = &target.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    let item = win::method_object(namespace, "GetItemFromID", args)
        .context("Message no longer available; EntryID may be stale")?;
    if !matches!(win::int_property(&item, "Class")?, 43 | 53) {
        bail!("that EntryID does not identify a mail item");
    }
    Ok(item)
}

fn human(size: i64) -> String {
    match size {
        s if s >= 1 << 20 => format!("{:.1} MB", s as f64 / f64::from(1 << 20)),
        s if s >= 1 << 10 => format!("{:.1} KB", s as f64 / f64::from(1 << 10)),
        s => format!("{s} bytes"),
    }
}

/// The requested attachments (all, or the 1-indexed one) with their verdicts.
fn gather(item: &IDispatch, index: Option<i64>) -> Result<Vec<Attachment>> {
    let collection = win::object_property(item, "Attachments")?;
    let count = win::int_property(&collection, "Count")?.clamp(0, MAX_ATTACHMENTS);
    if count == 0 {
        bail!("that message has no attachments");
    }
    let wanted: Vec<i32> = match index {
        Some(i) if (1..=i64::from(count)).contains(&i) => vec![i as i32],
        Some(i) => {
            bail!("attachment_index {i} is out of range (the message has {count}, 1-indexed)")
        }
        None => (1..=count).collect(),
    };
    let mut found = Vec::new();
    for i in wanted {
        let object = win::method_object(&collection, "Item", vec![VARIANT::from(i)])?;
        let name = optional_text(&object, "FileName");
        let refusal = if win::int_property(&object, "Type").unwrap_or(0) == OL_BY_REFERENCE {
            Some(format!(
                "{name} was not saved: it is a link to a file, not stored in the message"
            ))
        } else {
            refusal_reason(&name)
        };
        found.push(Attachment {
            size: i64::from(win::int_property(&object, "Size").unwrap_or(0)),
            object,
            name,
            refusal,
        });
    }
    Ok(found)
}

fn output_dir(target: &ActionTarget) -> Result<PathBuf> {
    let payload = target
        .payload
        .as_ref()
        .context("the request has no folder")?;
    let requested = payload["output_dir"].as_str().unwrap_or_default();
    resolve_output_dir(Path::new(requested), &default_roots()).map_err(anyhow::Error::msg)
}

fn index_of(target: &ActionTarget) -> Option<i64> {
    target.payload.as_ref()?["index"].as_i64()
}

fn identity(target: &ActionTarget, files: &[Attachment], dir: &Path) -> String {
    let list: Vec<String> = files
        .iter()
        .map(|f| format!("{}:{}", f.name, f.size))
        .collect();
    fnv1a(&format!(
        "{}|{}|{}",
        target.entry_id,
        list.join(";"),
        plain_path(dir)
    ))
}

pub(super) fn prepare(namespace: &IDispatch, target: &ActionTarget) -> Result<Value> {
    let item = item_by_id(namespace, target)?;
    let dir = output_dir(target)?;
    let files = gather(&item, index_of(target))?;
    let refused: Vec<&str> = files.iter().filter_map(|f| f.refusal.as_deref()).collect();
    let allowed: Vec<&Attachment> = files.iter().filter(|f| f.refusal.is_none()).collect();
    if allowed.is_empty() {
        // Nothing could be saved: say which files were refused and why, without a prompt.
        bail!("{}", refused.join("; "));
    }
    let fact = |label: &str, value: String| Fact {
        label: label.into(),
        value,
    };
    let sender = optional_text(&item, "SenderName");
    let address = public_address(&optional_text(&item, "SenderEmailAddress"));
    let mut summary = format!(
        "Save {} attachment(s) to {}.",
        allowed.len(),
        plain_path(&dir)
    );
    let mut warnings =
        vec!["Existing files are never overwritten; a taken name gets ' (1)'.".to_string()];
    for reason in &refused {
        warnings.push(format!("NOT SAVED: {reason}."));
        summary.push_str(&format!(" NOT saved: {reason}."));
    }
    let listing = allowed
        .iter()
        .map(|f| format!("{} ({})", f.name, human(f.size)))
        .collect::<Vec<_>>()
        .join("; ");
    Ok(json!({"preview": ActionPreview {
        action_id: new_id(),
        summary,
        facts: vec![
            fact("From", format!("{sender} <{address}>")),
            fact("Subject", optional_text(&item, "Subject")),
            fact("Received", date_iso(&item, "ReceivedTime").unwrap_or_default()),
            fact("Save to", plain_path(&dir)),
            fact("Files", listing),
        ],
        warnings,
        hash: identity(target, &files, &dir),
    }}))
}

pub(super) fn execute(
    namespace: &IDispatch,
    target: &ActionTarget,
    preview: &ActionPreview,
) -> Result<Value> {
    let fail = |e: anyhow::Error| not_done(&format!("{e:#}"));
    let item = item_by_id(namespace, target).map_err(fail)?;
    let dir = output_dir(target).map_err(fail)?;
    let files = gather(&item, index_of(target)).map_err(fail)?;
    if identity(target, &files, &dir) != preview.hash {
        return Err(not_done(
            "the message, its attachments or the folder changed after they were shown to you",
        ));
    }
    let allowed: Vec<&Attachment> = files.iter().filter(|f| f.refusal.is_none()).collect();
    let mut saved = Vec::new();
    for file in &allowed {
        let progress = format!("{} of {} saved", saved.len(), allowed.len());
        let (path, handle) = create_unique(&dir, &file.name)
            .map_err(|e| not_done(&format!("could not create {} ({e}); {progress}", file.name)))?;
        // Outlook writes the file itself, so the handle must be closed first.
        drop(handle);
        if let Err(e) = win::invoke(
            &file.object,
            "SaveAsFile",
            DISPATCH_METHOD,
            vec![VARIANT::from(plain_path(&path).as_str())],
        ) {
            let _ = std::fs::remove_file(&path);
            return Err(not_done(&format!(
                "Outlook could not save {} ({e:#}); {progress}",
                file.name
            )));
        }
        saved.push(
            path.file_name()
                .map_or_else(String::new, |n| n.to_string_lossy().into_owned()),
        );
    }
    Ok(json!({"done": true, "saved": saved}))
}
