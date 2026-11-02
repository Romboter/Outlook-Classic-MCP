//! Outlook side of the send pipeline (`crate::send`): build and stamp a
//! draft, re-read it, send it once, and find it again by its stamp.
//!
//! Rules that matter here:
//! - nothing is saved until every recipient has resolved;
//! - the item is never read after `Send()` starts;
//! - any failure from `Send()` onward is reported as an unknown outcome.
use super::mailbox::{default_folder, optional_text};
use super::win;
use crate::send::{Addr, DraftRef, PrepareRequest, Preview};
use crate::ticket::{NOT_SENT, UNKNOWN_OUTCOME};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::sync::atomic::{AtomicU64, Ordering};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const OL_MAIL_ITEM: i32 = 0;
const OL_FOLDER_OUTBOX: i32 = 4;
const OL_FOLDER_SENT: i32 = 5;
const OL_FOLDER_DRAFTS: i32 = 16;
const OL_TEXT: i32 = 1;
const STAMP: &str = "OutlookMcpSendId";
const SMTP_PROPTAG: &str = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
const EXCERPT_CHARS: usize = 600;
const MAX_RECIPIENTS: i32 = 500;

fn fnv1a(text: &str) -> String {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in text.bytes() {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
    format!("{hash:016x}")
}

fn new_send_id() -> String {
    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    format!(
        "{:x}-{:x}-{:x}",
        nanos,
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::SeqCst)
    )
}

fn not_sent(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

fn mail_by_id(namespace: &IDispatch, draft: &DraftRef) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(draft.entry_id.as_str())];
    if let Some(store) = &draft.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Message no longer available; EntryID may be stale")
}

fn kind_name(kind: i32) -> &'static str {
    match kind {
        2 => "cc",
        3 => "bcc",
        _ => "to",
    }
}

fn smtp_of(recipient: &IDispatch) -> Result<String> {
    let from_property = win::object_property(recipient, "PropertyAccessor")
        .and_then(|accessor| {
            win::invoke(
                &accessor,
                "GetProperty",
                DISPATCH_METHOD,
                vec![VARIANT::from(SMTP_PROPTAG)],
            )
        })
        .ok()
        .and_then(|v| windows::core::BSTR::try_from(&v).ok())
        .map(|b| b.to_string())
        .filter(|s| s.contains('@'));
    if let Some(address) = from_property {
        return Ok(address.to_ascii_lowercase());
    }
    let plain = optional_text(recipient, "Address");
    if plain.contains('@') && !plain.starts_with('/') {
        return Ok(plain.to_ascii_lowercase());
    }
    bail!("could not determine an SMTP address for a recipient")
}

fn recipients_of(item: &IDispatch) -> Result<Vec<Addr>> {
    let recipients = win::object_property(item, "Recipients")?;
    let count = win::int_property(&recipients, "Count")?.clamp(0, MAX_RECIPIENTS);
    let mut out = Vec::new();
    for i in 1..=count {
        let r = win::method_object(&recipients, "Item", vec![VARIANT::from(i)])?;
        out.push(Addr {
            kind: kind_name(win::int_property(&r, "Type").unwrap_or(1)).into(),
            address: smtp_of(&r)?,
        });
    }
    Ok(out)
}

fn sender_of(namespace: &IDispatch) -> String {
    let entry = win::object_property(namespace, "CurrentUser")
        .and_then(|user| win::object_property(&user, "AddressEntry"));
    let Ok(entry) = entry else {
        return String::new();
    };
    if let Ok(user) = win::method_object(&entry, "GetExchangeUser", vec![]) {
        let smtp = optional_text(&user, "PrimarySmtpAddress");
        if smtp.contains('@') {
            return smtp.to_ascii_lowercase();
        }
    }
    optional_text(&entry, "Address").to_ascii_lowercase()
}

fn canonical(recipients: &[Addr], subject: &str, body: &str) -> String {
    let mut text = String::new();
    for r in recipients {
        text.push_str(&format!("{}:{}\n", r.kind, r.address));
    }
    text.push_str(subject);
    text.push('\n');
    text.push_str(&body.replace("\r\n", "\n"));
    text
}

fn preview_of(
    namespace: &IDispatch,
    item: &IDispatch,
    send_id: &str,
    kind: &str,
) -> Result<Preview> {
    let recipients = recipients_of(item)?;
    let subject = optional_text(item, "Subject");
    let body = optional_text(item, "Body");
    Ok(Preview {
        send_id: send_id.to_string(),
        sender: sender_of(namespace),
        hash: fnv1a(&canonical(&recipients, &subject, &body)),
        recipients,
        subject,
        body_excerpt: body
            .replace("\r\n", "\n")
            .chars()
            .take(EXCERPT_CHARS)
            .collect(),
        body_chars: body.chars().count(),
        kind: kind.to_string(),
    })
}

fn stamp_of(item: &IDispatch) -> Option<String> {
    let props = win::object_property(item, "UserProperties").ok()?;
    let prop = win::method_object(&props, "Find", vec![VARIANT::from(STAMP)]).ok()?;
    win::string_property(&prop, "Value").ok()
}

fn in_drafts(namespace: &IDispatch, item: &IDispatch) -> bool {
    let (Ok(parent), Ok(drafts)) = (
        win::object_property(item, "Parent"),
        default_folder(namespace, OL_FOLDER_DRAFTS),
    ) else {
        return false;
    };
    let id = optional_text(&drafts, "EntryID");
    !id.is_empty() && optional_text(&parent, "EntryID") == id
}

fn set_recipients(item: &IDispatch, property: &str, list: &[String]) -> Result<()> {
    if !list.is_empty() {
        win::put_property(item, property, VARIANT::from(list.join("; ").as_str()))?;
    }
    Ok(())
}

pub(super) fn prepare(
    application: &IDispatch,
    namespace: &IDispatch,
    request: PrepareRequest,
    resolve_source: impl Fn(&DraftRef) -> Result<DraftRef>,
) -> Result<Value> {
    let (item, kind) = match &request {
        PrepareRequest::New {
            to,
            cc,
            bcc,
            subject,
            body,
            importance,
        } => {
            let item =
                win::method_object(application, "CreateItem", vec![VARIANT::from(OL_MAIL_ITEM)])?;
            win::put_property(&item, "BodyFormat", VARIANT::from(1i32))?;
            set_recipients(&item, "To", to)?;
            set_recipients(&item, "CC", cc)?;
            set_recipients(&item, "BCC", bcc)?;
            win::put_property(&item, "Subject", VARIANT::from(subject.as_str()))?;
            win::put_property(&item, "Body", VARIANT::from(body.as_str()))?;
            let level = match importance.as_str() {
                "low" => 0i32,
                "high" => 2,
                _ => 1,
            };
            win::put_property(&item, "Importance", VARIANT::from(level))?;
            (item, "new")
        }
        PrepareRequest::Reply {
            draft_of,
            body,
            reply_all,
        } => {
            let source = mail_by_id(namespace, &resolve_source(draft_of)?)?;
            let reply = win::method_object(
                &source,
                if *reply_all { "ReplyAll" } else { "Reply" },
                vec![],
            )?;
            let quoted = optional_text(&reply, "Body");
            win::put_property(
                &reply,
                "Body",
                VARIANT::from(format!("{body}\n\n{quoted}").as_str()),
            )?;
            (reply, if *reply_all { "reply_all" } else { "reply" })
        }
        PrepareRequest::Forward {
            draft_of,
            to,
            cc,
            body,
        } => {
            let source = mail_by_id(namespace, &resolve_source(draft_of)?)?;
            let forward = win::method_object(&source, "Forward", vec![])?;
            set_recipients(&forward, "To", to)?;
            set_recipients(&forward, "CC", cc)?;
            if !body.is_empty() {
                let quoted = optional_text(&forward, "Body");
                win::put_property(
                    &forward,
                    "Body",
                    VARIANT::from(format!("{body}\n\n{quoted}").as_str()),
                )?;
            }
            (forward, "forward")
        }
    };

    // Nothing is saved until every recipient is a real, resolved address.
    let recipients = win::object_property(&item, "Recipients")?;
    let resolved = win::invoke(&recipients, "ResolveAll", DISPATCH_METHOD, vec![])
        .ok()
        .and_then(|v| bool::try_from(&v).ok())
        .unwrap_or(false);
    let count = win::int_property(&recipients, "Count")?.clamp(0, MAX_RECIPIENTS);
    if count == 0 {
        bail!("the message has no recipients");
    }
    if !resolved {
        bail!("one or more recipients could not be resolved to a mailbox; nothing was saved");
    }

    let send_id = new_send_id();
    let props = win::object_property(&item, "UserProperties")?;
    let stamp = win::method_object(
        &props,
        "Add",
        vec![
            VARIANT::from(STAMP),
            VARIANT::from(OL_TEXT),
            VARIANT::from(false),
        ],
    )?;
    win::put_property(&stamp, "Value", VARIANT::from(send_id.as_str()))?;
    win::invoke(&item, "Save", DISPATCH_METHOD, vec![])?;

    // The preview comes from the SAVED item, so the hash matches what is re-read at send time.
    let preview = preview_of(namespace, &item, &send_id, kind)?;
    Ok(json!({
        "entry_id": optional_text(&item, "EntryID"),
        "store_id": "",
        "preview": preview,
    }))
}

pub(super) fn inspect(namespace: &IDispatch, draft: &DraftRef) -> Result<Value> {
    let item = mail_by_id(namespace, draft)?;
    if !in_drafts(namespace, &item) {
        bail!("that message is not in Drafts (it may already have been sent)");
    }
    let Some(send_id) = stamp_of(&item) else {
        bail!("that draft was not prepared by this server, so it will not be sent from here");
    };
    let preview = preview_of(namespace, &item, &send_id, "new")?;
    Ok(json!({
        "entry_id": optional_text(&item, "EntryID"),
        "store_id": "",
        "preview": preview,
    }))
}

pub(super) fn send(namespace: &IDispatch, draft: &DraftRef, preview: &Preview) -> Result<Value> {
    // Everything before Send() can fail safely: the draft stays in Drafts.
    let item = mail_by_id(namespace, draft).map_err(|e| not_sent(&format!("{e:#}")))?;
    if !in_drafts(namespace, &item) {
        return Err(not_sent("the message is not in Drafts"));
    }
    if stamp_of(&item).as_deref() != Some(preview.send_id.as_str()) {
        return Err(not_sent("the draft does not carry this send id"));
    }
    let current = preview_of(namespace, &item, &preview.send_id, &preview.kind)
        .map_err(|e| not_sent(&format!("{e:#}")))?;
    if current.hash != preview.hash {
        return Err(not_sent("the draft changed after it was previewed"));
    }

    // From here the outcome is unknown unless Send() returns cleanly, and the
    // item must not be touched again: it moves out of Drafts when sent.
    match win::invoke(&item, "Send", DISPATCH_METHOD, vec![]) {
        Ok(_) => Ok(json!({"sent": true})),
        Err(error) => Err(anyhow::anyhow!(
            "{UNKNOWN_OUTCOME}Send() reported an error ({error:#}); it may still have been sent"
        )),
    }
}

pub(super) fn locate(namespace: &IDispatch, send_id: &str) -> Result<Value> {
    if send_id.is_empty()
        || send_id.len() > 80
        || !send_id
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-')
    {
        bail!("invalid send_id");
    }
    let filter = format!(
        "@SQL=\"http://schemas.microsoft.com/mapi/string/{{00020329-0000-0000-C000-000000000046}}/{STAMP}\" = '{send_id}'"
    );
    for (folder, place) in [
        (OL_FOLDER_SENT, "sent"),
        (OL_FOLDER_OUTBOX, "queued"),
        (OL_FOLDER_DRAFTS, "drafts"),
    ] {
        let found = default_folder(namespace, folder)
            .and_then(|f| win::object_property(&f, "Items"))
            .and_then(|items| {
                win::method_object(&items, "Restrict", vec![VARIANT::from(filter.as_str())])
            })
            .and_then(|hits| win::int_property(&hits, "Count"))
            .unwrap_or(0);
        if found > 0 {
            return Ok(json!({"location": place}));
        }
    }
    Ok(json!({"location": "unknown"}))
}
