//! Outlook side of confirmed meeting cancellation (`crate::action`).
//!
//! Setting `MeetingStatus` to cancelled and calling `Send()` mails the
//! cancellation to every attendee; `Delete()` then moves the organizer's copy
//! to Deleted Items. Nothing here changes the item until the person has
//! approved, the item is stamped first so it can be found afterwards, and any
//! doubt once the status is changed is an unknown outcome.
use super::compose::{fnv1a, new_id};
use super::delete::{in_deleted_items, stamp_of};
use super::mailbox::{default_folder, optional_text};
use super::schedule::date_iso;
use super::win;
use crate::action::{ActionPreview, ActionTarget, Fact};
use crate::ticket::{NOT_SENT, UNKNOWN_OUTCOME};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const STAMP: &str = "OutlookMcpActionId";
const OL_TEXT: i32 = 1;
const OL_CLASS_APPOINTMENT: i32 = 26;
const OL_FOLDER_CALENDAR: i32 = 9;
const OL_FOLDER_DELETED: i32 = 3;
const SCAN_NEWEST: i32 = 100;
const MAX_LISTED: i32 = 50;
/// `MeetingStatus`: a meeting you organized and have not cancelled.
const OL_MEETING: i32 = 1;
/// `MeetingStatus`: a meeting you organized and cancelled.
const OL_MEETING_CANCELED: i32 = 5;

fn not_done(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

fn item_by_id(namespace: &IDispatch, target: &ActionTarget) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(target.entry_id.as_str())];
    if let Some(store) = &target.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Event no longer available; EntryID may be stale")
}

/// Pre-prompt refusals: only an open meeting you organized can be cancelled.
fn check_allowed(namespace: &IDispatch, item: &IDispatch) -> Result<()> {
    if win::int_property(item, "Class")? != OL_CLASS_APPOINTMENT {
        bail!("that EntryID does not identify a calendar event");
    }
    if in_deleted_items(namespace, item) {
        bail!("that event is in Deleted Items; it cannot be cancelled from there");
    }
    match win::int_property(item, "MeetingStatus").unwrap_or(0) {
        OL_MEETING => Ok(()),
        0 => bail!("that event has no attendees; use outlook_delete_event instead"),
        OL_MEETING_CANCELED => bail!("that meeting is already cancelled"),
        _ => bail!("you did not organize that meeting, so you cannot cancel it"),
    }
}

/// Readable address: Exchange recipients carry an internal `/o=...` address,
/// so ask the address entry for the SMTP one and show nothing rather than the DN.
pub(super) fn smtp_address(recipient: &IDispatch) -> String {
    let plain = optional_text(recipient, "Address");
    if !plain.starts_with('/') {
        return plain;
    }
    win::object_property(recipient, "AddressEntry")
        .and_then(|entry| win::method_object(&entry, "GetExchangeUser", vec![]))
        .map(|user| optional_text(&user, "PrimarySmtpAddress"))
        .unwrap_or_default()
}

/// `(count, "Name <address>" for the first MAX_LISTED)`.
pub(super) fn attendees(item: &IDispatch) -> (i32, Vec<String>) {
    let Ok(recipients) = win::object_property(item, "Recipients") else {
        return (0, vec![]);
    };
    let count = win::int_property(&recipients, "Count").unwrap_or(0).max(0);
    let listed = (1..=count.min(MAX_LISTED))
        .filter_map(|i| win::method_object(&recipients, "Item", vec![VARIANT::from(i)]).ok())
        .map(|r| format!("{} <{}>", optional_text(&r, "Name"), smtp_address(&r)))
        .collect();
    (count, listed)
}

fn identity_of(entry_id: &str, item: &IDispatch) -> String {
    let (count, listed) = attendees(item);
    format!(
        "{entry_id}|{}|{}|{}|{count}|{}",
        optional_text(item, "Subject"),
        date_iso(item, "Start").unwrap_or_default(),
        win::int_property(item, "RecurrenceState").unwrap_or(0),
        listed.join(";")
    )
}

pub(super) fn prepare(namespace: &IDispatch, target: &ActionTarget) -> Result<Value> {
    let item = item_by_id(namespace, target)?;
    check_allowed(namespace, &item)?;
    let (count, listed) = attendees(&item);
    if count == 0 {
        bail!("that meeting lists no attendees; use outlook_delete_event instead");
    }
    let fact = |label: &str, value: String| Fact {
        label: label.into(),
        value,
    };
    let mut attendee_text = listed.join("; ");
    if count > MAX_LISTED {
        attendee_text.push_str(&format!("; and {} more", count - MAX_LISTED));
    }
    let mut warnings = vec![format!(
        "Cancellation emails are sent to all {count} attendees and cannot be recalled."
    )];
    if win::int_property(&item, "RecurrenceState").unwrap_or(0) == 1 {
        warnings.push("This cancels the whole recurring series.".into());
    }
    let action_id = new_id();
    // The stamp survives the move to Deleted Items, so a later check can find it.
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
    win::put_property(&stamp, "Value", VARIANT::from(action_id.as_str()))?;
    win::invoke(&item, "Save", DISPATCH_METHOD, vec![])?;
    Ok(json!({"preview": ActionPreview {
        action_id,
        summary: "Cancel this meeting and email every attendee.".into(),
        facts: vec![
            fact("Meeting", optional_text(&item, "Subject")),
            fact("Starts", date_iso(&item, "Start").unwrap_or_default()),
            fact("Ends", date_iso(&item, "End").unwrap_or_default()),
            fact("Location", optional_text(&item, "Location")),
            fact("Attendees", attendee_text),
        ],
        warnings,
        hash: fnv1a(&identity_of(&target.entry_id, &item)),
    }}))
}

pub(super) fn execute(
    namespace: &IDispatch,
    target: &ActionTarget,
    preview: &ActionPreview,
) -> Result<Value> {
    // Everything before the status change can fail safely: nobody is told.
    let item = item_by_id(namespace, target).map_err(|e| not_done(&format!("{e:#}")))?;
    check_allowed(namespace, &item).map_err(|e| not_done(&format!("{e:#}")))?;
    if stamp_of(&item).as_deref() != Some(preview.action_id.as_str()) {
        return Err(not_done("the event does not carry this action id"));
    }
    if fnv1a(&identity_of(&target.entry_id, &item)) != preview.hash {
        return Err(not_done(
            "the meeting or its attendee list changed after it was shown to you",
        ));
    }

    // From here the cancellation may already have gone out.
    let unknown = |why: String| anyhow::anyhow!("{UNKNOWN_OUTCOME}{why}");
    win::put_property(
        &item,
        "MeetingStatus",
        VARIANT::from(OL_MEETING_CANCELED),
    )
    .map_err(|e| {
        unknown(format!(
            "marking the meeting cancelled reported an error ({e:#}); nobody should have been told, but check"
        ))
    })?;
    win::invoke(&item, "Send", DISPATCH_METHOD, vec![]).map_err(|e| {
        unknown(format!(
            "sending the cancellations reported an error ({e:#}); attendees may or may not have been told"
        ))
    })?;
    // The cancellations are out; moving the organizer's copy is tidying up and
    // must not turn a sent cancellation into a reported failure.
    let moved = win::invoke(&item, "Delete", DISPATCH_METHOD, vec![]).is_ok();
    Ok(json!({"done": true, "cancellations_sent": true, "moved_to_deleted_items": moved}))
}

/// Find the stamped item: cancelled, or moved to Deleted Items, means done.
pub(super) fn locate(
    namespace: &IDispatch,
    action_id: &str,
    target: Option<&ActionTarget>,
) -> Result<Value> {
    let verdict = |item: &IDispatch| {
        let done = win::int_property(item, "MeetingStatus").unwrap_or(0) == OL_MEETING_CANCELED
            || in_deleted_items(namespace, item);
        json!({"located": if done { "done" } else { "not_done" }})
    };
    if let Some(target) = target
        && let Ok(item) = item_by_id(namespace, target)
        && stamp_of(&item).as_deref() == Some(action_id)
    {
        return Ok(verdict(&item));
    }
    for folder in [OL_FOLDER_DELETED, OL_FOLDER_CALENDAR] {
        let Ok(items) =
            default_folder(namespace, folder).and_then(|f| win::object_property(&f, "Items"))
        else {
            continue;
        };
        let _ = win::invoke(
            &items,
            "Sort",
            DISPATCH_METHOD,
            vec![VARIANT::from("[LastModificationTime]"), VARIANT::from(true)],
        );
        let count = win::int_property(&items, "Count")
            .unwrap_or(0)
            .clamp(0, SCAN_NEWEST);
        for i in 1..=count {
            if let Ok(item) = win::method_object(&items, "Item", vec![VARIANT::from(i)])
                && stamp_of(&item).as_deref() == Some(action_id)
            {
                return Ok(verdict(&item));
            }
        }
    }
    Ok(json!({"located": "unknown"}))
}
