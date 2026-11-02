//! Outlook side of confirmed meeting responses (`crate::action`).
//!
//! `Respond` changes the calendar item at once (accepting or tentatively
//! accepting replaces it with a new item that has a different EntryID, and
//! declining removes it) and returns a response that is only emailed when
//! `Send()` is called. So nothing here calls `Respond` until the person has
//! approved, the item is stamped first so it can be found afterwards, and any
//! doubt once `Respond` begins is an unknown outcome.
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
/// `MeetingStatus`: an invitation received from someone else.
const OL_MEETING_RECEIVED: i32 = 3;

fn not_done(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

/// `(OlMeetingResponse for Respond, OlResponseStatus read back, wording)`.
fn response_codes(response: &str) -> Result<(i32, i32, &'static str)> {
    match response {
        "accept" => Ok((3, 3, "Accepted")),
        "tentative" => Ok((2, 2, "Tentative")),
        "decline" => Ok((4, 4, "Declined")),
        _ => bail!("response must be accept, tentative or decline"),
    }
}

fn status_word(status: i32) -> &'static str {
    match status {
        1 => "Organizer",
        2 => "Tentative",
        3 => "Accepted",
        4 => "Declined",
        5 => "Not yet answered",
        _ => "None",
    }
}

fn item_by_id(namespace: &IDispatch, target: &ActionTarget) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(target.entry_id.as_str())];
    if let Some(store) = &target.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Event no longer available; EntryID may be stale")
}

/// Pre-prompt refusals: only an invitation someone else sent can be answered.
fn check_allowed(namespace: &IDispatch, item: &IDispatch) -> Result<()> {
    if win::int_property(item, "Class")? != OL_CLASS_APPOINTMENT {
        bail!("that EntryID does not identify a calendar event");
    }
    if in_deleted_items(namespace, item) {
        bail!("that event is in Deleted Items; it cannot be answered from there");
    }
    match win::int_property(item, "MeetingStatus").unwrap_or(0) {
        OL_MEETING_RECEIVED => Ok(()),
        0 => bail!("that event is not a meeting invitation, so there is nothing to answer"),
        1 => bail!("you organized that meeting; an organizer cannot respond to it"),
        _ => bail!("that meeting was cancelled or is not an open invitation"),
    }
}

fn identity_of(entry_id: &str, item: &IDispatch) -> String {
    format!(
        "{entry_id}|{}|{}|{}|{}",
        optional_text(item, "Subject"),
        date_iso(item, "Start").unwrap_or_default(),
        optional_text(item, "Organizer"),
        win::int_property(item, "ResponseStatus").unwrap_or(0)
    )
}

pub(super) fn prepare(namespace: &IDispatch, target: &ActionTarget) -> Result<Value> {
    let (_, _, wording) = response_codes(target.response.as_deref().unwrap_or(""))?;
    let item = item_by_id(namespace, target)?;
    check_allowed(namespace, &item)?;
    let send = target.enabled.unwrap_or(true);
    let organizer = optional_text(&item, "Organizer");
    let current = win::int_property(&item, "ResponseStatus").unwrap_or(0);
    let fact = |label: &str, value: String| Fact {
        label: label.into(),
        value,
    };
    let mut warnings = Vec::new();
    if send {
        warnings.push(format!(
            "A reply will be emailed to the organizer ({organizer})."
        ));
    } else {
        warnings.push("The organizer will NOT be told; only your own calendar changes.".into());
    }
    if target.response.as_deref() == Some("decline") {
        warnings.push("Declining removes the meeting from your calendar.".into());
    } else {
        warnings.push("Your calendar item is replaced with a new copy.".into());
    }
    let action_id = new_id();
    // The stamp survives onto the replacement item, so a later check can find it.
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
        summary: format!("Answer this meeting invitation: {wording}."),
        facts: vec![
            fact("Meeting", optional_text(&item, "Subject")),
            fact("Starts", date_iso(&item, "Start").unwrap_or_default()),
            fact("Ends", date_iso(&item, "End").unwrap_or_default()),
            fact("Location", optional_text(&item, "Location")),
            fact("Organizer", organizer),
            fact("Your answer now", status_word(current).into()),
            fact("Your new answer", wording.into()),
            fact("Reply emailed to organizer", if send { "yes" } else { "no" }.into()),
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
    let (code, _, _) = response_codes(target.response.as_deref().unwrap_or(""))
        .map_err(|e| not_done(&format!("{e:#}")))?;
    let send = target.enabled.unwrap_or(true);
    // Everything before Respond() can fail safely: the invitation is untouched.
    let item = item_by_id(namespace, target).map_err(|e| not_done(&format!("{e:#}")))?;
    check_allowed(namespace, &item).map_err(|e| not_done(&format!("{e:#}")))?;
    if stamp_of(&item).as_deref() != Some(preview.action_id.as_str()) {
        return Err(not_done("the event does not carry this action id"));
    }
    if fnv1a(&identity_of(&target.entry_id, &item)) != preview.hash {
        return Err(not_done("the meeting changed after it was shown to you"));
    }

    // From here the calendar may already have changed. The response is created
    // by Respond() and emailed only by Send().
    let unknown = |why: String| anyhow::anyhow!("{UNKNOWN_OUTCOME}{why}");
    let response = win::method_object(
        &item,
        "Respond",
        vec![VARIANT::from(code), VARIANT::from(true)],
    );
    if !send {
        return match response {
            Ok(_) => Ok(json!({"done": true, "emailed": false})),
            Err(e) => Err(unknown(format!(
                "Respond() reported an error ({e:#}); your calendar may have changed"
            ))),
        };
    }
    let response = response.map_err(|e| {
        unknown(format!("Respond() reported an error ({e:#}); your calendar may have changed and no reply was sent"))
    })?;
    win::invoke(&response, "Send", DISPATCH_METHOD, vec![]).map_err(|e| {
        unknown(format!("your answer was recorded but sending the reply reported an error ({e:#}); the reply may or may not have gone"))
    })?;
    Ok(json!({"done": true, "emailed": true}))
}

/// Find the stamped item on the calendar or, after a decline, in Deleted Items.
pub(super) fn locate(
    namespace: &IDispatch,
    action_id: &str,
    target: Option<&ActionTarget>,
) -> Result<Value> {
    let Some(target) = target else {
        return Ok(json!({"located": "unknown"}));
    };
    let (_, wanted, _) = response_codes(target.response.as_deref().unwrap_or(""))?;
    for folder in [OL_FOLDER_CALENDAR, OL_FOLDER_DELETED] {
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
            let Ok(item) = win::method_object(&items, "Item", vec![VARIANT::from(i)]) else {
                continue;
            };
            if stamp_of(&item).as_deref() == Some(action_id) {
                let status = win::int_property(&item, "ResponseStatus").unwrap_or(0);
                return Ok(json!({"located": if status == wanted { "done" } else { "not_done" }}));
            }
        }
    }
    Ok(json!({"located": "unknown"}))
}
