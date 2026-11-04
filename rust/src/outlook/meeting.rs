//! Outlook side of confirmed meeting invitations and updates (`crate::action`).
//!
//! Invitations: nothing is created while the person is being asked. Prepare
//! builds a throwaway item only to resolve the attendees and read the times
//! back, then discards it; execute builds the item again, checks it still
//! matches what was shown, stamps it and calls `Send()` once. A declined
//! request therefore leaves nothing on the calendar.
//!
//! Updates: prepare only reads (and stamps) the existing meeting; execute
//! applies the changes and `Send()` mails the update to every attendee.
//! Any doubt once `Send()` begins is an unknown outcome and is never replayed.
use super::cancel::{attendees, smtp_address};
use super::compose::{fnv1a, new_id};
use super::delete::{in_deleted_items, stamp_of};
use super::mailbox::{bool_property, default_folder, optional_text};
use super::schedule::{date_iso, date_variant, resolve_bound};
use super::win;
use crate::action::{ActionPreview, ActionTarget, Fact};
use crate::contracts::{
    RecurrenceParams, describe_recurrence, describe_series, series_gate, series_warnings,
};
use crate::ticket::{NOT_SENT, UNKNOWN_OUTCOME};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const STAMP: &str = "OutlookMcpActionId";
const OL_TEXT: i32 = 1;
const OL_APPOINTMENT_ITEM: i32 = 1;
const OL_CLASS_APPOINTMENT: i32 = 26;
const OL_FOLDER_CALENDAR: i32 = 9;
const SCAN_NEWEST: i32 = 100;
const MAX_LISTED: usize = 50;
/// `MeetingStatus`: a meeting you organized and have not cancelled.
const OL_MEETING: i32 = 1;
const OL_REQUIRED: i32 = 1;
const OL_DISCARD: i32 = 1;

fn not_done(why: &str) -> anyhow::Error {
    anyhow::anyhow!("{NOT_SENT}{why}")
}

fn unknown(why: String) -> anyhow::Error {
    anyhow::anyhow!("{UNKNOWN_OUTCOME}{why}")
}

fn payload_of(target: &ActionTarget) -> Result<&Value> {
    target
        .payload
        .as_ref()
        .context("the request has no meeting details")
}

fn text(payload: &Value, key: &str) -> Option<String> {
    payload.get(key)?.as_str().map(str::to_string)
}

fn fact(label: &str, value: String) -> Fact {
    Fact {
        label: label.into(),
        value,
    }
}

fn stamp(item: &IDispatch, action_id: &str) -> Result<()> {
    let props = win::object_property(item, "UserProperties")?;
    let prop = win::method_object(
        &props,
        "Add",
        vec![
            VARIANT::from(STAMP),
            VARIANT::from(OL_TEXT),
            VARIANT::from(false),
        ],
    )?;
    win::put_property(&prop, "Value", VARIANT::from(action_id))
}

/// The user's own mail domain, to mark attendees outside it. Unknown means no marking.
fn own_domain(namespace: &IDispatch) -> Option<String> {
    let user = win::object_property(namespace, "CurrentUser").ok()?;
    let smtp = smtp_address(&user);
    smtp.split_once('@')
        .map(|(_, domain)| domain.to_ascii_lowercase())
}

fn own_address(namespace: &IDispatch) -> Option<String> {
    let user = win::object_property(namespace, "CurrentUser").ok()?;
    let smtp = smtp_address(&user).to_ascii_lowercase();
    (!smtp.is_empty()).then_some(smtp)
}

fn describe_attendees(namespace: &IDispatch, addresses: &[String]) -> (String, usize) {
    describe_with_domain(own_domain(namespace), addresses)
}

fn describe_with_domain(own: Option<String>, addresses: &[String]) -> (String, usize) {
    let mut external = 0;
    let mut parts: Vec<String> = addresses
        .iter()
        .take(MAX_LISTED)
        .map(|a| {
            let outside = own.as_deref().is_some_and(|d| {
                a.rsplit_once('@')
                    .is_some_and(|(_, domain)| !domain.eq_ignore_ascii_case(d))
            });
            if outside {
                external += 1;
                format!("{a} (EXTERNAL)")
            } else {
                a.clone()
            }
        })
        .collect();
    if addresses.len() > MAX_LISTED {
        parts.push(format!("and {} more", addresses.len() - MAX_LISTED));
    }
    (parts.join("; "), external)
}

// ---- invitations -------------------------------------------------------

/// Build an unsaved meeting from the request. Returns it with the resolved
/// SMTP address of every attendee, sorted.
fn build_invite(application: &IDispatch, payload: &Value) -> Result<(IDispatch, Vec<String>)> {
    let item = win::method_object(
        application,
        "CreateItem",
        vec![VARIANT::from(OL_APPOINTMENT_ITEM)],
    )?;
    let start = resolve_bound(&text(payload, "start").unwrap_or_default())?;
    let end = resolve_bound(&text(payload, "end").unwrap_or_default())?;
    win::put_property(
        &item,
        "Subject",
        VARIANT::from(text(payload, "subject").unwrap_or_default().as_str()),
    )?;
    win::put_property(&item, "Start", date_variant(&start)?)?;
    win::put_property(&item, "End", date_variant(&end)?)?;
    if let Some(location) = text(payload, "location") {
        win::put_property(&item, "Location", VARIANT::from(location.as_str()))?;
    }
    if let Some(body) = text(payload, "body") {
        win::put_property(&item, "Body", VARIANT::from(body.as_str()))?;
    }
    let reminder = !payload["no_reminder"].as_bool().unwrap_or(false);
    win::put_property(&item, "ReminderSet", VARIANT::from(reminder))?;
    if reminder {
        win::put_property(
            &item,
            "ReminderMinutesBeforeStart",
            VARIANT::from(payload["reminder_minutes"].as_i64().unwrap_or(15) as i32),
        )?;
    }
    win::put_property(&item, "MeetingStatus", VARIANT::from(OL_MEETING))?;
    if let Some(spec) = recurrence_of(payload)? {
        super::schedule_write::apply_recurrence(&item, &spec, &start)?;
    }
    let recipients = win::object_property(&item, "Recipients")?;
    let wanted: Vec<&str> = payload["attendees"]
        .as_array()
        .map(|a| a.iter().filter_map(Value::as_str).collect())
        .unwrap_or_default();
    let mut resolved = Vec::new();
    for address in &wanted {
        let recipient = win::method_object(&recipients, "Add", vec![VARIANT::from(*address)])?;
        win::put_property(&recipient, "Type", VARIANT::from(OL_REQUIRED))?;
        let ok = win::invoke(&recipient, "Resolve", DISPATCH_METHOD, vec![])
            .ok()
            .and_then(|v| bool::try_from(&v).ok())
            .unwrap_or(false);
        if !ok {
            bail!("'{address}' could not be resolved to a mailbox; nothing was created");
        }
        let smtp = smtp_address(&recipient);
        if smtp.is_empty() {
            bail!("'{address}' resolved to an address that cannot be shown; nothing was created");
        }
        resolved.push(smtp.to_ascii_lowercase());
    }
    resolved.sort();
    resolved.dedup();
    Ok((item, resolved))
}

/// The recurrence requested in the payload, if any.
fn recurrence_of(payload: &Value) -> Result<Option<RecurrenceParams>> {
    match payload.get("recurrence") {
        None | Some(Value::Null) => Ok(None),
        Some(spec) => serde_json::from_value(spec.clone())
            .map(Some)
            .context("the recurrence in the request is invalid"),
    }
}

fn invite_hash(payload: &Value, resolved: &[String]) -> String {
    fnv1a(&format!("{payload}|{}", resolved.join(",")))
}

pub(super) fn prepare_invite(
    application: &IDispatch,
    namespace: &IDispatch,
    target: &ActionTarget,
) -> Result<Value> {
    let payload = payload_of(target)?;
    let (item, resolved) = build_invite(application, payload)?;
    let (attendee_text, external) = describe_attendees(namespace, &resolved);
    let mut warnings = vec![format!(
        "Invitations are emailed to {} attendee(s) and cannot be recalled.",
        resolved.len()
    )];
    if external > 0 {
        warnings.push(format!(
            "{external} attendee(s) are outside your organization."
        ));
    }
    let mut facts = vec![
        fact("Meeting", optional_text(&item, "Subject")),
        fact("Starts", date_iso(&item, "Start").unwrap_or_default()),
        fact("Ends", date_iso(&item, "End").unwrap_or_default()),
        fact("Location", optional_text(&item, "Location")),
    ];
    if let Some(spec) = recurrence_of(payload)? {
        let start = resolve_bound(&text(payload, "start").unwrap_or_default())?;
        facts.push(fact("Repeats", describe_recurrence(&spec, &start)));
        warnings.push("This is a recurring series: every occurrence is invited.".into());
    }
    facts.push(fact("Attendees", attendee_text));
    let preview = ActionPreview {
        action_id: new_id(),
        summary: "Create this meeting and email the invitations.".into(),
        facts,
        warnings,
        hash: invite_hash(payload, &resolved),
    };
    // Nothing was saved; drop the draft so no phantom meeting is left behind.
    let _ = win::invoke(
        &item,
        "Close",
        DISPATCH_METHOD,
        vec![VARIANT::from(OL_DISCARD)],
    );
    Ok(json!({"preview": preview}))
}

pub(super) fn execute_invite(
    application: &IDispatch,
    target: &ActionTarget,
    preview: &ActionPreview,
) -> Result<Value> {
    let payload = payload_of(target).map_err(|e| not_done(&format!("{e:#}")))?;
    let (item, resolved) =
        build_invite(application, payload).map_err(|e| not_done(&format!("{e:#}")))?;
    if invite_hash(payload, &resolved) != preview.hash {
        let _ = win::invoke(
            &item,
            "Close",
            DISPATCH_METHOD,
            vec![VARIANT::from(OL_DISCARD)],
        );
        return Err(not_done(
            "the meeting or its attendees resolved differently from what was shown to you",
        ));
    }
    stamp(&item, &preview.action_id).map_err(|e| not_done(&format!("{e:#}")))?;
    // From here the invitations may have gone out.
    win::invoke(&item, "Send", DISPATCH_METHOD, vec![]).map_err(|e| {
        unknown(format!(
            "sending the invitations reported an error ({e:#}); attendees may or may not have been told"
        ))
    })?;
    Ok(json!({"done": true, "invitations_sent": true}))
}

/// A sent meeting is saved on the calendar with its stamp. An absent stamp
/// proves nothing, so the answer is then unknown.
pub(super) fn locate_invite(namespace: &IDispatch, action_id: &str) -> Result<Value> {
    if let Ok(items) = default_folder(namespace, OL_FOLDER_CALENDAR)
        .and_then(|f| win::object_property(&f, "Items"))
    {
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
                return Ok(json!({"located": "done"}));
            }
        }
    }
    Ok(json!({"located": "unknown"}))
}

// ---- updates -----------------------------------------------------------

fn item_by_id(namespace: &IDispatch, target: &ActionTarget) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(target.entry_id.as_str())];
    if let Some(store) = &target.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Event no longer available; EntryID may be stale")
}

/// Pre-prompt refusals: only an open meeting you organized; a recurring one
/// only as a whole series and with `whole_series` acknowledged.
fn check_updatable(namespace: &IDispatch, item: &IDispatch, payload: &Value) -> Result<()> {
    if win::int_property(item, "Class")? != OL_CLASS_APPOINTMENT {
        bail!("that EntryID does not identify a calendar event");
    }
    if in_deleted_items(namespace, item) {
        bail!("that event is in Deleted Items");
    }
    match win::int_property(item, "MeetingStatus").unwrap_or(0) {
        OL_MEETING => {}
        0 => bail!("that event has no attendees; use outlook_update_event instead"),
        5 => bail!("that meeting is cancelled"),
        _ => bail!("you did not organize that meeting; only the organizer can change it"),
    }
    // An unreadable state is treated as an occurrence: refuse rather than guess.
    let state = win::int_property(item, "RecurrenceState").unwrap_or(2);
    let whole = payload["whole_series"].as_bool().unwrap_or(false);
    series_gate(state, whole).map_err(anyhow::Error::msg)
}

/// The series recurrence as plain words plus whether it has exceptions; empty
/// and false for a non-recurring item. Part of the approval hash.
fn series_of(item: &IDispatch) -> (String, bool) {
    if win::int_property(item, "RecurrenceState").unwrap_or(0) == 0 {
        return (String::new(), false);
    }
    let Ok(pattern) = win::method_object(item, "GetRecurrencePattern", vec![]) else {
        return ("an unreadable recurrence pattern".into(), false);
    };
    let end = if bool_property(&pattern, "NoEndDate") {
        None
    } else {
        date_iso(&pattern, "PatternEndDate")
    };
    let words = describe_series(
        win::int_property(&pattern, "RecurrenceType").unwrap_or(-1),
        win::int_property(&pattern, "Interval").unwrap_or(1),
        win::int_property(&pattern, "DaysOfWeekMask").unwrap_or(0),
        end.as_deref(),
        if bool_property(&pattern, "NoEndDate") {
            0
        } else {
            win::int_property(&pattern, "Occurrences").unwrap_or(0)
        },
    );
    let exceptions = win::object_property(&pattern, "Exceptions")
        .and_then(|e| win::int_property(&e, "Count"))
        .unwrap_or(0)
        > 0;
    (words, exceptions)
}

fn identity_of(entry_id: &str, item: &IDispatch, payload: &Value) -> String {
    let (count, listed) = attendees(item);
    fnv1a(&format!(
        "{entry_id}|{}|{}|{}|{}|{}|{count}|{}|{}|{payload}|{}",
        optional_text(item, "Subject"),
        date_iso(item, "Start").unwrap_or_default(),
        date_iso(item, "End").unwrap_or_default(),
        optional_text(item, "Location"),
        fnv1a(&optional_text(item, "Body")),
        listed.join(";"),
        current_attendees(item).join(","),
        series_of(item).0
    ))
}

const OL_ORGANIZER_TYPE: i32 = 0;
const MAX_ATTENDEES: usize = 100;

/// Lowercased SMTP address of every non-organizer recipient (no listing cap).
fn current_attendees(item: &IDispatch) -> Vec<String> {
    let Ok(recipients) = win::object_property(item, "Recipients") else {
        return vec![];
    };
    let count = win::int_property(&recipients, "Count").unwrap_or(0).max(0);
    (1..=count)
        .filter_map(|i| win::method_object(&recipients, "Item", vec![VARIANT::from(i)]).ok())
        .filter(|r| win::int_property(r, "Type").unwrap_or(OL_REQUIRED) != OL_ORGANIZER_TYPE)
        .map(|r| smtp_address(&r).to_ascii_lowercase())
        .filter(|a| !a.is_empty())
        .collect()
}

fn string_list(payload: &Value, key: &str) -> Vec<String> {
    payload[key]
        .as_array()
        .map(|a| {
            a.iter()
                .filter_map(Value::as_str)
                .map(String::from)
                .collect()
        })
        .unwrap_or_default()
}

/// The attendee list after the requested changes, or a plain refusal naming the address.
fn plan_attendees(
    current: &[String],
    own: Option<&str>,
    adds: &[String],
    removes: &[String],
) -> Result<Vec<String>> {
    for a in adds {
        if current.contains(a) {
            bail!("{a} is already an attendee; nothing was changed");
        }
    }
    for r in removes {
        if own == Some(r.as_str()) {
            bail!("{r} is the organizer and cannot be removed; nothing was changed");
        }
        if !current.contains(r) {
            bail!("{r} is not an attendee of this meeting; nothing was changed");
        }
    }
    let mut result: Vec<String> = current
        .iter()
        .filter(|a| !removes.contains(a))
        .cloned()
        .collect();
    result.extend(adds.iter().cloned());
    if result.is_empty() {
        bail!("that would leave the meeting with no attendees; use outlook_cancel_meeting instead");
    }
    if result.len() > MAX_ATTENDEES {
        bail!(
            "that would make {} attendees; the limit is {MAX_ATTENDEES}",
            result.len()
        );
    }
    Ok(result)
}

/// Prompt lines for the attendee diff, and how many named addresses are outside the organization.
fn diff_facts(own: Option<&str>, adds: &[String], removes: &[String]) -> (Vec<Fact>, usize) {
    let mut facts = Vec::new();
    let mut external = 0;
    if !adds.is_empty() {
        let (text, n) = describe_with_domain(own.map(String::from), adds);
        external += n;
        facts.push(fact("Adding", text));
    }
    if !removes.is_empty() {
        let (text, n) = describe_with_domain(own.map(String::from), removes);
        external += n;
        facts.push(fact(
            "Removing",
            format!("{text} (they will receive a cancellation)"),
        ));
    }
    (facts, external)
}

/// Done when every added address is listed and every removed one is not.
fn attendees_applied(current: &[String], adds: &[String], removes: &[String]) -> bool {
    adds.iter().all(|a| current.contains(a)) && removes.iter().all(|r| !current.contains(r))
}

/// Minute-precision local time of a request value, for comparing with Outlook's.
fn minutes_of(value: &str) -> Result<String> {
    Ok(resolve_bound(value)?.iso().chars().take(16).collect())
}

fn current_minutes(item: &IDispatch, name: &str) -> String {
    date_iso(item, name)
        .unwrap_or_default()
        .chars()
        .take(16)
        .collect()
}

pub(super) fn prepare_update(namespace: &IDispatch, target: &ActionTarget) -> Result<Value> {
    let payload = payload_of(target)?;
    let item = item_by_id(namespace, target)?;
    check_updatable(namespace, &item, payload)?;
    let new_start = text(payload, "start").map(|s| minutes_of(&s)).transpose()?;
    let new_end = text(payload, "end").map(|s| minutes_of(&s)).transpose()?;
    let start = new_start
        .clone()
        .unwrap_or_else(|| current_minutes(&item, "Start"));
    let end = new_end
        .clone()
        .unwrap_or_else(|| current_minutes(&item, "End"));
    if end <= start {
        bail!("end must be after start; nothing was changed");
    }
    let mut facts = vec![fact("Meeting", optional_text(&item, "Subject"))];
    let mut change = |label: &str, old: String, new: Option<String>| {
        if let Some(new) = new {
            facts.push(fact(label, format!("{old}  ->  {new}")));
        }
    };
    change(
        "Subject",
        optional_text(&item, "Subject"),
        text(payload, "subject"),
    );
    change("Starts", current_minutes(&item, "Start"), new_start);
    change("Ends", current_minutes(&item, "End"), new_end);
    change(
        "Location",
        optional_text(&item, "Location"),
        text(payload, "location"),
    );
    if text(payload, "body").is_some() {
        facts.push(fact("Description", "will be replaced".into()));
    }
    let current = current_attendees(&item);
    if current.is_empty() {
        bail!("that meeting lists no attendees; use outlook_update_event instead");
    }
    let (adds, removes) = (
        string_list(payload, "add_attendees"),
        string_list(payload, "remove_attendees"),
    );
    let own = own_address(namespace);
    let addresses = plan_attendees(&current, own.as_deref(), &adds, &removes)?;
    let count = addresses.len();
    let (diff, _) = diff_facts(own_domain(namespace).as_deref(), &adds, &removes);
    facts.extend(diff);
    let (attendee_text, external) = describe_attendees(namespace, &addresses);
    facts.push(fact("Attendees told", attendee_text));
    let mut warnings = vec![format!(
        "An update is emailed to all {count} attendees and cannot be recalled."
    )];
    let (series, has_exceptions) = series_of(&item);
    if !series.is_empty() {
        facts.push(fact("Repeats", series));
        warnings.extend(series_warnings(has_exceptions));
    }
    if external > 0 {
        warnings.push(format!(
            "{external} attendee(s) are outside your organization."
        ));
    }
    let action_id = new_id();
    stamp(&item, &action_id)?;
    win::invoke(&item, "Save", DISPATCH_METHOD, vec![])?;
    Ok(json!({"preview": ActionPreview {
        action_id,
        summary: "Change this meeting and email the update to every attendee.".into(),
        facts,
        warnings,
        hash: identity_of(&target.entry_id, &item, payload),
    }}))
}

pub(super) fn execute_update(
    namespace: &IDispatch,
    target: &ActionTarget,
    preview: &ActionPreview,
) -> Result<Value> {
    let payload = payload_of(target).map_err(|e| not_done(&format!("{e:#}")))?;
    let item = item_by_id(namespace, target).map_err(|e| not_done(&format!("{e:#}")))?;
    check_updatable(namespace, &item, payload).map_err(|e| not_done(&format!("{e:#}")))?;
    if stamp_of(&item).as_deref() != Some(preview.action_id.as_str()) {
        return Err(not_done("the event does not carry this action id"));
    }
    if identity_of(&target.entry_id, &item, payload) != preview.hash {
        return Err(not_done(
            "the meeting or its attendee list changed after it was shown to you",
        ));
    }
    // Changes are held in memory until Send() saves and mails them.
    let apply = || -> Result<()> {
        if let Some(subject) = text(payload, "subject") {
            win::put_property(&item, "Subject", VARIANT::from(subject.as_str()))?;
        }
        if let Some(start) = text(payload, "start") {
            win::put_property(&item, "Start", date_variant(&resolve_bound(&start)?)?)?;
        }
        if let Some(end) = text(payload, "end") {
            win::put_property(&item, "End", date_variant(&resolve_bound(&end)?)?)?;
        }
        if let Some(location) = text(payload, "location") {
            win::put_property(&item, "Location", VARIANT::from(location.as_str()))?;
        }
        if let Some(body) = text(payload, "body") {
            win::put_property(&item, "Body", VARIANT::from(body.as_str()))?;
        }
        change_attendees(
            &item,
            &string_list(payload, "add_attendees"),
            &string_list(payload, "remove_attendees"),
        )?;
        Ok(())
    };
    apply().map_err(|e| {
        not_done(&format!(
            "could not apply the changes ({e:#}); nobody was told, but check the meeting in Outlook"
        ))
    })?;
    if current_minutes(&item, "End") <= current_minutes(&item, "Start") {
        return Err(not_done(
            "end would not be after start; nobody was told, but check the meeting in Outlook",
        ));
    }
    // From here the update may have gone out.
    win::invoke(&item, "Send", DISPATCH_METHOD, vec![]).map_err(|e| {
        unknown(format!(
            "sending the update reported an error ({e:#}); attendees may or may not have been told"
        ))
    })?;
    Ok(json!({"done": true, "update_sent": true}))
}

/// Remove then add recipients on the in-memory item (nothing is mailed until Send).
fn change_attendees(item: &IDispatch, adds: &[String], removes: &[String]) -> Result<()> {
    if adds.is_empty() && removes.is_empty() {
        return Ok(());
    }
    let recipients = win::object_property(item, "Recipients")?;
    let mut pending: Vec<&String> = removes.iter().collect();
    let count = win::int_property(&recipients, "Count").unwrap_or(0).max(0);
    for i in (1..=count).rev() {
        let recipient = win::method_object(&recipients, "Item", vec![VARIANT::from(i)])?;
        let smtp = smtp_address(&recipient).to_ascii_lowercase();
        if let Some(pos) = pending.iter().position(|r| **r == smtp) {
            win::invoke(&recipient, "Delete", DISPATCH_METHOD, vec![])?;
            pending.remove(pos);
        }
    }
    if let Some(missing) = pending.first() {
        bail!("{missing} was not found among the attendees");
    }
    for address in adds {
        let recipient =
            win::method_object(&recipients, "Add", vec![VARIANT::from(address.as_str())])?;
        win::put_property(&recipient, "Type", VARIANT::from(OL_REQUIRED))?;
        let ok = win::invoke(&recipient, "Resolve", DISPATCH_METHOD, vec![])
            .ok()
            .and_then(|v| bool::try_from(&v).ok())
            .unwrap_or(false);
        if !ok {
            bail!("'{address}' could not be resolved to a mailbox");
        }
        if smtp_address(&recipient).is_empty() {
            bail!("'{address}' resolved to an address that cannot be shown");
        }
    }
    let total = current_attendees(item).len();
    if total == 0 || total > MAX_ATTENDEES {
        bail!("the attendee list would have {total} entries");
    }
    Ok(())
}

/// Done when every requested field now has its requested value.
/// Attendee changes are checked against the saved list; unreadable means unknown.
pub(super) fn locate_update(namespace: &IDispatch, target: Option<&ActionTarget>) -> Result<Value> {
    let Some(target) = target else {
        return Ok(json!({"located": "unknown"}));
    };
    let payload = payload_of(target)?;
    let Ok(item) = item_by_id(namespace, target) else {
        return Ok(json!({"located": "unknown"}));
    };
    let mut matches = true;
    if let Some(v) = text(payload, "subject") {
        matches &= optional_text(&item, "Subject") == v;
    }
    if let Some(v) = text(payload, "location") {
        matches &= optional_text(&item, "Location") == v;
    }
    if let Some(v) = text(payload, "start") {
        matches &= minutes_of(&v)? == current_minutes(&item, "Start");
    }
    if let Some(v) = text(payload, "end") {
        matches &= minutes_of(&v)? == current_minutes(&item, "End");
    }
    let (adds, removes) = (
        string_list(payload, "add_attendees"),
        string_list(payload, "remove_attendees"),
    );
    if !adds.is_empty() || !removes.is_empty() {
        let current = current_attendees(&item);
        if current.is_empty() {
            return Ok(json!({"located": "unknown"}));
        }
        matches &= attendees_applied(&current, &adds, &removes);
    }
    Ok(json!({"located": if matches { "done" } else { "not_done" }}))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn v(items: &[&str]) -> Vec<String> {
        items.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn attendee_plan_refuses_bad_changes_by_name() {
        let cur = v(&["a@x.example", "b@x.example"]);
        let me = Some("me@x.example");
        let err = |adds: &[&str], removes: &[&str]| {
            plan_attendees(&cur, me, &v(adds), &v(removes))
                .unwrap_err()
                .to_string()
        };
        assert!(err(&["a@x.example"], &[]).contains("a@x.example is already an attendee"));
        assert!(err(&[], &["z@x.example"]).contains("z@x.example is not an attendee"));
        assert!(err(&[], &["a@x.example", "b@x.example"]).contains("outlook_cancel_meeting"));
        let mut with_me = cur.clone();
        with_me.push("me@x.example".into());
        let e = plan_attendees(&with_me, me, &[], &v(&["me@x.example"]))
            .unwrap_err()
            .to_string();
        assert!(e.contains("me@x.example is the organizer"));
        let many: Vec<String> = (0..100).map(|i| format!("u{i}@x.example")).collect();
        assert!(
            plan_attendees(&many, me, &v(&["n@x.example"]), &[])
                .unwrap_err()
                .to_string()
                .contains("101")
        );
        let ok = plan_attendees(&cur, me, &v(&["c@x.example"]), &v(&["a@x.example"])).unwrap();
        assert_eq!(ok, v(&["b@x.example", "c@x.example"]));
        assert!(plan_attendees(&many, me, &v(&["n@x.example"]), &v(&["u0@x.example"])).is_ok());
    }

    #[test]
    fn prompt_shows_the_diff_and_marks_outsiders() {
        let (facts, external) = diff_facts(
            Some("x.example"),
            &v(&["new@x.example", "out@y.example"]),
            &v(&["old@y.example"]),
        );
        assert_eq!(facts[0].label, "Adding");
        assert!(facts[0].value.contains("out@y.example (EXTERNAL)"));
        assert!(!facts[0].value.contains("new@x.example (EXTERNAL)"));
        assert_eq!(facts[1].label, "Removing");
        assert!(facts[1].value.contains("old@y.example (EXTERNAL)"));
        assert!(facts[1].value.contains("they will receive a cancellation"));
        assert_eq!(external, 2);
        assert!(diff_facts(None, &[], &[]).0.is_empty());
    }

    #[test]
    fn readback_needs_adds_present_and_removes_absent() {
        let cur = v(&["a@x.example", "c@x.example"]);
        assert!(attendees_applied(
            &cur,
            &v(&["c@x.example"]),
            &v(&["b@x.example"])
        ));
        assert!(!attendees_applied(&cur, &v(&["d@x.example"]), &[]));
        assert!(!attendees_applied(&cur, &[], &v(&["a@x.example"])));
    }
}
