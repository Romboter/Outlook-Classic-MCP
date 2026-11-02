//! Task and calendar writes that notify nobody and delete nothing:
//! create/complete a task, create/update an appointment on the user's own
//! calendar. Inviting attendees, responding to invitations and deleting
//! events need the confirmation workflow and are not here.
use super::mailbox::{bool_property, default_folder, optional_text};
use super::schedule::{date_iso, date_variant, resolve_bound};
use super::win;
use crate::contracts::{
    CompleteTaskParams, CreateEventParams, CreateTaskParams, RecurrenceParams, UpdateEventParams,
};
use crate::localtime::LocalTime;
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use windows::Win32::System::{
    Com::{DISPATCH_METHOD, IDispatch},
    Variant::VARIANT,
};

const OL_APPOINTMENT_ITEM: i32 = 1;
const OL_TASK_ITEM: i32 = 3;
const OL_FOLDER_CALENDAR: i32 = 9;
const OL_FOLDER_TASKS: i32 = 13;
const OL_CLASS_APPOINTMENT: i32 = 26;
const OL_CLASS_TASK: i32 = 48;
const OL_NON_MEETING: i32 = 0;

fn item_by_id(namespace: &IDispatch, entry_id: &str, store_id: Option<&str>) -> Result<IDispatch> {
    let mut args = vec![VARIANT::from(entry_id)];
    if let Some(store) = store_id {
        args.push(VARIANT::from(store));
    }
    win::method_object(namespace, "GetItemFromID", args)
        .context("Item no longer available; EntryID may be stale")
}

fn save(item: &IDispatch) -> Result<()> {
    win::invoke(item, "Save", DISPATCH_METHOD, vec![]).map(|_| ())
}

fn put_text(item: &IDispatch, property: &str, value: &str) -> Result<()> {
    win::put_property(item, property, VARIANT::from(value))
}

fn put_date(item: &IDispatch, property: &str, text: &str) -> Result<()> {
    win::put_property(item, property, date_variant(&resolve_bound(text)?)?)
}

fn importance(level: &str) -> i32 {
    match level.to_ascii_lowercase().as_str() {
        "low" => 0,
        "high" => 2,
        _ => 1,
    }
}

pub(super) fn create_task(
    application: &IDispatch,
    namespace: &IDispatch,
    params: CreateTaskParams,
) -> Result<Value> {
    let task = win::method_object(application, "CreateItem", vec![VARIANT::from(OL_TASK_ITEM)])?;
    put_text(&task, "Subject", &params.subject)?;
    if let Some(due) = &params.due_date {
        put_date(&task, "DueDate", due)?;
    }
    if let Some(body) = &params.body {
        put_text(&task, "Body", body)?;
    }
    win::put_property(
        &task,
        "Importance",
        VARIANT::from(importance(params.importance.as_deref().unwrap_or("normal"))),
    )?;
    if let Some(reminder) = &params.reminder {
        win::put_property(&task, "ReminderSet", VARIANT::from(true))?;
        put_date(&task, "ReminderTime", reminder)?;
    }
    save(&task)?;
    let store = default_folder(namespace, OL_FOLDER_TASKS)
        .map(|f| optional_text(&f, "StoreID"))
        .unwrap_or_default();
    Ok(json!({
        "status": "created",
        "entry_id": optional_text(&task, "EntryID"),
        "store_id": store,
        "subject": optional_text(&task, "Subject"),
        "due_date": date_iso(&task, "DueDate"),
    }))
}

pub(super) fn complete_task(namespace: &IDispatch, params: CompleteTaskParams) -> Result<Value> {
    let task = item_by_id(namespace, &params.entry_id, params.store_id.as_deref())?;
    if win::int_property(&task, "Class").ok() != Some(OL_CLASS_TASK) {
        bail!("EntryID does not identify a task");
    }
    win::invoke(&task, "MarkComplete", DISPATCH_METHOD, vec![])?;
    save(&task)?;
    let complete = bool_property(&task, "Complete");
    Ok(json!({
        "status": if complete { "completed" } else { "unverified" },
        "entry_id": optional_text(&task, "EntryID"),
        "store_id": params.store_id.clone().unwrap_or_default(),
        "subject": optional_text(&task, "Subject"),
        "complete": complete,
    }))
}

fn apply_recurrence(event: &IDispatch, spec: &RecurrenceParams) -> Result<()> {
    let pattern = win::method_object(event, "GetRecurrencePattern", vec![])?;
    let kind = match spec.kind.as_str() {
        "daily" => 0,
        "weekly" => 1,
        "monthly" => 2,
        _ => 5, // yearly
    };
    win::put_property(&pattern, "RecurrenceType", VARIANT::from(kind))?;
    win::put_property(
        &pattern,
        "Interval",
        VARIANT::from(spec.interval.unwrap_or(1) as i32),
    )?;
    if let Some(count) = spec.occurrences {
        win::put_property(&pattern, "Occurrences", VARIANT::from(count as i32))?;
    }
    if let Some(end) = &spec.end_date {
        put_date(&pattern, "PatternEndDate", end)?;
    }
    Ok(())
}

pub(super) fn create_event(
    application: &IDispatch,
    namespace: &IDispatch,
    params: CreateEventParams,
) -> Result<Value> {
    let start = resolve_bound(&params.start)?;
    let end = resolve_bound(&params.end)?;
    if end.minutes() <= start.minutes() {
        bail!("end must be after start");
    }
    let event = win::method_object(
        application,
        "CreateItem",
        vec![VARIANT::from(OL_APPOINTMENT_ITEM)],
    )?;
    put_text(&event, "Subject", &params.subject)?;
    win::put_property(&event, "Start", date_variant(&start)?)?;
    win::put_property(&event, "End", date_variant(&end)?)?;
    if let Some(location) = &params.location {
        put_text(&event, "Location", location)?;
    }
    if let Some(body) = &params.body {
        put_text(&event, "Body", body)?;
    }
    win::put_property(&event, "ReminderSet", VARIANT::from(true))?;
    win::put_property(
        &event,
        "ReminderMinutesBeforeStart",
        VARIANT::from(params.reminder_minutes.unwrap_or(15) as i32),
    )?;
    if let Some(spec) = &params.recurrence {
        apply_recurrence(&event, spec)?;
    }
    save(&event)?;
    let store = default_folder(namespace, OL_FOLDER_CALENDAR)
        .map(|f| optional_text(&f, "StoreID"))
        .unwrap_or_default();
    Ok(json!({
        "status": "created",
        "entry_id": optional_text(&event, "EntryID"),
        "store_id": store,
        "subject": optional_text(&event, "Subject"),
        "start": date_iso(&event, "Start"),
        "end": date_iso(&event, "End"),
        "is_recurring": win::int_property(&event, "RecurrenceState").unwrap_or(0) != 0,
    }))
}

pub(super) fn update_event(namespace: &IDispatch, params: UpdateEventParams) -> Result<Value> {
    let event = item_by_id(namespace, &params.entry_id, params.store_id.as_deref())?;
    if win::int_property(&event, "Class").ok() != Some(OL_CLASS_APPOINTMENT) {
        bail!("EntryID does not identify a calendar event");
    }
    // Saving a change to a meeting would notify its attendees.
    if win::int_property(&event, "MeetingStatus").unwrap_or(OL_NON_MEETING) != OL_NON_MEETING {
        bail!(
            "this event is a meeting with attendees; changing it would notify them, which is not supported yet"
        );
    }
    let start = params.start.as_deref().map(resolve_bound).transpose()?;
    let end = params.end.as_deref().map(resolve_bound).transpose()?;
    if let (Some(start), Some(end)) = (&start, &end)
        && end.minutes() <= start.minutes()
    {
        bail!("end must be after start");
    }
    if let Some(subject) = &params.subject {
        put_text(&event, "Subject", subject)?;
    }
    if let Some(start) = &start {
        win::put_property(&event, "Start", date_variant(start)?)?;
    }
    if let Some(end) = &end {
        win::put_property(&event, "End", date_variant(end)?)?;
    }
    if let Some(location) = &params.location {
        put_text(&event, "Location", location)?;
    }
    if let Some(body) = &params.body {
        put_text(&event, "Body", body)?;
    }
    // Changing only one bound can invert the pair; check the result before saving.
    let (now_start, now_end) = (date_iso(&event, "Start"), date_iso(&event, "End"));
    if let (Some(s), Some(e)) = (&now_start, &now_end)
        && let (Ok(s), Ok(e)) = (LocalTime::parse(&s[..19]), LocalTime::parse(&e[..19]))
        && e.minutes() <= s.minutes()
    {
        bail!("end must be after start; nothing was saved");
    }
    save(&event)?;
    Ok(json!({
        "status": "updated",
        "entry_id": optional_text(&event, "EntryID"),
        "store_id": params.store_id.clone().unwrap_or_default(),
        "subject": optional_text(&event, "Subject"),
        "start": date_iso(&event, "Start"),
        "end": date_iso(&event, "End"),
    }))
}
