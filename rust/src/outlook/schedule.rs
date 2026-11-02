//! Read-only calendar and task reads.
//!
//! Dates are Outlook local wall-clock values rendered as ISO-8601 without an
//! offset. Every call is bounded by a result limit, a scan cap and a deadline.
use super::mailbox::{bool_property, default_folder, optional_text};
use super::win;
use crate::contracts::{
    GetEventParams, ListEventsParams, ListTasksParams, MAX_EVENT_RANGE_DAYS, checked_limit_in,
};
use crate::localtime::LocalTime;
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::time::{Duration, Instant};
use windows::Win32::{
    Foundation::SYSTEMTIME,
    System::{
        Com::{DISPATCH_METHOD, IDispatch},
        SystemInformation::GetLocalTime,
        Variant::{VAR_CHANGE_FLAGS, VARIANT, VT_R8, VariantChangeType, VariantTimeToSystemTime},
    },
};

const OL_FOLDER_CALENDAR: i32 = 9;
const OL_FOLDER_TASKS: i32 = 13;
const OL_CLASS_APPOINTMENT: i32 = 26;
const OL_CLASS_TASK: i32 = 48;
const MAX_ITEMS_SCANNED: usize = 5000;
const BUDGET: Duration = Duration::from_secs(30);
const BODY_CHARS: usize = 4000;
const PREVIEW_CHARS: usize = 200;
const MAX_ATTENDEES: i32 = 100;
/// Outlook reports "no date" as the year 4501.
const NO_DATE_YEAR: u16 = 4500;

fn nth(collection: &IDispatch, index: i32) -> Result<IDispatch> {
    win::method_object(collection, "Item", vec![VARIANT::from(index)])
}

fn local_now() -> LocalTime {
    // SAFETY: GetLocalTime has no preconditions.
    let t = unsafe { GetLocalTime() };
    LocalTime::from_parts(
        i32::from(t.wYear),
        u32::from(t.wMonth),
        u32::from(t.wDay),
        u32::from(t.wHour),
        u32::from(t.wMinute),
        u32::from(t.wSecond),
    )
}

/// A DATE property as local ISO-8601 text, or None when unset or unreadable.
fn date_iso(object: &IDispatch, name: &str) -> Option<String> {
    let value = win::property(object, name).ok()?;
    let mut converted = VARIANT::default();
    // SAFETY: both VARIANTs are live locals; VT_DATE coerces to its VT_R8 value.
    unsafe { VariantChangeType(&mut converted, &value, VAR_CHANGE_FLAGS(0), VT_R8) }.ok()?;
    let vtime = f64::try_from(&converted).ok()?;
    let mut st = SYSTEMTIME::default();
    // SAFETY: `st` is a live local.
    if unsafe { VariantTimeToSystemTime(vtime, &mut st) } == 0 || st.wYear >= NO_DATE_YEAR {
        return None;
    }
    Some(
        LocalTime::from_parts(
            i32::from(st.wYear),
            u32::from(st.wMonth),
            u32::from(st.wDay),
            u32::from(st.wHour),
            u32::from(st.wMinute),
            u32::from(st.wSecond),
        )
        .iso(),
    )
}

fn capped(text: String, max: usize) -> (String, bool) {
    let text = text.replace("\r\n", "\n");
    let truncated = text.chars().count() > max;
    (text.chars().take(max).collect(), truncated)
}

fn event_summary(event: &IDispatch, store_id: &str) -> Value {
    let preview = capped(
        optional_text(event, "Body").trim().to_string(),
        PREVIEW_CHARS,
    )
    .0;
    json!({
        "entry_id": optional_text(event, "EntryID"),
        "store_id": store_id,
        "subject": optional_text(event, "Subject"),
        "start": date_iso(event, "Start"),
        "end": date_iso(event, "End"),
        "location": optional_text(event, "Location"),
        "organizer": optional_text(event, "Organizer"),
        "is_recurring": win::int_property(event, "RecurrenceState").unwrap_or(0) != 0,
        "all_day": bool_property(event, "AllDayEvent"),
        "preview": preview,
    })
}

pub(super) fn list_events(namespace: &IDispatch, params: ListEventsParams) -> Result<Value> {
    let limit = checked_limit_in(params.limit, 50, 200).map_err(anyhow::Error::msg)?;
    let start = match params.start.as_deref() {
        Some(text) => LocalTime::parse(text).map_err(anyhow::Error::msg)?,
        None => local_now(),
    };
    let end = match params.end.as_deref() {
        Some(text) => LocalTime::parse(text).map_err(anyhow::Error::msg)?,
        None => start.add_days(14),
    };
    let minutes = end.minutes() - start.minutes();
    if minutes <= 0 {
        bail!("end must be after start");
    }
    if minutes > MAX_EVENT_RANGE_DAYS * 1440 {
        bail!("date range may not exceed {MAX_EVENT_RANGE_DAYS} days");
    }

    let calendar = default_folder(namespace, OL_FOLDER_CALENDAR)?;
    let store_id = optional_text(&calendar, "StoreID");
    let items = win::object_property(&calendar, "Items")?;
    // Outlook needs the collection sorted by [Start] BEFORE IncludeRecurrences
    // is set, or recurring series expand incorrectly. The window below bounds
    // the expansion; this only configures the in-memory collection.
    win::invoke(
        &items,
        "Sort",
        DISPATCH_METHOD,
        vec![VARIANT::from("[Start]")],
    )?;
    if params.include_recurrences.unwrap_or(true) {
        win::put_property(&items, "IncludeRecurrences", VARIANT::from(true))?;
    }
    let filter = format!(
        "[Start] >= '{}' AND [Start] <= '{}'",
        start.jet(),
        end.jet()
    );
    let filtered = win::method_object(&items, "Restrict", vec![VARIANT::from(filter.as_str())])
        .context("calendar Restrict failed")?;

    let deadline = Instant::now() + BUDGET;
    let mut rows = Vec::new();
    let (mut has_more, mut scan_truncated) = (false, false);
    // GetFirst/GetNext: Count is unreliable once recurrences are expanded.
    let mut current = win::method_object(&filtered, "GetFirst", vec![]).ok();
    while let Some(event) = current {
        if rows.len() >= limit {
            has_more = true;
            break;
        }
        if Instant::now() > deadline {
            scan_truncated = true;
            break;
        }
        rows.push(event_summary(&event, &store_id));
        current = win::method_object(&filtered, "GetNext", vec![]).ok();
    }
    Ok(
        json!({"start": start.iso(), "end": end.iso(), "count": rows.len(),
        "has_more": has_more, "scan_truncated": scan_truncated, "items": rows}),
    )
}

pub(super) fn get_event(namespace: &IDispatch, params: GetEventParams) -> Result<Value> {
    let mut args = vec![VARIANT::from(params.entry_id.as_str())];
    if let Some(ref store) = params.store_id {
        args.push(VARIANT::from(store.as_str()));
    }
    let event = win::method_object(namespace, "GetItemFromID", args)
        .context("Event no longer available; EntryID may be stale")?;
    if win::int_property(&event, "Class").ok() != Some(OL_CLASS_APPOINTMENT) {
        bail!("EntryID does not identify a calendar event");
    }
    let mut data = event_summary(&event, params.store_id.as_deref().unwrap_or(""));

    let mut attendees = Vec::new();
    if let Ok(recipients) = win::object_property(&event, "Recipients") {
        let count = win::int_property(&recipients, "Count")
            .unwrap_or(0)
            .clamp(0, MAX_ATTENDEES);
        for i in 1..=count {
            let Ok(r) = nth(&recipients, i) else { continue };
            attendees.push(json!({
                "name": optional_text(&r, "Name"),
                "address": optional_text(&r, "Address"),
                "type": win::int_property(&r, "Type").unwrap_or(0),
                "response": win::int_property(&r, "MeetingResponseStatus").unwrap_or(0),
            }));
        }
    }
    let (body, body_truncated) = capped(optional_text(&event, "Body"), BODY_CHARS);
    let object = data.as_object_mut().context("invalid event result")?;
    object.insert("body".into(), json!(body));
    object.insert("body_truncated".into(), json!(body_truncated));
    object.insert("attendees".into(), json!(attendees));
    object.insert(
        "reminder_minutes".into(),
        json!(win::int_property(&event, "ReminderMinutesBeforeStart").ok()),
    );
    object.insert(
        "categories".into(),
        json!(optional_text(&event, "Categories")),
    );
    Ok(data)
}

pub(super) fn list_tasks(namespace: &IDispatch, params: ListTasksParams) -> Result<Value> {
    let limit = checked_limit_in(params.limit, 50, 200).map_err(anyhow::Error::msg)?;
    let include_completed = params.include_completed.unwrap_or(false);
    let folder = default_folder(namespace, OL_FOLDER_TASKS)?;
    let store_id = optional_text(&folder, "StoreID");
    let items = win::object_property(&folder, "Items")?;
    let _ = win::invoke(
        &items,
        "Sort",
        DISPATCH_METHOD,
        vec![VARIANT::from("[DueDate]")],
    );
    let count = win::int_property(&items, "Count")?.max(0);
    let deadline = Instant::now() + BUDGET;
    let (mut rows, mut scanned) = (Vec::new(), 0usize);
    let (mut has_more, mut scan_truncated) = (false, false);
    for i in 1..=count {
        if scanned >= MAX_ITEMS_SCANNED || Instant::now() > deadline {
            scan_truncated = true;
            break;
        }
        scanned += 1;
        let Ok(task) = nth(&items, i) else { continue };
        if win::int_property(&task, "Class").ok() != Some(OL_CLASS_TASK) {
            continue;
        }
        let complete = bool_property(&task, "Complete");
        if complete && !include_completed {
            continue;
        }
        if rows.len() >= limit {
            has_more = true;
            break;
        }
        rows.push(json!({
            "entry_id": optional_text(&task, "EntryID"),
            "store_id": store_id,
            "subject": optional_text(&task, "Subject"),
            "due_date": date_iso(&task, "DueDate"),
            "start_date": date_iso(&task, "StartDate"),
            "complete": complete,
            "percent_complete": win::int_property(&task, "PercentComplete").unwrap_or(0),
            "importance": win::int_property(&task, "Importance").ok(),
            "status": win::int_property(&task, "Status").ok(),
        }));
    }
    Ok(json!({"count": rows.len(), "has_more": has_more,
        "scan_truncated": scan_truncated, "items": rows}))
}
