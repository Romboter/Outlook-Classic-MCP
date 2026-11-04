//! Read-only calendar and task reads.
//!
//! Dates are Outlook local wall-clock values rendered as ISO-8601 without an
//! offset. Every call is bounded by a result limit, a scan cap and a deadline.
use super::mailbox::{bool_property, default_folder, optional_text};
use super::win;
use crate::contracts::{GetEventParams, ListEventsParams, ListTasksParams, checked_limit_in};
use crate::localtime::LocalTime;
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::time::{Duration, Instant};
use windows::Win32::{
    Foundation::SYSTEMTIME,
    System::{
        Com::{DISPATCH_METHOD, IDispatch},
        SystemInformation::GetLocalTime,
        Time::{
            GetTimeZoneInformation, SystemTimeToTzSpecificLocalTime, TIME_ZONE_INFORMATION,
            TzSpecificLocalTimeToSystemTime,
        },
        Variant::{
            SystemTimeToVariantTime, VAR_CHANGE_FLAGS, VARIANT, VT_DATE, VT_R8, VariantChangeType,
            VariantTimeToSystemTime,
        },
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
const MAX_ATTENDEES: i32 = 500;
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

fn to_system_time(t: &LocalTime) -> SYSTEMTIME {
    SYSTEMTIME {
        wYear: t.year as u16,
        wMonth: t.month as u16,
        wDay: t.day as u16,
        wHour: t.hour as u16,
        wMinute: t.minute as u16,
        wSecond: t.second as u16,
        ..Default::default()
    }
}

fn from_system_time(t: &SYSTEMTIME) -> LocalTime {
    LocalTime::from_parts(
        i32::from(t.wYear),
        u32::from(t.wMonth),
        u32::from(t.wDay),
        u32::from(t.wHour),
        u32::from(t.wMinute),
        u32::from(t.wSecond),
    )
}

/// Convert a UTC wall clock to this machine's local time.
fn utc_to_local(utc: &LocalTime) -> Result<LocalTime> {
    let (src, mut dst) = (to_system_time(utc), SYSTEMTIME::default());
    // SAFETY: both SYSTEMTIMEs are live locals; None selects the current zone.
    unsafe { SystemTimeToTzSpecificLocalTime(None, &src, &mut dst) }
        .context("local time conversion failed")?;
    Ok(from_system_time(&dst))
}

/// This machine's UTC offset (DST-correct) at a local wall-clock time, as
/// `+HH:MM`. Empty when the zone cannot be resolved.
fn offset_suffix(local: &LocalTime) -> String {
    let (src, mut dst) = (to_system_time(local), SYSTEMTIME::default());
    // SAFETY: both SYSTEMTIMEs are live locals; None selects the current zone.
    if unsafe { TzSpecificLocalTimeToSystemTime(None, &src, &mut dst) }.is_err() {
        return String::new();
    }
    let minutes = local.minutes() - from_system_time(&dst).minutes();
    let sign = if minutes < 0 { '-' } else { '+' };
    format!("{sign}{:02}:{:02}", minutes.abs() / 60, minutes.abs() % 60)
}

/// This machine's clock: local ISO time with offset, zone name and `+HHMM` offset.
pub(super) fn clock_info() -> (String, String, String) {
    let now = local_now();
    let offset = offset_suffix(&now);
    let mut tz = TIME_ZONE_INFORMATION::default();
    // SAFETY: `tz` is a live local.
    let id = unsafe { GetTimeZoneInformation(&mut tz) };
    let raw = if id == 2 {
        &tz.DaylightName
    } else {
        &tz.StandardName
    };
    let len = raw.iter().position(|&c| c == 0).unwrap_or(raw.len());
    (
        format!("{}{offset}", now.iso()),
        String::from_utf16_lossy(&raw[..len]),
        offset.replace(':', ""),
    )
}

/// Bound given as local time or with an offset, resolved to local wall clock.
pub(super) fn resolve_bound(text: &str) -> Result<LocalTime> {
    let (time, offset) = LocalTime::parse_with_offset(text).map_err(anyhow::Error::msg)?;
    match offset {
        None => Ok(time),
        Some(minutes) => utc_to_local(&time.shift_minutes(-i64::from(minutes))),
    }
}

/// An Outlook DATE (`VT_DATE`) for a local wall-clock time.
pub(super) fn date_variant(time: &LocalTime) -> Result<VARIANT> {
    if !(1900..=4000).contains(&time.year) {
        bail!("year must be between 1900 and 4000");
    }
    let st = to_system_time(time);
    let mut vtime = 0f64;
    // SAFETY: `st` and `vtime` are live locals.
    if unsafe { SystemTimeToVariantTime(&st, &mut vtime) } == 0 {
        bail!("date is out of range");
    }
    let source = VARIANT::from(vtime);
    let mut date = VARIANT::default();
    // SAFETY: both VARIANTs are live locals; VT_R8 coerces to VT_DATE.
    unsafe { VariantChangeType(&mut date, &source, VAR_CHANGE_FLAGS(0), VT_DATE) }
        .context("date conversion failed")?;
    Ok(date)
}

/// A DATE property as local ISO-8601 text, or None when unset or unreadable.
pub(super) fn date_iso(object: &IDispatch, name: &str) -> Option<String> {
    variant_iso(&win::property(object, name).ok()?)
}

/// A DATE value (item property or table cell) as local ISO-8601 text with offset.
pub(super) fn variant_iso(value: &VARIANT) -> Option<String> {
    let mut converted = VARIANT::default();
    // SAFETY: both VARIANTs are live locals; VT_DATE coerces to its VT_R8 value.
    unsafe { VariantChangeType(&mut converted, value, VAR_CHANGE_FLAGS(0), VT_R8) }.ok()?;
    let vtime = f64::try_from(&converted).ok()?;
    let mut st = SYSTEMTIME::default();
    // SAFETY: `st` is a live local.
    if unsafe { VariantTimeToSystemTime(vtime, &mut st) } == 0 || st.wYear >= NO_DATE_YEAR {
        return None;
    }
    let local = from_system_time(&st);
    Some(format!("{}{}", local.iso(), offset_suffix(&local)))
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
        Some(text) => resolve_bound(text)?,
        None => local_now(),
    };
    let end = match params.end.as_deref() {
        Some(text) => resolve_bound(text)?,
        None => start.add_days(14),
    };
    let minutes = end.minutes() - start.minutes();
    if minutes <= 0 {
        bail!("end must be after start");
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

/// Hosts whose https links are accepted as meeting joins: exact or a subdomain.
const JOIN_HOSTS: [&str; 6] = [
    "teams.microsoft.com",
    "teams.live.com",
    "meet.google.com",
    "zoom.us",
    "zoomgov.com",
    "webex.com",
];
const MAX_LINK_CHARS: usize = 500;

fn is_join_host(host: &str) -> bool {
    JOIN_HOSTS
        .iter()
        .any(|h| host == *h || host.strip_suffix(h).is_some_and(|p| p.ends_with('.')))
}

/// The first Teams/Zoom/Meet/Webex https join link in `body`, or None. Pure
/// text matching: the host must be on the list exactly (or a subdomain of a
/// listed domain), with no userinfo and a non-empty path.
pub(super) fn extract_join_link(body: &str) -> Option<String> {
    let lower = body.to_ascii_lowercase();
    let mut from = 0;
    while let Some(at) = lower[from..].find("https://") {
        let begin = from + at;
        from = begin + "https://".len();
        let rest = &body[begin..];
        let end = rest
            .find(|c: char| c.is_whitespace() || "<>\"'()[]{}|^`".contains(c) || c.is_control())
            .unwrap_or(rest.len());
        let url = rest[..end].trim_end_matches(['.', ',', ';', ':', '!', '?']);
        if url.len() > MAX_LINK_CHARS {
            continue;
        }
        let after = &url["https://".len()..];
        let authority_end = after.find(['/', '?', '#']).unwrap_or(after.len());
        let (authority, path) = after.split_at(authority_end);
        let host = authority.strip_suffix(":443").unwrap_or(authority);
        let host = host.to_ascii_lowercase();
        let clean_host = !host.is_empty()
            && host
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || c == '.' || c == '-');
        if clean_host && is_join_host(&host) && path.len() > 1 && path.starts_with('/') {
            return Some(url.to_string());
        }
    }
    None
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
    let body_limit = params.max_body_chars.unwrap_or(BODY_CHARS as u32) as usize;
    let raw_body = optional_text(&event, "Body");
    let join_link = extract_join_link(&raw_body);
    let (body, body_truncated) = capped(raw_body, body_limit);
    let object = data.as_object_mut().context("invalid event result")?;
    // Classic's object model has no verified online-meeting property, so the
    // link only ever comes from the (untrusted) body. Never fetched or logged.
    object.insert("online_meeting_link".into(), json!(join_link));
    object.insert(
        "online_meeting_link_source".into(),
        json!(join_link.as_ref().map(|_| "body")),
    );
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
    // OlMeetingStatus: 0 none, 1 meeting, 3 received, 5 canceled, 7 received+canceled.
    object.insert(
        "meeting_status".into(),
        json!(win::int_property(&event, "MeetingStatus").ok()),
    );
    // OlResponseStatus for the mailbox owner: 0 none, 1 organized, 2 tentative, 3 accepted, 4 declined, 5 not responded.
    object.insert(
        "my_response".into(),
        json!(win::int_property(&event, "ResponseStatus").ok()),
    );
    object.insert(
        "organizer".into(),
        json!(optional_text(&event, "Organizer")),
    );
    let recurring = super::mailbox::bool_property(&event, "IsRecurring");
    object.insert("is_recurring".into(), json!(recurring));
    if recurring && let Ok(pattern) = win::method_object(&event, "GetRecurrencePattern", vec![]) {
        // OlRecurrenceType: 0 daily, 1 weekly, 2 monthly, 3 month-nth, 5 yearly, 6 year-nth.
        object.insert(
            "recurrence".into(),
            json!({
                "type": win::int_property(&pattern, "RecurrenceType").ok(),
                "interval": win::int_property(&pattern, "Interval").ok(),
                "days_of_week_mask": win::int_property(&pattern, "DaysOfWeekMask").ok(),
                "no_end_date": super::mailbox::bool_property(&pattern, "NoEndDate"),
                "occurrences": win::int_property(&pattern, "Occurrences").ok(),
            }),
        );
    }
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

#[cfg(test)]
mod tests {
    use super::extract_join_link;

    #[test]
    fn finds_join_links_on_known_hosts() {
        let teams = "Join: https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0?context=x";
        assert_eq!(
            extract_join_link(teams).as_deref(),
            Some("https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0?context=x")
        );
        assert_eq!(
            extract_join_link("<https://us02web.zoom.us/j/123456789?pwd=abc>.").as_deref(),
            Some("https://us02web.zoom.us/j/123456789?pwd=abc")
        );
        assert_eq!(
            extract_join_link("Meet: https://meet.google.com/abc-defg-hij, see you").as_deref(),
            Some("https://meet.google.com/abc-defg-hij")
        );
        assert_eq!(
            extract_join_link("https://acme.webex.com/meet/jane").as_deref(),
            Some("https://acme.webex.com/meet/jane")
        );
        // The first acceptable link wins, skipping unrelated ones.
        assert_eq!(
            extract_join_link("see https://example.org/a then HTTPS://TEAMS.LIVE.COM/meet/9")
                .as_deref(),
            Some("HTTPS://TEAMS.LIVE.COM/meet/9")
        );
    }

    #[test]
    fn rejects_missing_and_look_alike_links() {
        for body in [
            "",
            "no link here",
            "http://teams.microsoft.com/l/meetup-join/x",
            "https://teams.microsoft.com.evil.example/l/meetup-join/x",
            "https://evilteams.microsoft.com.example/x",
            "https://evilzoom.us/j/1",
            "https://zoom.us.evil.example/j/1",
            "https://teams.microsoft.com@evil.example/l/x",
            "https://evil.example/?u=https://teams.microsoft.com",
            "https://teams.microsoft.com",
            "https://teams.microsoft.com/",
            "https://teams.microsoft.com:8443/l/x",
        ] {
            assert_eq!(extract_join_link(body), None, "{body}");
        }
    }
}
