//! Read-only MCP request contracts. Values are validated again on the server
//! side; JSON Schema is a usability hint, not an authorization boundary.
use crate::action::{ActionKind, ActionPreview, ActionTarget};
use crate::localtime::LocalTime;
use crate::policy::Capability;
use crate::send::{DraftRef, PrepareRequest, Preview};
use rmcp::schemars::JsonSchema;
use serde::Deserialize;

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct FolderParams {
    /// Folder path. Omit to start at the default mailbox root.
    pub root: Option<String>,
    /// Maximum subfolder traversal depth; defaults to 2, capped at 4.
    pub max_depth: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ListMailsParams {
    /// Inbox, Sent Items, Drafts, or a relative Inbox/subfolder path.
    pub folder: Option<String>,
    /// Number of messages to return (default 10, maximum 30).
    pub limit: Option<u32>,
    /// Number of matching messages to skip (default 0, maximum 1000).
    pub offset: Option<u32>,
    /// Limit results to unread mail.
    pub unread_only: Option<bool>,
    /// Filter by sender address or display name (substring).
    pub from_address: Option<String>,
    /// Local wall-clock lower bound on received time (`YYYY-MM-DD[THH:MM[:SS]]`), inclusive.
    pub since: Option<String>,
    /// Local wall-clock upper bound on received time, inclusive.
    pub until: Option<String>,
    /// Force auto, table or legacy backend to compare identical queries.
    pub backend: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SearchMailsParams {
    /// Search text, never raw DASL.
    pub query: String,
    /// Folder to search, defaults to Inbox.
    pub folder: Option<String>,
    /// Maximum result count (default 10, maximum 30).
    pub limit: Option<u32>,
    /// Search fields: subject, from, or subject_body (default).
    pub scope: Option<String>,
    /// Backend: auto (default), table, legacy, indexed. Indexed uses experimental event-driven AdvancedSearch.
    pub backend: Option<String>,
    /// Include subfolders when backend=indexed (default false). Other backends reject this option.
    pub search_subfolders: Option<bool>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct GetMailParams {
    /// Short session-scoped entry_id handle from list/search (e.g. m1), or a raw Outlook EntryID for compatibility.
    pub entry_id: String,
    /// Outlook StoreID when reading an item from a non-default store.
    pub store_id: Option<String>,
    /// Include plain-text email body (default true).
    pub include_body: Option<bool>,
    /// Also return the raw HTML source as `html_body` (default false). It is
    /// untrusted text, never rendered or fetched, and uses the same
    /// `max_body_chars` / `body_offset` window as the plain body.
    pub include_html: Option<bool>,
    /// Body budget per field (default 4000, maximum 12000 characters).
    pub max_body_chars: Option<u32>,
    /// Start the body at this character offset (default 0). Use the previous
    /// result's `body_next_offset` to read a long body in pages.
    pub body_offset: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ThreadParams {
    /// Short handle from search/list or raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// Default 10, range 1..30.
    pub max_messages: Option<u32>,
    /// Include plaintext bodies, default true.
    pub include_bodies: Option<bool>,
    /// Omit recognized quoted history, default true.
    pub trim_quoted: Option<bool>,
    /// Per-message Unicode character limit, default 1500, max 6000.
    pub per_message_chars: Option<u32>,
    /// Total Unicode character limit, default 8000, max 24000.
    pub total_body_chars: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ListContactsParams {
    /// Number of contacts to return (default 50, maximum 200).
    pub limit: Option<u32>,
    /// Number of contacts to skip across all contact folders (default 0, maximum 5000).
    pub offset: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SearchContactsParams {
    /// Words that must all appear in name, email, company or job title (1..160 characters).
    pub query: String,
    /// Maximum result count (default 25, maximum 100).
    pub limit: Option<u32>,
    /// Also scan the Exchange Global Address List by display name (default true).
    pub include_directory: Option<bool>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct GetContactParams {
    /// Short handle from list/search contacts (e.g. m1), or a raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// Notes character budget (default 4000, maximum 50000).
    pub max_notes_chars: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ResolveNameParams {
    /// Display name or address to resolve against the address book (1..160 characters).
    pub name: String,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ListEventsParams {
    /// ISO-8601 start of range, e.g. 2026-04-25T09:00 (local) or 2026-04-25T09:00+02:00. Defaults to now.
    pub start: Option<String>,
    /// ISO-8601 end of range. Defaults to start + 14 days.
    pub end: Option<String>,
    /// Maximum events (default 50, maximum 200).
    pub limit: Option<u32>,
    /// Expand recurring series into occurrences (default true).
    pub include_recurrences: Option<bool>,
    /// Calendar folder: same spec as the mail `folder` parameter (alias, `<store>/Path` prefix or path). Must hold appointments. Defaults to your own calendar.
    pub calendar: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct GetEventParams {
    /// Short handle from list_events (e.g. m1), or a raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// Body character budget (default 4000, maximum 50000).
    pub max_body_chars: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct GetTaskParams {
    /// Short handle from list_tasks (e.g. m1), or a raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// Body character budget (default 4000, maximum 50000).
    pub max_body_chars: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct UpdateTaskParams {
    /// Short handle from list_tasks, or a raw Outlook EntryID.
    pub entry_id: String,
    pub store_id: Option<String>,
    pub subject: Option<String>,
    /// Due date, local ISO-8601 (2026-05-01) or with an offset.
    pub due_date: Option<String>,
    /// Replaces the task notes.
    pub body: Option<String>,
    /// low, normal or high.
    pub importance: Option<String>,
    /// 0..100. 100 marks the task complete.
    pub percent_complete: Option<u32>,
    /// not_started, in_progress, complete, waiting or deferred. Not combinable with percent_complete.
    pub status: Option<String>,
    /// Reminder time, local ISO-8601 or with an offset (turns the reminder on).
    pub reminder: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ListTasksParams {
    /// Maximum tasks (default 50, maximum 200).
    pub limit: Option<u32>,
    /// Include completed tasks (default false).
    pub include_completed: Option<bool>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct MarkMailParams {
    /// Short handle from list/search (e.g. m1), or a raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// true marks read, false marks unread, omit for no change.
    pub read: Option<bool>,
    /// true flags for follow-up, false clears the flag, omit for no change.
    pub flagged: Option<bool>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SetCategoryParams {
    /// Short handle from list/search, or a raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// Semicolon-separated category names that REPLACE the item's categories. Empty clears them.
    pub categories: String,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CreateFolderParams {
    /// Name of the new folder (1..120 characters, no slashes).
    pub name: String,
    /// Parent folder. Defaults to Inbox.
    pub parent: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct MoveMailParams {
    /// Short handle from list/search, or a raw Outlook EntryID.
    pub entry_id: String,
    /// StoreID for raw identifiers in non-default stores.
    pub store_id: Option<String>,
    /// Destination folder: Inbox, Sent Items, Drafts or a relative path. Deleted Items is refused.
    pub target_folder: String,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SendMailParams {
    /// Recipient addresses or names Outlook can resolve.
    pub to: Vec<String>,
    pub subject: String,
    /// Plain-text body.
    pub body: String,
    pub cc: Option<Vec<String>>,
    pub bcc: Option<Vec<String>>,
    /// Treat `body` as HTML (default false). The approval prompt says so and shows a plain-text rendering.
    pub html: Option<bool>,
    /// Paths of files to attach: existing regular files inside the user profile, at most 10 and 20 MiB in total; executables and scripts are refused. The approval prompt lists each file.
    pub attachments: Option<Vec<String>>,
    /// low, normal (default) or high.
    pub importance: Option<String>,
    /// Only save to Drafts; never send.
    pub save_only: Option<bool>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ReplyMailParams {
    /// Short handle from list/search, or a raw Outlook EntryID.
    pub entry_id: String,
    pub store_id: Option<String>,
    /// Plain-text reply, placed above the quoted original.
    pub body: String,
    /// Reply to everyone (default false).
    pub reply_all: Option<bool>,
    /// Treat `body` as HTML (default false).
    pub html: Option<bool>,
    /// Files to attach (same rules and limits as `outlook_send_mail`).
    pub attachments: Option<Vec<String>>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ForwardMailParams {
    /// Short handle from list/search, or a raw Outlook EntryID.
    pub entry_id: String,
    pub store_id: Option<String>,
    pub to: Vec<String>,
    pub cc: Option<Vec<String>>,
    /// Optional plain-text note above the forwarded message.
    pub body: Option<String>,
    /// Treat `body` as HTML (default false).
    pub html: Option<bool>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SendDraftParams {
    /// Handle of a draft this server saved (the entry_id it returned).
    pub entry_id: String,
    pub store_id: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CheckSendParams {
    /// The send_id returned by a send, save or unknown_outcome result.
    pub send_id: String,
}

fn checked_addresses(list: &[String], what: &str) -> Result<(), String> {
    if list.len() > 500 {
        return Err(format!("{what} has too many entries"));
    }
    for entry in list {
        if entry.trim().is_empty()
            || entry.chars().count() > 320
            || entry.chars().any(|c| c.is_control() || c == ';')
        {
            return Err(format!("{what} contains an invalid entry"));
        }
    }
    Ok(())
}

/// Validate outgoing attachment paths (user profile only) and return the
/// canonical paths to attach.
fn checked_attachments(attachments: Option<&Vec<String>>) -> Result<Vec<String>, String> {
    let requested = attachments.map(Vec::as_slice).unwrap_or_default();
    crate::safe_path::resolve_attachment_files(requested, &crate::safe_path::default_roots())
        .map(|files| files.into_iter().map(|f| f.path).collect())
}

fn checked_body(body: &str) -> Result<(), String> {
    if body.chars().count() > 100_000 {
        return Err("body must be at most 100000 characters".into());
    }
    Ok(())
}

pub fn prepare_send_mail(x: &SendMailParams) -> Result<PrepareRequest, String> {
    let attachments = checked_attachments(x.attachments.as_ref())?;
    if x.to.is_empty() {
        return Err("at least one recipient is required".into());
    }
    let (cc, bcc) = (
        x.cc.clone().unwrap_or_default(),
        x.bcc.clone().unwrap_or_default(),
    );
    for (list, what) in [(&x.to, "to"), (&cc, "cc"), (&bcc, "bcc")] {
        checked_addresses(list, what)?;
    }
    if x.subject.chars().count() > 255 || x.subject.chars().any(char::is_control) {
        return Err("subject must be at most 255 characters on one line".into());
    }
    checked_body(&x.body)?;
    let importance = x
        .importance
        .clone()
        .unwrap_or_else(|| "normal".into())
        .to_ascii_lowercase();
    if !matches!(importance.as_str(), "low" | "normal" | "high") {
        return Err("importance must be low, normal or high".into());
    }
    Ok(PrepareRequest::New {
        to: x.to.clone(),
        cc,
        bcc,
        subject: x.subject.clone(),
        body: x.body.clone(),
        importance,
        html: x.html.unwrap_or(false),
        attachments,
    })
}

pub fn prepare_reply_mail(x: &ReplyMailParams) -> Result<PrepareRequest, String> {
    let attachments = checked_attachments(x.attachments.as_ref())?;
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    checked_body(&x.body)?;
    Ok(PrepareRequest::Reply {
        draft_of: DraftRef {
            entry_id: x.entry_id.clone(),
            store_id: x.store_id.clone(),
        },
        body: x.body.clone(),
        reply_all: x.reply_all.unwrap_or(false),
        html: x.html.unwrap_or(false),
        attachments,
    })
}

pub fn prepare_forward_mail(x: &ForwardMailParams) -> Result<PrepareRequest, String> {
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    if x.to.is_empty() {
        return Err("at least one recipient is required".into());
    }
    let cc = x.cc.clone().unwrap_or_default();
    checked_addresses(&x.to, "to")?;
    checked_addresses(&cc, "cc")?;
    let body = x.body.clone().unwrap_or_default();
    checked_body(&body)?;
    Ok(PrepareRequest::Forward {
        draft_of: DraftRef {
            entry_id: x.entry_id.clone(),
            store_id: x.store_id.clone(),
        },
        to: x.to.clone(),
        cc,
        body,
        html: x.html.unwrap_or(false),
    })
}

pub fn checked_send_id(send_id: &str) -> Result<(), String> {
    if send_id.is_empty()
        || send_id.len() > 80
        || !send_id
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-')
    {
        return Err("invalid send_id".into());
    }
    Ok(())
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CreateTaskParams {
    pub subject: String,
    /// Due date, local ISO-8601 (2026-05-01) or with an offset.
    pub due_date: Option<String>,
    pub body: Option<String>,
    /// low, normal (default) or high.
    pub importance: Option<String>,
    /// Reminder time, local ISO-8601 or with an offset.
    pub reminder: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CompleteTaskParams {
    /// Short handle from list_tasks, or a raw Outlook EntryID.
    pub entry_id: String,
    pub store_id: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct RecurrenceParams {
    /// daily, weekly, monthly or yearly.
    #[serde(rename = "type")]
    pub kind: String,
    /// Every N units (default 1).
    pub interval: Option<u32>,
    /// End after N occurrences (alternative to end_date).
    pub occurrences: Option<u32>,
    /// ISO-8601 date the series ends (alternative to occurrences).
    pub end_date: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CreateEventParams {
    pub subject: String,
    /// Local ISO-8601 start, or with an offset (converted to local).
    pub start: String,
    pub end: String,
    pub location: Option<String>,
    pub body: Option<String>,
    /// Plain email addresses to invite. Non-empty makes this a meeting that needs approval (send permission).
    pub attendees: Option<Vec<String>>,
    /// Online meetings are not supported yet; must be false or omitted.
    pub is_online_meeting: Option<bool>,
    /// Minutes before start (default 15).
    pub reminder_minutes: Option<u32>,
    /// Create the event with no reminder (default false). Cannot be combined with reminder_minutes.
    pub no_reminder: Option<bool>,
    pub recurrence: Option<RecurrenceParams>,
    /// Calendar folder to create the event in: same spec as the mail `folder` parameter. Must hold appointments. Defaults to your own calendar. Not allowed together with attendees.
    pub calendar: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct UpdateEventParams {
    /// Short handle from list_events, or a raw Outlook EntryID.
    pub entry_id: String,
    pub store_id: Option<String>,
    pub subject: Option<String>,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: Option<String>,
    pub body: Option<String>,
}

fn checked_free_text(value: &str, what: &str, max: usize) -> Result<(), String> {
    if value.chars().count() > max {
        return Err(format!("{what} must be at most {max} characters"));
    }
    Ok(())
}

fn checked_title(value: &str, what: &str) -> Result<(), String> {
    if value.trim().is_empty() || value.chars().count() > 255 || value.chars().any(char::is_control)
    {
        return Err(format!("{what} must be 1..255 characters on one line"));
    }
    Ok(())
}

fn parsed(text: &str, what: &str) -> Result<LocalTime, String> {
    let (time, offset) = LocalTime::parse_with_offset(text).map_err(|e| format!("{what}: {e}"))?;
    // Offsets are converted to local time by the Outlook layer; for ordering
    // here compare on UTC-equivalent minutes.
    Ok(time.shift_minutes(-i64::from(offset.unwrap_or(0))))
}

pub fn validate_create_task(x: &CreateTaskParams) -> Result<(), String> {
    checked_title(&x.subject, "subject")?;
    checked_free_text(x.body.as_deref().unwrap_or(""), "body", 100_000)?;
    if let Some(level) = &x.importance
        && !matches!(
            level.to_ascii_lowercase().as_str(),
            "low" | "normal" | "high"
        )
    {
        return Err("importance must be low, normal or high".into());
    }
    for (value, what) in [(&x.due_date, "due_date"), (&x.reminder, "reminder")] {
        if let Some(text) = value {
            LocalTime::parse_with_offset(text).map_err(|e| format!("{what}: {e}"))?;
        }
    }
    Ok(())
}

/// A `calendar` folder spec: present means non-blank, one line, at most 300 characters.
pub fn validate_calendar_spec(spec: Option<&str>) -> Result<(), String> {
    if let Some(spec) = spec
        && (spec.trim().is_empty()
            || spec.chars().count() > 300
            || spec.chars().any(char::is_control))
    {
        return Err("calendar must be 1..300 characters on one line".into());
    }
    Ok(())
}

pub const TASK_STATUSES: [&str; 5] = [
    "not_started",
    "in_progress",
    "complete",
    "waiting",
    "deferred",
];

pub fn validate_update_task(x: &UpdateTaskParams) -> Result<(), String> {
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    if let Some(subject) = &x.subject {
        checked_title(subject, "subject")?;
    }
    checked_free_text(x.body.as_deref().unwrap_or(""), "body", 100_000)?;
    if let Some(level) = &x.importance
        && !matches!(
            level.to_ascii_lowercase().as_str(),
            "low" | "normal" | "high"
        )
    {
        return Err("importance must be low, normal or high".into());
    }
    for (value, what) in [(&x.due_date, "due_date"), (&x.reminder, "reminder")] {
        if let Some(text) = value {
            LocalTime::parse_with_offset(text).map_err(|e| format!("{what}: {e}"))?;
        }
    }
    if x.percent_complete.is_some_and(|p| p > 100) {
        return Err("percent_complete must be 0..100".into());
    }
    if let Some(status) = &x.status
        && !TASK_STATUSES.contains(&status.to_ascii_lowercase().as_str())
    {
        return Err(
            "status must be not_started, in_progress, complete, waiting or deferred".into(),
        );
    }
    if x.percent_complete.is_some() && x.status.is_some() {
        return Err("give percent_complete or status, not both".into());
    }
    if x.subject.is_none()
        && x.due_date.is_none()
        && x.body.is_none()
        && x.importance.is_none()
        && x.percent_complete.is_none()
        && x.status.is_none()
        && x.reminder.is_none()
    {
        return Err("nothing to change; give at least one field to update".into());
    }
    Ok(())
}

pub fn validate_create_event(x: &CreateEventParams) -> Result<(), String> {
    checked_title(&x.subject, "subject")?;
    validate_calendar_spec(x.calendar.as_deref())?;
    if x.attendees.as_ref().is_some_and(|a| !a.is_empty()) {
        return Err(
            "inviting attendees sends mail and is not supported yet; create the event without attendees"
                .into(),
        );
    }
    if x.is_online_meeting == Some(true) {
        return Err("online meetings are not supported yet".into());
    }
    let (start, end) = (parsed(&x.start, "start")?, parsed(&x.end, "end")?);
    if end.minutes() <= start.minutes() {
        return Err("end must be after start".into());
    }
    checked_free_text(x.location.as_deref().unwrap_or(""), "location", 255)?;
    checked_free_text(x.body.as_deref().unwrap_or(""), "body", 100_000)?;
    if x.no_reminder.unwrap_or(false) && x.reminder_minutes.is_some() {
        return Err("give reminder_minutes or no_reminder, not both".into());
    }
    if x.reminder_minutes.is_some_and(|m| m > 525_600) {
        return Err("reminder_minutes must be at most 525600".into());
    }
    if let Some(r) = &x.recurrence {
        if !matches!(r.kind.as_str(), "daily" | "weekly" | "monthly" | "yearly") {
            return Err("recurrence type must be daily, weekly, monthly or yearly".into());
        }
        if r.interval.is_some_and(|i| !(1..=999).contains(&i)) {
            return Err("recurrence interval must be 1..999".into());
        }
        if r.occurrences.is_some() && r.end_date.is_some() {
            return Err("give recurrence occurrences or end_date, not both".into());
        }
        if r.occurrences.is_some_and(|o| !(1..=999).contains(&o)) {
            return Err("recurrence occurrences must be 1..999".into());
        }
        if let Some(date) = &r.end_date
            && parsed(date, "recurrence end_date")?.minutes() <= start.minutes()
        {
            return Err("recurrence end_date must be after the start".into());
        }
    }
    Ok(())
}

pub fn validate_update_event(x: &UpdateEventParams) -> Result<(), String> {
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    if x.subject.is_none()
        && x.start.is_none()
        && x.end.is_none()
        && x.location.is_none()
        && x.body.is_none()
    {
        return Err("nothing to change".into());
    }
    if let Some(subject) = &x.subject {
        checked_title(subject, "subject")?;
    }
    for (value, what) in [(&x.start, "start"), (&x.end, "end")] {
        if let Some(text) = value {
            parsed(text, what)?;
        }
    }
    if let (Some(start), Some(end)) = (&x.start, &x.end)
        && parsed(end, "end")?.minutes() <= parsed(start, "start")?.minutes()
    {
        return Err("end must be after start".into());
    }
    checked_free_text(x.location.as_deref().unwrap_or(""), "location", 255)?;
    checked_free_text(x.body.as_deref().unwrap_or(""), "body", 100_000)?;
    Ok(())
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct DeleteItemParams {
    /// Short handle from list/search, or a raw Outlook EntryID.
    pub entry_id: String,
    pub store_id: Option<String>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct CheckActionParams {
    /// The action_id returned by a delete result.
    pub action_id: String,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct RespondEventParams {
    /// Short handle from list_events, or a raw Outlook EntryID, of a meeting invitation you received.
    pub entry_id: String,
    pub store_id: Option<String>,
    /// accept, tentative or decline.
    pub response: String,
    /// Email the answer to the organizer (default true). false only changes your own calendar.
    pub send_response: Option<bool>,
}

pub fn prepare_respond_event(x: &RespondEventParams) -> Result<ActionTarget, String> {
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    let response = x.response.trim().to_ascii_lowercase();
    if !matches!(response.as_str(), "accept" | "tentative" | "decline") {
        return Err("response must be accept, tentative or decline".into());
    }
    Ok(ActionTarget {
        kind: ActionKind::RespondEvent,
        entry_id: x.entry_id.clone(),
        store_id: x.store_id.clone(),
        enabled: Some(x.send_response.unwrap_or(true)),
        response: Some(response),
        payload: None,
    })
}

const MAX_INVITEES: usize = 100;

/// Plain SMTP addresses only: names are ambiguous and would resolve differently
/// at approval time than at send time.
fn checked_attendees(list: &[String]) -> Result<Vec<String>, String> {
    if list.is_empty() || list.len() > MAX_INVITEES {
        return Err(format!("attendees must list 1..{MAX_INVITEES} addresses"));
    }
    let mut seen = Vec::new();
    for raw in list {
        let address = raw.trim().to_ascii_lowercase();
        let valid = address.len() <= 254
            && address.matches('@').count() == 1
            && !address.starts_with('@')
            && !address.ends_with('@')
            && !address
                .chars()
                .any(|c| c.is_whitespace() || c.is_control() || c == ';' || c == ',');
        if !valid {
            return Err(format!("'{raw}' is not a plain email address"));
        }
        if !seen.contains(&address) {
            seen.push(address);
        }
    }
    Ok(seen)
}

/// 0 = Sunday .. 6 = Saturday.
pub fn weekday_of(t: &LocalTime) -> usize {
    const T: [i32; 12] = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
    let y = if t.month < 3 { t.year - 1 } else { t.year };
    ((y + y / 4 - y / 100 + y / 400 + T[(t.month - 1) as usize] + t.day as i32) % 7) as usize
}

const WEEKDAYS: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];
const MONTHS: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

/// Outlook's day-of-week mask bit for the start date (Sunday 1 .. Saturday 64).
pub fn weekday_mask(start: &LocalTime) -> i32 {
    1 << weekday_of(start)
}

/// The recurrence in plain words, for the approval prompt. The series runs on
/// the weekday / day of month / date of the start, which is what is applied.
pub fn describe_recurrence(r: &RecurrenceParams, start: &LocalTime) -> String {
    let n = r.interval.unwrap_or(1);
    let every = |unit: &str| {
        if n == 1 {
            format!("every {unit}")
        } else {
            format!("every {n} {unit}s")
        }
    };
    let pattern = match r.kind.as_str() {
        "daily" => every("day"),
        "weekly" => format!("{} on {}", every("week"), WEEKDAYS[weekday_of(start)]),
        "monthly" => format!("{} on day {}", every("month"), start.day),
        _ => format!(
            "{} on {} {}",
            every("year"),
            start.day,
            MONTHS[(start.month - 1) as usize]
        ),
    };
    let end = match (r.occurrences, &r.end_date) {
        (Some(o), _) => format!("{o} times"),
        (None, Some(d)) => format!("until {}", d.get(..10).unwrap_or(d)),
        (None, None) => "with no end date".into(),
    };
    format!("repeats {pattern}, starting {}, {end}", start.iso())
}

pub fn prepare_invite_meeting(x: &CreateEventParams) -> Result<ActionTarget, String> {
    if x.calendar.is_some() {
        return Err(
            "meetings with attendees can only be created on your own default calendar; omit calendar or omit attendees"
                .into(),
        );
    }
    let attendees = checked_attendees(x.attendees.as_deref().unwrap_or(&[]))?;
    let mut plain = x.clone();
    plain.attendees = None;
    validate_create_event(&plain)?;
    let (start, end) = (parsed(&x.start, "start")?, parsed(&x.end, "end")?);
    let recurrence = x.recurrence.as_ref().map(|r| {
        serde_json::json!({
            "type": r.kind, "interval": r.interval,
            "occurrences": r.occurrences, "end_date": r.end_date,
        })
    });
    let payload = serde_json::json!({
        "subject": x.subject, "start": x.start, "end": x.end,
        "location": x.location, "body": x.body, "attendees": attendees,
        "reminder_minutes": x.reminder_minutes.unwrap_or(15),
        "no_reminder": x.no_reminder.unwrap_or(false),
        "recurrence": recurrence,
    });
    Ok(ActionTarget {
        kind: ActionKind::InviteMeeting,
        // Identifies the request, so an identical retry after an unknown outcome is blocked.
        entry_id: format!(
            "invite|{}|{}|{}|{}|{}",
            x.subject.trim(),
            start.minutes(),
            end.minutes(),
            attendees.join(","),
            x.recurrence
                .as_ref()
                .map(|r| describe_recurrence(r, &start))
                .unwrap_or_default()
        ),
        store_id: None,
        enabled: None,
        response: None,
        payload: Some(payload),
    })
}

pub fn prepare_update_meeting(x: &UpdateEventParams) -> Result<ActionTarget, String> {
    validate_update_event(x)?;
    let payload = serde_json::json!({
        "subject": x.subject, "start": x.start, "end": x.end,
        "location": x.location, "body": x.body,
    });
    Ok(ActionTarget {
        kind: ActionKind::UpdateMeeting,
        entry_id: x.entry_id.clone(),
        store_id: x.store_id.clone(),
        enabled: None,
        response: None,
        payload: Some(payload),
    })
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SaveAttachmentsParams {
    /// Short handle from list/search, or a raw Outlook EntryID, of the mail.
    pub entry_id: String,
    pub store_id: Option<String>,
    /// Absolute path of an existing directory under your user profile.
    pub output_dir: String,
    /// 1-indexed attachment. Omit to save all.
    pub attachment_index: Option<u32>,
}

pub fn prepare_save_attachments(x: &SaveAttachmentsParams) -> Result<ActionTarget, String> {
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    let dir = x.output_dir.trim();
    if dir.is_empty() || dir.chars().count() > 1000 || dir.chars().any(char::is_control) {
        return Err("output_dir must be an absolute path of 1..1000 characters".into());
    }
    if x.attachment_index == Some(0) {
        return Err("attachment_index is 1-indexed".into());
    }
    Ok(ActionTarget {
        kind: ActionKind::SaveAttachments,
        entry_id: x.entry_id.clone(),
        store_id: x.store_id.clone(),
        enabled: None,
        response: None,
        payload: Some(serde_json::json!({
            "output_dir": dir, "index": x.attachment_index,
        })),
    })
}

pub fn prepare_cancel_meeting(x: &DeleteItemParams) -> Result<ActionTarget, String> {
    prepare_delete(x, ActionKind::CancelMeeting)
}

pub fn prepare_delete(x: &DeleteItemParams, kind: ActionKind) -> Result<ActionTarget, String> {
    checked_id(&x.entry_id, x.store_id.as_deref())?;
    Ok(ActionTarget {
        kind,
        entry_id: x.entry_id.clone(),
        store_id: x.store_id.clone(),
        enabled: None,
        response: None,
        payload: None,
    })
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ToggleRuleParams {
    /// Exact name of the rule, as shown by outlook_list_rules.
    pub rule_name: String,
    /// true switches the rule on, false switches it off.
    pub enabled: bool,
}

pub fn prepare_toggle_rule(x: &ToggleRuleParams) -> Result<ActionTarget, String> {
    let name = x.rule_name.trim();
    if name.is_empty()
        || x.rule_name.chars().count() > 255
        || x.rule_name.chars().any(char::is_control)
    {
        return Err("rule_name must be 1..255 characters on one line".into());
    }
    Ok(ActionTarget {
        kind: ActionKind::ToggleRule,
        entry_id: name.to_string(),
        store_id: None,
        enabled: Some(x.enabled),
        response: None,
        payload: None,
    })
}

#[derive(Debug, Clone)]
pub enum MailCommand {
    /// Confirmed-action steps on the Outlook worker; the engine is in `crate::action`.
    ActionPrepare(ActionTarget),
    ActionExecute {
        target: ActionTarget,
        preview: ActionPreview,
    },
    ActionLocate {
        action_id: String,
        target: Option<ActionTarget>,
    },
    CreateTask(CreateTaskParams),
    CompleteTask(CompleteTaskParams),
    UpdateTask(UpdateTaskParams),
    CreateEvent(CreateEventParams),
    UpdateEvent(UpdateEventParams),
    /// Send pipeline steps on the Outlook worker; the pipeline itself is in `crate::send`.
    ComposePrepare(PrepareRequest),
    ComposeInspect(DraftRef),
    ComposeSend {
        draft: DraftRef,
        preview: Preview,
    },
    ComposeLocate(String),
    MarkMail(MarkMailParams),
    SetCategory(SetCategoryParams),
    CreateFolder(CreateFolderParams),
    MoveMail(MoveMailParams),
    ListEvents(ListEventsParams),
    GetEvent(GetEventParams),
    GetTask(GetTaskParams),
    ListTasks(ListTasksParams),
    WhoAmI,
    ListContacts(ListContactsParams),
    SearchContacts(SearchContactsParams),
    GetContact(GetContactParams),
    ResolveName(ResolveNameParams),
    ListCategories,
    ListRules,
    GetOutOfOffice,
    ListFolders(FolderParams),
    ListMails(ListMailsParams),
    SearchMails(SearchMailsParams),
    GetMail(GetMailParams),
    GetThread(ThreadParams),
}

/// Lowercased whitespace-separated query words (never empty for a non-blank query).
pub fn query_words(query: &str) -> Vec<String> {
    query.split_whitespace().map(str::to_lowercase).collect()
}

pub fn matches_all(haystack: &str, words: &[String]) -> bool {
    let haystack = haystack.to_lowercase();
    words.iter().all(|w| haystack.contains(w.as_str()))
}

fn checked_id(entry_id: &str, store_id: Option<&str>) -> Result<(), String> {
    if entry_id.is_empty()
        || entry_id.len() > 4096
        || store_id.is_some_and(|s| s.is_empty() || s.len() > 4096)
    {
        return Err("invalid EntryID/StoreID".into());
    }
    Ok(())
}

fn checked_text(value: &str, what: &str) -> Result<(), String> {
    if value.trim().is_empty() || value.chars().count() > 160 {
        return Err(format!("{what} must contain 1..160 characters"));
    }
    Ok(())
}

/// Limit check for tools whose Python defaults are larger than the mail tools'.
impl MailCommand {
    /// What the command needs permission to do. The match is exhaustive on
    /// purpose: a new command cannot compile until it is classified.
    pub fn capability(&self) -> Capability {
        match self {
            MailCommand::WhoAmI
            | MailCommand::ListEvents(_)
            | MailCommand::GetEvent(_)
            | MailCommand::GetTask(_)
            | MailCommand::ListTasks(_)
            | MailCommand::ListContacts(_)
            | MailCommand::SearchContacts(_)
            | MailCommand::GetContact(_)
            | MailCommand::ResolveName(_)
            | MailCommand::ListCategories
            | MailCommand::ListRules
            | MailCommand::GetOutOfOffice
            | MailCommand::ListFolders(_)
            | MailCommand::ListMails(_)
            | MailCommand::SearchMails(_)
            | MailCommand::GetMail(_)
            | MailCommand::GetThread(_) => Capability::Read,
            MailCommand::ComposeInspect(_)
            | MailCommand::ComposeLocate(_)
            | MailCommand::ActionLocate { .. } => Capability::Read,
            MailCommand::ActionPrepare(target) | MailCommand::ActionExecute { target, .. } => {
                target.capability()
            }
            MailCommand::ComposePrepare(_) => Capability::Draft,
            MailCommand::ComposeSend { .. } => Capability::Send,
            MailCommand::CreateEvent(_) | MailCommand::UpdateEvent(_) => Capability::Events,
            MailCommand::CreateTask(_)
            | MailCommand::CompleteTask(_)
            | MailCommand::UpdateTask(_) => Capability::Modify,
            MailCommand::MarkMail(_)
            | MailCommand::SetCategory(_)
            | MailCommand::CreateFolder(_)
            | MailCommand::MoveMail(_) => Capability::Modify,
        }
    }
}

pub fn checked_limit_in(input: Option<u32>, default: u32, max: u32) -> Result<usize, String> {
    let n = input.unwrap_or(default);
    if !(1..=max).contains(&n) {
        return Err(format!("limit must be between 1 and {max}"));
    }
    Ok(n as usize)
}

pub fn checked_limit(input: Option<u32>) -> Result<usize, String> {
    let n = input.unwrap_or(10);
    if !(1..=30).contains(&n) {
        return Err("limit must be between 1 and 30".into());
    }
    Ok(n as usize)
}

fn validate_backend(name: Option<&str>, allow_indexed: bool) -> Result<(), String> {
    if matches!(name.unwrap_or("auto"), "auto" | "table" | "legacy")
        || (allow_indexed && name == Some("indexed"))
    {
        Ok(())
    } else {
        Err("backend must be auto, table, legacy or indexed (search only)".into())
    }
}

pub fn validate_command(command: &MailCommand) -> Result<(), String> {
    match command {
        MailCommand::WhoAmI
        | MailCommand::ListCategories
        | MailCommand::ListRules
        | MailCommand::GetOutOfOffice => Ok(()),
        MailCommand::ActionPrepare(_) | MailCommand::ActionExecute { .. } => Ok(()),
        MailCommand::ActionLocate { action_id, .. } => checked_send_id(action_id),
        MailCommand::ComposePrepare(_)
        | MailCommand::ComposeInspect(_)
        | MailCommand::ComposeSend { .. } => Ok(()),
        MailCommand::ComposeLocate(id) => checked_send_id(id),
        MailCommand::CreateTask(x) => validate_create_task(x),
        MailCommand::CompleteTask(x) => checked_id(&x.entry_id, x.store_id.as_deref()),
        MailCommand::UpdateTask(x) => validate_update_task(x),
        MailCommand::GetTask(x) => {
            checked_id(&x.entry_id, x.store_id.as_deref())?;
            if x.max_body_chars.unwrap_or(4000) > 50_000 {
                return Err("max_body_chars must be <= 50000".into());
            }
            Ok(())
        }
        MailCommand::CreateEvent(x) => validate_create_event(x),
        MailCommand::UpdateEvent(x) => validate_update_event(x),
        MailCommand::MarkMail(x) => {
            checked_id(&x.entry_id, x.store_id.as_deref())?;
            if x.read.is_none() && x.flagged.is_none() {
                return Err("set read and/or flagged; nothing to change".into());
            }
            Ok(())
        }
        MailCommand::SetCategory(x) => {
            checked_id(&x.entry_id, x.store_id.as_deref())?;
            if x.categories.chars().count() > 500 || x.categories.chars().any(char::is_control) {
                return Err(
                    "categories must be at most 500 characters without control characters".into(),
                );
            }
            Ok(())
        }
        MailCommand::CreateFolder(x) => {
            let name = x.name.trim();
            if name.is_empty()
                || x.name.chars().count() > 120
                || name == "."
                || name == ".."
                || x.name
                    .chars()
                    .any(|c| c.is_control() || c == '/' || c == '\\')
            {
                return Err("folder name must be 1..120 characters without slashes".into());
            }
            if x.parent.as_deref().is_some_and(|p| p.chars().count() > 300) {
                return Err("parent path is too long".into());
            }
            Ok(())
        }
        MailCommand::MoveMail(x) => {
            checked_id(&x.entry_id, x.store_id.as_deref())?;
            checked_text(&x.target_folder, "target_folder")?;
            let key = x.target_folder.trim().to_ascii_lowercase();
            if key == "deleted" || key == "deleted items" {
                return Err(
                    "moving mail into Deleted Items is a delete and is not allowed here".into(),
                );
            }
            Ok(())
        }
        MailCommand::ListEvents(x) => {
            checked_limit_in(x.limit, 50, 200)?;
            validate_calendar_spec(x.calendar.as_deref())?;
            let start = x
                .start
                .as_deref()
                .map(LocalTime::parse_with_offset)
                .transpose()?;
            let end = x
                .end
                .as_deref()
                .map(LocalTime::parse_with_offset)
                .transpose()?;
            // Mixed local/offset bounds are compared after conversion at run time.
            if let (Some((start, so)), Some((end, eo))) = (start, end)
                && so.is_some() == eo.is_some()
            {
                let minutes = (end.minutes() - i64::from(eo.unwrap_or(0)))
                    - (start.minutes() - i64::from(so.unwrap_or(0)));
                if minutes <= 0 {
                    return Err("end must be after start".into());
                }
            }
            Ok(())
        }
        MailCommand::GetEvent(x) => {
            if x.entry_id.is_empty()
                || x.entry_id.len() > 4096
                || x.store_id
                    .as_ref()
                    .is_some_and(|s| s.is_empty() || s.len() > 4096)
            {
                return Err("invalid event EntryID/StoreID".into());
            }
            if x.max_body_chars.unwrap_or(4000) > 50_000 {
                return Err("max_body_chars must be <= 50000".into());
            }
            Ok(())
        }
        MailCommand::ListTasks(x) => checked_limit_in(x.limit, 50, 200).map(|_| ()),
        MailCommand::ListContacts(x) => {
            checked_limit_in(x.limit, 50, 200)?;
            if x.offset.unwrap_or(0) > 5000 {
                return Err("offset must be 0..5000".into());
            }
            Ok(())
        }
        MailCommand::SearchContacts(x) => {
            checked_limit_in(x.limit, 25, 100)?;
            checked_text(&x.query, "query")
        }
        MailCommand::GetContact(x) => {
            if x.entry_id.is_empty()
                || x.entry_id.len() > 4096
                || x.store_id
                    .as_ref()
                    .is_some_and(|s| s.is_empty() || s.len() > 4096)
            {
                return Err("invalid contact EntryID/StoreID".into());
            }
            if x.max_notes_chars.unwrap_or(4000) > 50_000 {
                return Err("max_notes_chars must be <= 50000".into());
            }
            Ok(())
        }
        MailCommand::ResolveName(x) => checked_text(&x.name, "name"),
        MailCommand::ListFolders(x) => {
            if x.max_depth.unwrap_or(2) > 4 {
                Err("max_depth must be between 0 and 4".into())
            } else {
                Ok(())
            }
        }
        MailCommand::ListMails(x) => {
            checked_limit(x.limit)?;
            validate_backend(x.backend.as_deref(), false)?;
            if x.offset.unwrap_or(0) > 1000 {
                return Err("offset must be 0..1000".into());
            }
            received_filter(x.since.as_deref(), x.until.as_deref())?;
            Ok(())
        }
        MailCommand::SearchMails(x) => {
            checked_limit(x.limit)?;
            validate_backend(x.backend.as_deref(), true)?;
            if x.search_subfolders.is_some() && x.backend.as_deref() != Some("indexed") {
                return Err("search_subfolders requires backend=indexed".into());
            }
            if x.query.trim().is_empty() || x.query.len() > 160 {
                return Err("query must contain 1..160 characters".into());
            }
            if !matches!(
                x.scope.as_deref().unwrap_or("subject_body"),
                "subject" | "from" | "subject_body"
            ) {
                return Err("scope must be subject, from, or subject_body".into());
            }
            Ok(())
        }
        MailCommand::GetThread(x) => {
            if x.entry_id.is_empty()
                || x.entry_id.len() > 4096
                || x.store_id
                    .as_ref()
                    .is_some_and(|s| s.is_empty() || s.len() > 4096)
            {
                return Err("invalid thread EntryID/StoreID".into());
            }
            if !(1..=30).contains(&x.max_messages.unwrap_or(10)) {
                return Err("max_messages must be 1..30".into());
            }
            if !(1..=6000).contains(&x.per_message_chars.unwrap_or(1500)) {
                return Err("per_message_chars must be 1..6000".into());
            }
            if !(1..=24000).contains(&x.total_body_chars.unwrap_or(8000)) {
                return Err("total_body_chars must be 1..24000".into());
            }
            Ok(())
        }
        MailCommand::GetMail(x) => {
            if x.entry_id.is_empty() || x.entry_id.len() > 4096 {
                return Err("invalid EntryID length".into());
            }
            if x.max_body_chars.unwrap_or(4000) > 12000 {
                return Err("max_body_chars must be <= 12000".into());
            }
            Ok(())
        }
    }
}

/// Restrict/GetTable filter on `[ReceivedTime]` for local wall-clock bounds
/// (inclusive, like Python). Empty when neither bound is set.
pub fn received_filter(since: Option<&str>, until: Option<&str>) -> Result<String, String> {
    let parse =
        |value: &str, name: &str| LocalTime::parse(value).map_err(|e| format!("{name}: {e}"));
    let since = since.map(|v| parse(v, "since")).transpose()?;
    let until = until.map(|v| parse(v, "until")).transpose()?;
    if let (Some(s), Some(u)) = (&since, &until)
        && s.minutes() > u.minutes()
    {
        return Err("since must not be after until".into());
    }
    let clauses: Vec<String> = since
        .iter()
        .map(|t| format!("[ReceivedTime] >= '{}'", t.jet()))
        .chain(
            until
                .iter()
                .map(|t| format!("[ReceivedTime] <= '{}'", t.jet())),
        )
        .collect();
    Ok(clauses.join(" AND "))
}

/// Split a search query into (anchor, remaining words), lowercased, like Python.
/// The anchor is the longest word (first on ties): the most selective term to
/// push into the Outlook filter. The remaining words are then required per
/// item, because ANDing two LIKEs on one property returns nothing in Outlook.
pub fn split_search_words(query: &str) -> (String, Vec<String>) {
    let mut words: Vec<String> = query.split_whitespace().map(str::to_lowercase).collect();
    if words.is_empty() {
        return (query.to_lowercase(), words);
    }
    let longest = words.iter().map(|w| w.chars().count()).max().unwrap_or(0);
    let at = words
        .iter()
        .position(|w| w.chars().count() == longest)
        .unwrap_or(0);
    let anchor = words.remove(at);
    (anchor, words)
}

/// True when every word occurs in the (already lowercased) haystack.
pub fn has_all_words(haystack: &str, words: &[String]) -> bool {
    words.iter().all(|w| haystack.contains(w.as_str()))
}

pub fn escape_dasl(raw: &str) -> String {
    // User text is always a single-quoted DASL literal; double apostrophes
    // to prevent injection. Double quotes inside the value are harmless and
    // must not be changed (they delimit property identifiers, not values).
    raw.replace('\'', "''")
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn get_mail_accepts_include_html_within_the_same_body_ceiling() {
        let params = |chars| GetMailParams {
            entry_id: "m1".into(),
            store_id: None,
            include_body: None,
            include_html: Some(true),
            max_body_chars: Some(chars),
            body_offset: None,
        };
        assert!(validate_command(&MailCommand::GetMail(params(12000))).is_ok());
        assert!(validate_command(&MailCommand::GetMail(params(12001))).is_err());
    }

    #[test]
    fn received_filter_builds_inclusive_local_bounds() {
        assert_eq!(received_filter(None, None).unwrap(), "");
        assert_eq!(
            received_filter(Some("2026-11-02T00:00:00"), None).unwrap(),
            "[ReceivedTime] >= '11/02/2026 12:00 AM'"
        );
        assert_eq!(
            received_filter(Some("2026-11-02"), Some("2026-11-03T14:30")).unwrap(),
            "[ReceivedTime] >= '11/02/2026 12:00 AM' AND [ReceivedTime] <= '11/03/2026 02:30 PM'"
        );
    }
    #[test]
    fn search_words_pick_the_longest_as_anchor() {
        let (anchor, rest) = split_search_words("Budget  the FORECAST");
        assert_eq!(
            (anchor.as_str(), rest),
            ("forecast", vec!["budget".to_string(), "the".to_string()])
        );
        // First of equal-length words wins, as in Python.
        let (anchor, rest) = split_search_words("alpha beta");
        assert_eq!((anchor.as_str(), rest), ("alpha", vec!["beta".to_string()]));
        // One word has nothing left to verify.
        assert_eq!(
            split_search_words("invoice"),
            ("invoice".to_string(), vec![])
        );
    }
    #[test]
    fn all_words_must_appear_anywhere() {
        let words = vec!["budget".to_string(), "q3".to_string()];
        assert!(has_all_words("q3 forecast for the budget", &words));
        assert!(!has_all_words("budget forecast", &words));
        assert!(has_all_words("anything", &[]));
    }
    #[test]
    fn received_filter_rejects_bad_or_reversed_bounds() {
        assert!(received_filter(Some("yesterday"), None).is_err());
        assert!(received_filter(Some("2026-11-02T00:00:00Z"), None).is_err());
        assert!(received_filter(Some("2026-11-03"), Some("2026-11-02")).is_err());
    }
    #[test]
    fn enforces_bounds() {
        assert_eq!(checked_limit(None).unwrap(), 10);
        assert!(checked_limit(Some(0)).is_err());
        assert!(checked_limit(Some(31)).is_err());
    }
    #[test]
    fn validates_thread_limits_and_handle() {
        let params = ThreadParams {
            entry_id: "m7".into(),
            store_id: None,
            max_messages: Some(30),
            include_bodies: Some(true),
            trim_quoted: Some(false),
            per_message_chars: Some(6000),
            total_body_chars: Some(24000),
        };
        assert!(validate_command(&MailCommand::GetThread(params.clone())).is_ok());
        let mut too_many = params.clone();
        too_many.max_messages = Some(31);
        assert!(validate_command(&MailCommand::GetThread(too_many)).is_err());
        let mut too_large = params.clone();
        too_large.total_body_chars = Some(24001);
        assert!(validate_command(&MailCommand::GetThread(too_large)).is_err());
        let mut empty = params;
        empty.entry_id.clear();
        assert!(validate_command(&MailCommand::GetThread(empty)).is_err());
    }

    #[test]
    fn modifying_commands_need_the_modify_capability_and_valid_input() {
        let mark = |read, flagged| {
            MailCommand::MarkMail(MarkMailParams {
                entry_id: "m1".into(),
                store_id: None,
                read,
                flagged,
            })
        };
        assert_eq!(mark(Some(true), None).capability(), Capability::Modify);
        assert!(validate_command(&mark(Some(true), None)).is_ok());
        assert!(validate_command(&mark(None, None)).is_err());
        let category = |c: &str| {
            MailCommand::SetCategory(SetCategoryParams {
                entry_id: "m1".into(),
                store_id: None,
                categories: c.into(),
            })
        };
        assert_eq!(category("").capability(), Capability::Modify);
        assert!(validate_command(&category("Red Category; Blue")).is_ok());
        assert!(validate_command(&category("")).is_ok());
        assert!(validate_command(&category("bad\u{7}")).is_err());
        let folder = |n: &str| {
            MailCommand::CreateFolder(CreateFolderParams {
                name: n.into(),
                parent: None,
            })
        };
        assert_eq!(folder("x").capability(), Capability::Modify);
        for bad in ["", " ", "..", "a/b", "a\\b", "bad\nname"] {
            assert!(validate_command(&folder(bad)).is_err(), "{bad:?}");
        }
        assert!(validate_command(&folder("Projects 2026")).is_ok());
        let moving = |t: &str| {
            MailCommand::MoveMail(MoveMailParams {
                entry_id: "m1".into(),
                store_id: None,
                target_folder: t.into(),
            })
        };
        assert_eq!(moving("Archive").capability(), Capability::Modify);
        assert!(validate_command(&moving("Inbox/Projects")).is_ok());
        for blocked in ["Deleted Items", " deleted ", "DELETED ITEMS"] {
            assert!(validate_command(&moving(blocked)).is_err(), "{blocked}");
        }
    }

    #[test]
    fn task_and_event_writes_use_modify_and_events_capabilities() {
        let task = CreateTaskParams {
            subject: "Pay invoice".into(),
            due_date: Some("2026-05-01".into()),
            body: None,
            importance: Some("High".into()),
            reminder: None,
        };
        assert!(validate_create_task(&task).is_ok());
        assert_eq!(
            MailCommand::CreateTask(task).capability(),
            Capability::Modify
        );
        assert_eq!(
            MailCommand::CompleteTask(CompleteTaskParams {
                entry_id: "m1".into(),
                store_id: None
            })
            .capability(),
            Capability::Modify
        );
        let event = || CreateEventParams {
            subject: "Review".into(),
            start: "2026-05-01T09:00".into(),
            end: "2026-05-01T10:00".into(),
            location: None,
            body: None,
            attendees: None,
            is_online_meeting: None,
            reminder_minutes: None,
            no_reminder: None,
            calendar: None,
            recurrence: None,
        };
        assert!(validate_create_event(&event()).is_ok());
        assert_eq!(
            MailCommand::CreateEvent(event()).capability(),
            Capability::Events
        );
        assert_eq!(
            MailCommand::UpdateEvent(UpdateEventParams {
                entry_id: "m1".into(),
                store_id: None,
                subject: Some("x".into()),
                start: None,
                end: None,
                location: None,
                body: None
            })
            .capability(),
            Capability::Events
        );
    }

    #[test]
    fn event_inputs_reject_notifying_and_malformed_requests() {
        let base = || CreateEventParams {
            subject: "Review".into(),
            start: "2026-05-01T09:00".into(),
            end: "2026-05-01T10:00".into(),
            location: None,
            body: None,
            attendees: None,
            is_online_meeting: None,
            reminder_minutes: None,
            no_reminder: None,
            calendar: None,
            recurrence: None,
        };
        let mut x = base();
        x.no_reminder = Some(true);
        assert!(validate_create_event(&x).is_ok());
        x.reminder_minutes = Some(10);
        assert!(validate_create_event(&x).is_err());
        let mut x = base();
        x.attendees = Some(vec!["a@b.example".into()]);
        assert!(
            validate_create_event(&x).is_err(),
            "attendees would send invitations"
        );
        let mut x = base();
        x.attendees = Some(vec![]);
        assert!(validate_create_event(&x).is_ok());
        let mut x = base();
        x.is_online_meeting = Some(true);
        assert!(validate_create_event(&x).is_err());
        let mut x = base();
        x.end = "2026-05-01T09:00".into();
        assert!(validate_create_event(&x).is_err());
        let mut x = base();
        x.start = "2026-05-01T09:00+02:00".into();
        x.end = "2026-05-01T07:30Z".into();
        assert!(
            validate_create_event(&x).is_ok(),
            "offsets compare in UTC: 07:00Z to 07:30Z"
        );
        let mut x = base();
        x.start = "not a date".into();
        assert!(validate_create_event(&x).is_err());
        let recurrence = |kind: &str, occ: Option<u32>, end: Option<&str>| RecurrenceParams {
            kind: kind.into(),
            interval: None,
            occurrences: occ,
            end_date: end.map(Into::into),
        };
        let mut x = base();
        x.recurrence = Some(recurrence("weekly", Some(4), None));
        assert!(validate_create_event(&x).is_ok());
        for bad in [
            recurrence("hourly", None, None),
            recurrence("daily", Some(0), None),
            recurrence("daily", Some(2), Some("2026-06-01")),
            recurrence("daily", None, Some("2026-04-01")),
        ] {
            let mut x = base();
            x.recurrence = Some(bad);
            assert!(validate_create_event(&x).is_err());
        }
        let update = |subject: Option<&str>| UpdateEventParams {
            entry_id: "m1".into(),
            store_id: None,
            subject: subject.map(Into::into),
            start: None,
            end: None,
            location: None,
            body: None,
        };
        assert!(
            validate_update_event(&update(None)).is_err(),
            "nothing to change"
        );
        assert!(validate_update_event(&update(Some("New title"))).is_ok());
        let task = |s: &str, imp: Option<&str>| CreateTaskParams {
            subject: s.into(),
            due_date: None,
            body: None,
            importance: imp.map(Into::into),
            reminder: None,
        };
        assert!(validate_create_task(&task("", None)).is_err());
        assert!(validate_create_task(&task("ok", Some("urgent"))).is_err());
    }

    #[test]
    fn delete_steps_need_the_delete_capability_and_locate_is_a_read() {
        let target = prepare_delete(
            &DeleteItemParams {
                entry_id: "m1".into(),
                store_id: None,
            },
            ActionKind::DeleteMail,
        )
        .unwrap();
        assert_eq!(
            MailCommand::ActionPrepare(target.clone()).capability(),
            Capability::Delete
        );
        let preview = ActionPreview {
            action_id: "a-1".into(),
            summary: String::new(),
            facts: vec![],
            warnings: vec![],
            hash: String::new(),
        };
        assert_eq!(
            MailCommand::ActionExecute {
                target: target.clone(),
                preview
            }
            .capability(),
            Capability::Delete
        );
        let locate = MailCommand::ActionLocate {
            action_id: "a-1".into(),
            target: Some(target),
        };
        assert_eq!(locate.capability(), Capability::Read);
        assert!(validate_command(&locate).is_ok());
        assert!(
            validate_command(&MailCommand::ActionLocate {
                action_id: "bad id!".into(),
                target: None
            })
            .is_err()
        );
        assert!(
            prepare_delete(
                &DeleteItemParams {
                    entry_id: "".into(),
                    store_id: None
                },
                ActionKind::DeleteEvent
            )
            .is_err()
        );
    }

    #[test]
    fn meeting_responses_validate_and_pick_capability_from_whether_a_reply_is_sent() {
        let params = |response: &str, send: Option<bool>| RespondEventParams {
            entry_id: "m1".into(),
            store_id: None,
            response: response.into(),
            send_response: send,
        };
        let sending = prepare_respond_event(&params("Accept", None)).unwrap();
        assert_eq!(sending.response.as_deref(), Some("accept"));
        assert_eq!(sending.enabled, Some(true));
        assert_eq!(
            MailCommand::ActionPrepare(sending).capability(),
            Capability::Send
        );
        let local = prepare_respond_event(&params("decline", Some(false))).unwrap();
        assert_eq!(
            MailCommand::ActionPrepare(local).capability(),
            Capability::Events
        );
        assert!(prepare_respond_event(&params("maybe", None)).is_err());
        assert!(
            prepare_respond_event(&RespondEventParams {
                entry_id: "".into(),
                store_id: None,
                response: "accept".into(),
                send_response: None
            })
            .is_err()
        );
    }

    #[test]
    fn rule_changes_need_the_rules_capability_and_a_clean_name() {
        let target = prepare_toggle_rule(&ToggleRuleParams {
            rule_name: " Move newsletters ".into(),
            enabled: true,
        })
        .unwrap();
        assert_eq!(target.entry_id, "Move newsletters");
        assert_eq!(target.enabled, Some(true));
        assert_eq!(
            MailCommand::ActionPrepare(target).capability(),
            Capability::Rules
        );
        for bad in ["", "   ", "bad\nname"] {
            assert!(
                prepare_toggle_rule(&ToggleRuleParams {
                    rule_name: bad.into(),
                    enabled: false
                })
                .is_err(),
                "{bad:?}"
            );
        }
        assert!(
            prepare_toggle_rule(&ToggleRuleParams {
                rule_name: "x".repeat(256),
                enabled: false
            })
            .is_err()
        );
        // A rule named like a session handle is still a rule name.
        assert_eq!(
            prepare_toggle_rule(&ToggleRuleParams {
                rule_name: "m1".into(),
                enabled: true
            })
            .unwrap()
            .entry_id,
            "m1"
        );
    }

    #[test]
    fn send_pipeline_commands_map_to_draft_and_send_capabilities() {
        let draft = DraftRef {
            entry_id: "m1".into(),
            store_id: None,
        };
        assert_eq!(
            MailCommand::ComposeInspect(draft.clone()).capability(),
            Capability::Read
        );
        assert_eq!(
            MailCommand::ComposeLocate("a-1".into()).capability(),
            Capability::Read
        );
        let spec = prepare_send_mail(&SendMailParams {
            to: vec!["a@b.example".into()],
            subject: "s".into(),
            body: "b".into(),
            cc: None,
            bcc: None,
            html: None,
            attachments: None,
            importance: None,
            save_only: None,
        })
        .unwrap();
        assert_eq!(
            MailCommand::ComposePrepare(spec).capability(),
            Capability::Draft
        );
        let preview = Preview {
            send_id: "x".into(),
            sender: String::new(),
            recipients: vec![],
            subject: String::new(),
            body_excerpt: String::new(),
            body_chars: 0,
            html: false,
            attachments: vec![],
            hash: String::new(),
            kind: "new".into(),
        };
        assert_eq!(
            MailCommand::ComposeSend { draft, preview }.capability(),
            Capability::Send
        );
    }

    #[test]
    fn html_flag_is_carried_on_send_reply_and_forward() {
        let send = |html| SendMailParams {
            to: vec!["a@b.example".into()],
            subject: "Hi".into(),
            body: "<b>x</b>".into(),
            cc: None,
            bcc: None,
            html,
            attachments: None,
            importance: None,
            save_only: None,
        };
        let flag = |r: PrepareRequest| match r {
            PrepareRequest::New { html, .. }
            | PrepareRequest::Reply { html, .. }
            | PrepareRequest::Forward { html, .. } => html,
        };
        assert!(flag(prepare_send_mail(&send(Some(true))).unwrap()));
        assert!(!flag(prepare_send_mail(&send(None)).unwrap()));
        let reply = |html| ReplyMailParams {
            entry_id: "abc".into(),
            store_id: None,
            body: "x".into(),
            reply_all: None,
            html,
            attachments: None,
        };
        assert!(flag(prepare_reply_mail(&reply(Some(true))).unwrap()));
        assert!(!flag(prepare_reply_mail(&reply(Some(false))).unwrap()));
        let forward = |html| ForwardMailParams {
            entry_id: "abc".into(),
            store_id: None,
            to: vec!["a@b.example".into()],
            cc: None,
            body: None,
            html,
        };
        assert!(flag(prepare_forward_mail(&forward(Some(true))).unwrap()));
        assert!(!flag(prepare_forward_mail(&forward(None)).unwrap()));
    }

    #[test]
    fn attachment_paths_are_validated_on_send_and_reply() {
        let send = |attachments| SendMailParams {
            to: vec!["a@b.example".into()],
            subject: "Hi".into(),
            body: "x".into(),
            cc: None,
            bcc: None,
            html: None,
            attachments: Some(attachments),
            importance: None,
            save_only: None,
        };
        let reply = |attachments| ReplyMailParams {
            entry_id: "abc".into(),
            store_id: None,
            body: "x".into(),
            reply_all: None,
            html: None,
            attachments: Some(attachments),
        };
        // Empty list is fine; a missing file, traversal and too many are not.
        assert!(prepare_send_mail(&send(vec![])).is_ok());
        assert!(prepare_reply_mail(&reply(vec![])).is_ok());
        for bad in [
            vec!["C:/definitely/not/here/x.txt".to_string()],
            vec!["C:/Users/../Windows/win.ini".to_string()],
            vec![String::new()],
            vec!["x.txt".to_string(); crate::safe_path::MAX_ATTACH_FILES + 1],
        ] {
            assert!(prepare_send_mail(&send(bad.clone())).is_err(), "{bad:?}");
            assert!(prepare_reply_mail(&reply(bad.clone())).is_err(), "{bad:?}");
        }
        // A real file inside the profile passes; an executable beside it does not.
        let Some(profile) = std::env::var_os("USERPROFILE").map(std::path::PathBuf::from) else {
            return;
        };
        let dir = std::env::temp_dir().join(format!("outlook-mcp-contract-{}", std::process::id()));
        if !dir.starts_with(&profile) || std::fs::create_dir_all(&dir).is_err() {
            return;
        }
        let (ok, exe) = (dir.join("a.txt"), dir.join("run.exe"));
        std::fs::write(&ok, b"hi").unwrap();
        std::fs::write(&exe, b"MZ").unwrap();
        let text = |p: &std::path::Path| p.to_string_lossy().into_owned();
        assert!(prepare_send_mail(&send(vec![text(&ok)])).is_ok());
        let paths = |r: PrepareRequest| match r {
            PrepareRequest::Reply { attachments, .. } => attachments,
            _ => vec![],
        };
        assert_eq!(
            paths(prepare_reply_mail(&reply(vec![text(&ok)])).unwrap()).len(),
            1
        );
        let error = prepare_send_mail(&send(vec![text(&exe)])).unwrap_err();
        assert!(error.starts_with("run.exe was not attached"), "{error}");
        let _ = std::fs::remove_dir_all(dir);
    }

    #[test]
    fn send_inputs_reject_unsupported_and_malformed_values() {
        let base = || SendMailParams {
            to: vec!["a@b.example".into()],
            subject: "Hi".into(),
            body: "text".into(),
            cc: None,
            bcc: None,
            html: None,
            attachments: None,
            importance: None,
            save_only: None,
        };
        assert!(prepare_send_mail(&base()).is_ok());
        let mut x = base();
        x.attachments = Some(vec!["C:/definitely/not/here/x.txt".into()]);
        assert!(prepare_send_mail(&x).is_err());
        let mut x = base();
        x.to = vec![];
        assert!(prepare_send_mail(&x).is_err());
        let mut x = base();
        x.to = vec!["a@b.example; evil@x.example".into()];
        assert!(prepare_send_mail(&x).is_err(), "separator smuggling");
        let mut x = base();
        x.subject = "line\nbreak".into();
        assert!(prepare_send_mail(&x).is_err());
        let mut x = base();
        x.importance = Some("urgent".into());
        assert!(prepare_send_mail(&x).is_err());
        assert!(checked_send_id("abc-123").is_ok());
        for bad in ["", "a b", "x'; DROP", &"x".repeat(81)] {
            assert!(checked_send_id(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn every_current_command_is_a_read() {
        let commands = [
            MailCommand::WhoAmI,
            MailCommand::ListCategories,
            MailCommand::ListRules,
            MailCommand::GetOutOfOffice,
        ];
        assert!(commands.iter().all(|c| c.capability() == Capability::Read));
    }

    #[test]
    fn validates_directory_commands() {
        let search = |q: &str, limit| {
            validate_command(&MailCommand::SearchContacts(SearchContactsParams {
                query: q.into(),
                limit,
                include_directory: None,
            }))
        };
        assert!(search("jane doe", None).is_ok());
        assert!(search("   ", None).is_err());
        assert!(search(&"x".repeat(161), None).is_err());
        assert!(search("jane", Some(100)).is_ok());
        assert!(search("jane", Some(101)).is_err());
        let list = |offset| {
            validate_command(&MailCommand::ListContacts(ListContactsParams {
                limit: None,
                offset,
            }))
        };
        assert!(
            validate_command(&MailCommand::ListContacts(ListContactsParams {
                limit: Some(200),
                offset: None
            }))
            .is_ok()
        );
        assert!(
            validate_command(&MailCommand::ListContacts(ListContactsParams {
                limit: Some(201),
                offset: None
            }))
            .is_err()
        );
        assert!(list(Some(5000)).is_ok());
        assert!(list(Some(5001)).is_err());
        assert!(
            validate_command(&MailCommand::ResolveName(ResolveNameParams {
                name: "".into()
            }))
            .is_err()
        );
        assert!(
            validate_command(&MailCommand::GetContact(GetContactParams {
                entry_id: "".into(),
                store_id: None,
                max_notes_chars: None
            }))
            .is_err()
        );
    }

    #[test]
    fn calendar_spec_is_validated_and_keeps_events_capability() {
        assert!(validate_calendar_spec(None).is_ok());
        assert!(validate_calendar_spec(Some("Archive/Calendar")).is_ok());
        for bad in ["", "   ", "a\nb"] {
            assert!(validate_calendar_spec(Some(bad)).is_err(), "{bad:?}");
        }
        assert!(validate_calendar_spec(Some(&"x".repeat(301))).is_err());
        let mut event = CreateEventParams {
            subject: "Review".into(),
            start: "2026-05-01T09:00".into(),
            end: "2026-05-01T10:00".into(),
            location: None,
            body: None,
            attendees: None,
            is_online_meeting: None,
            reminder_minutes: None,
            no_reminder: None,
            calendar: Some("Shared/Calendar".into()),
            recurrence: None,
        };
        assert!(validate_command(&MailCommand::CreateEvent(event.clone())).is_ok());
        assert_eq!(
            MailCommand::CreateEvent(event.clone()).capability(),
            Capability::Events
        );
        event.calendar = Some(" ".into());
        assert!(validate_command(&MailCommand::CreateEvent(event.clone())).is_err());
        // Attendee meetings stay on the default calendar.
        event.calendar = Some("Shared/Calendar".into());
        event.attendees = Some(vec!["a@b.example".into()]);
        assert!(
            prepare_invite_meeting(&event)
                .unwrap_err()
                .contains("default calendar")
        );
        event.calendar = None;
        assert!(prepare_invite_meeting(&event).is_ok());
        let list = |calendar: Option<&str>| {
            validate_command(&MailCommand::ListEvents(ListEventsParams {
                start: None,
                end: None,
                limit: None,
                include_recurrences: None,
                calendar: calendar.map(Into::into),
            }))
        };
        assert!(list(Some("Archive/Calendar")).is_ok());
        assert!(list(Some("")).is_err());
    }

    #[test]
    fn task_get_and_update_validate_and_map_capabilities() {
        let get = |max| GetTaskParams {
            entry_id: "m1".into(),
            store_id: None,
            max_body_chars: max,
        };
        assert!(validate_command(&MailCommand::GetTask(get(Some(50_000)))).is_ok());
        assert!(validate_command(&MailCommand::GetTask(get(Some(50_001)))).is_err());
        assert_eq!(
            MailCommand::GetTask(get(None)).capability(),
            Capability::Read
        );
        let update = || UpdateTaskParams {
            entry_id: "m1".into(),
            store_id: None,
            subject: None,
            due_date: None,
            body: None,
            importance: None,
            percent_complete: None,
            status: None,
            reminder: None,
        };
        assert!(
            validate_update_task(&update()).is_err(),
            "nothing to change"
        );
        let mut x = update();
        x.subject = Some("Pay invoice".into());
        x.due_date = Some("2026-05-01".into());
        x.importance = Some("High".into());
        x.status = Some("In_Progress".into());
        assert!(validate_update_task(&x).is_ok());
        assert_eq!(
            MailCommand::UpdateTask(x.clone()).capability(),
            Capability::Modify
        );
        x.percent_complete = Some(50);
        assert!(
            validate_update_task(&x).is_err(),
            "status and percent together"
        );
        let bad: [fn(&mut UpdateTaskParams); 6] = [
            |x| x.subject = Some(" ".into()),
            |x| x.due_date = Some("tomorrow".into()),
            |x| x.importance = Some("urgent".into()),
            |x| x.percent_complete = Some(101),
            |x| x.status = Some("done".into()),
            |x| x.entry_id = String::new(),
        ];
        for change in bad {
            let mut y = update();
            y.body = Some("notes".into());
            change(&mut y);
            assert!(validate_update_task(&y).is_err());
        }
    }

    #[test]
    fn validates_event_and_task_commands() {
        let events = |start: Option<&str>, end: Option<&str>, limit| {
            validate_command(&MailCommand::ListEvents(ListEventsParams {
                start: start.map(Into::into),
                end: end.map(Into::into),
                limit,
                include_recurrences: None,
                calendar: None,
            }))
        };
        assert!(events(None, None, None).is_ok());
        assert!(events(Some("2026-04-25T09:00"), Some("2026-05-09"), Some(200)).is_ok());
        assert!(events(None, None, Some(201)).is_err());
        assert!(events(Some("2026-04-25T09:00:00Z"), None, None).is_ok());
        assert!(
            events(
                Some("2026-04-25T09:00:00+02:00"),
                Some("2026-04-25T08:00:00Z"),
                None
            )
            .is_ok()
        );
        assert!(
            events(
                Some("2026-04-25T09:00:00+02:00"),
                Some("2026-04-25T07:00:00Z"),
                None
            )
            .is_err()
        );
        assert!(events(Some("2026-04-25T09:00:00+99:00"), None, None).is_err());
        assert!(events(Some("2026-05-09"), Some("2026-04-25"), None).is_err());
        assert!(events(Some("2026-01-01"), Some("2026-01-01"), None).is_err());
        assert!(events(Some("2026-01-01"), Some("2031-01-01"), None).is_ok());
        let tasks = |limit| {
            validate_command(&MailCommand::ListTasks(ListTasksParams {
                limit,
                include_completed: None,
            }))
        };
        assert!(tasks(Some(200)).is_ok());
        assert!(tasks(Some(0)).is_err());
        assert!(
            validate_command(&MailCommand::GetEvent(GetEventParams {
                entry_id: "".into(),
                store_id: None,
                max_body_chars: None
            }))
            .is_err()
        );
        let long_body = |n| {
            validate_command(&MailCommand::GetEvent(GetEventParams {
                entry_id: "m1".into(),
                store_id: None,
                max_body_chars: Some(n),
            }))
        };
        assert!(long_body(50_000).is_ok());
        assert!(long_body(50_001).is_err());
    }

    #[test]
    fn invitations_need_plain_addresses_and_no_recurrence() {
        let base = || CreateEventParams {
            subject: "Sync".into(),
            start: "2030-02-01T10:00:00".into(),
            end: "2030-02-01T10:30:00".into(),
            location: None,
            body: None,
            attendees: Some(vec!["A@B.example".into(), "a@b.example".into()]),
            is_online_meeting: None,
            reminder_minutes: None,
            no_reminder: None,
            calendar: None,
            recurrence: None,
        };
        let target = prepare_invite_meeting(&base()).unwrap();
        assert_eq!(target.kind, ActionKind::InviteMeeting);
        assert_eq!(
            target.payload.unwrap()["attendees"],
            serde_json::json!(["a@b.example"])
        );
        for bad in [
            "Bob",
            "a@b@c",
            "a b@c.example",
            "x@y.example;z@y.example",
            "@x",
        ] {
            let mut x = base();
            x.attendees = Some(vec![bad.into()]);
            assert!(prepare_invite_meeting(&x).is_err(), "{bad}");
        }
        let mut x = base();
        x.attendees = Some(vec![]);
        assert!(prepare_invite_meeting(&x).is_err());
        let mut x = base();
        x.is_online_meeting = Some(true);
        assert!(prepare_invite_meeting(&x).is_err());
    }

    #[test]
    fn recurring_meetings_are_accepted_described_and_hashed_into_the_payload() {
        let base = |kind: &str, occ: Option<u32>| CreateEventParams {
            subject: "Sync".into(),
            start: "2030-02-04T10:00:00".into(),
            end: "2030-02-04T10:30:00".into(),
            location: None,
            body: None,
            attendees: Some(vec!["a@b.example".into()]),
            is_online_meeting: None,
            reminder_minutes: None,
            no_reminder: None,
            calendar: None,
            recurrence: Some(RecurrenceParams {
                kind: kind.into(),
                interval: None,
                occurrences: occ,
                end_date: None,
            }),
        };
        let weekly = prepare_invite_meeting(&base("weekly", Some(10))).unwrap();
        let payload = weekly.payload.clone().unwrap();
        assert_eq!(payload["recurrence"]["type"], "weekly");
        assert_eq!(payload["recurrence"]["occurrences"], 10);
        let other = prepare_invite_meeting(&base("weekly", Some(11))).unwrap();
        assert_ne!(payload, other.payload.clone().unwrap());
        assert_ne!(weekly.entry_id, other.entry_id);
        let daily = prepare_invite_meeting(&base("daily", Some(10))).unwrap();
        assert_ne!(payload, daily.payload.unwrap());
        // Validation still applies to the recurrence of a meeting.
        assert!(prepare_invite_meeting(&base("hourly", None)).is_err());
        let mut x = base("daily", None);
        x.recurrence.as_mut().unwrap().interval = Some(0);
        assert!(prepare_invite_meeting(&x).is_err());
    }

    #[test]
    fn recurrence_is_described_in_plain_words() {
        let start = LocalTime::parse("2030-02-04T10:00:00").unwrap(); // a Monday
        assert_eq!(weekday_of(&start), 1);
        assert_eq!(weekday_mask(&start), 2);
        let r = |kind: &str, interval, occurrences, end_date: Option<&str>| RecurrenceParams {
            kind: kind.into(),
            interval,
            occurrences,
            end_date: end_date.map(Into::into),
        };
        let text = describe_recurrence(&r("weekly", None, Some(10), None), &start);
        assert!(text.contains("every week on Monday"), "{text}");
        assert!(text.contains("10 times"), "{text}");
        let text = describe_recurrence(&r("weekly", Some(2), None, Some("2030-06-01")), &start);
        assert!(text.contains("every 2 weeks on Monday"), "{text}");
        assert!(text.contains("until 2030-06-01"), "{text}");
        let text = describe_recurrence(&r("monthly", None, None, None), &start);
        assert!(text.contains("every month on day 4"), "{text}");
        assert!(text.contains("no end date"), "{text}");
        let text = describe_recurrence(&r("yearly", None, Some(3), None), &start);
        assert!(text.contains("every year on 4 February"), "{text}");
    }

    #[test]
    fn contact_words_match_case_insensitively_and_all_required() {
        let words = query_words("  JANE   acme ");
        assert_eq!(words, ["jane", "acme"]);
        assert!(matches_all("Jane Doe jane@x.com ACME Corp", &words));
        assert!(!matches_all("Jane Doe Initech", &words));
    }

    #[test]
    fn escapes_dasl_literals() {
        assert_eq!(escape_dasl("O'Brien"), "O''Brien");
        assert_eq!(escape_dasl("\"hello\""), "\"hello\"");
    }
}
