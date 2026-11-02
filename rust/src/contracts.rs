//! Read-only MCP request contracts. Values are validated again on the server
//! side; JSON Schema is a usability hint, not an authorization boundary.
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
    /// ISO-8601 lower date bound, reserved for future compatibility.
    pub since: Option<String>,
    /// ISO-8601 upper date bound, reserved for future compatibility.
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
    /// HTML retrieval is deliberately disabled until explicit safety review.
    pub include_html: Option<bool>,
    /// Plain-text body budget (default 4000, maximum 12000 characters).
    pub max_body_chars: Option<u32>,
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
    /// Number of contacts to return (default 10, maximum 30).
    pub limit: Option<u32>,
    /// Number of contacts to skip across all contact folders (default 0, maximum 5000).
    pub offset: Option<u32>,
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct SearchContactsParams {
    /// Words that must all appear in name, email, company or job title (1..160 characters).
    pub query: String,
    /// Maximum result count (default 10, maximum 30).
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
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct ResolveNameParams {
    /// Display name or address to resolve against the address book (1..160 characters).
    pub name: String,
}

#[derive(Debug, Clone)]
pub enum MailCommand {
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

fn checked_text(value: &str, what: &str) -> Result<(), String> {
    if value.trim().is_empty() || value.chars().count() > 160 {
        return Err(format!("{what} must contain 1..160 characters"));
    }
    Ok(())
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
        MailCommand::ListContacts(x) => {
            checked_limit(x.limit)?;
            if x.offset.unwrap_or(0) > 5000 {
                return Err("offset must be 0..5000".into());
            }
            Ok(())
        }
        MailCommand::SearchContacts(x) => {
            checked_limit(x.limit)?;
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
            if x.since.is_some() || x.until.is_some() {
                return Err(
                    "since/until filters are not yet implemented; no date filter was applied"
                        .into(),
                );
            }
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
            if x.include_html.unwrap_or(false) {
                return Err("HTML bodies are disabled in this read-only prototype".into());
            }
            if x.max_body_chars.unwrap_or(4000) > 12000 {
                return Err("max_body_chars must be <= 12000".into());
            }
            Ok(())
        }
    }
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
    fn rejects_unimplemented_options_rather_than_ignoring_them() {
        let req = MailCommand::ListMails(ListMailsParams {
            folder: None,
            limit: None,
            offset: None,
            unread_only: None,
            from_address: None,
            since: Some("2026-01-01".into()),
            until: None,
            backend: None,
        });
        assert!(validate_command(&req).is_err());
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
        assert!(search("jane", Some(31)).is_err());
        let list = |offset| {
            validate_command(&MailCommand::ListContacts(ListContactsParams {
                limit: None,
                offset,
            }))
        };
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
                store_id: None
            }))
            .is_err()
        );
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
