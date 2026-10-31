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
}

#[derive(Debug, Clone, Deserialize, JsonSchema)]
pub struct GetMailParams {
    /// Outlook EntryID returned by a list/search result.
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

#[derive(Debug, Clone)]
pub enum MailCommand {
    WhoAmI,
    ListFolders(FolderParams),
    ListMails(ListMailsParams),
    SearchMails(SearchMailsParams),
    GetMail(GetMailParams),
}

pub fn checked_limit(input: Option<u32>) -> Result<usize, String> {
    let n = input.unwrap_or(10);
    if !(1..=30).contains(&n) {
        return Err("limit must be between 1 and 30".into());
    }
    Ok(n as usize)
}

pub fn validate_command(command: &MailCommand) -> Result<(), String> {
    match command {
        MailCommand::WhoAmI => Ok(()),
        MailCommand::ListFolders(x) => {
            if x.max_depth.unwrap_or(2) > 4 {
                Err("max_depth must be between 0 and 4".into())
            } else {
                Ok(())
            }
        }
        MailCommand::ListMails(x) => {
            checked_limit(x.limit)?;
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
    fn escapes_dasl_literals() {
        assert_eq!(escape_dasl("O'Brien"), "O''Brien");
        assert_eq!(escape_dasl("\"hello\""), "\"hello\"");
    }
}
