//! Server-enforced permission modes.
//!
//! The mode comes only from the server's own environment, never from a tool
//! argument or message text, so an LLM (or an email it reads) cannot raise it.
//! Parsing is fail-closed: an unrecognized value is an error, not a default
//! to something broader.

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Capability {
    /// Reads that change nothing.
    Read,
    /// In-place changes: flag, mark read, categorize, move, create folder/task.
    Modify,
    /// Create or edit drafts without sending.
    Draft,
    /// Send mail, replies, forwards and meeting invitations or responses.
    Send,
    /// Delete mail or events.
    Delete,
    /// Change mail rules.
    Rules,
    /// Create, update or delete calendar events.
    Events,
    /// Write files to disk (attachments).
    Files,
}

impl Capability {
    pub const ALL: [Capability; 8] = [
        Capability::Read,
        Capability::Modify,
        Capability::Draft,
        Capability::Send,
        Capability::Delete,
        Capability::Rules,
        Capability::Events,
        Capability::Files,
    ];

    pub fn name(self) -> &'static str {
        match self {
            Capability::Read => "read",
            Capability::Modify => "modify",
            Capability::Draft => "draft",
            Capability::Send => "send",
            Capability::Delete => "delete",
            Capability::Rules => "rules",
            Capability::Events => "events",
            Capability::Files => "files",
        }
    }

    /// Only pure reads may be repeated after a lost COM connection. Anything
    /// else could duplicate a send or delete if the first attempt took effect.
    pub fn idempotent(self) -> bool {
        self == Capability::Read
    }

    fn from_name(name: &str) -> Option<Self> {
        Self::ALL.into_iter().find(|c| c.name() == name)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    ReadOnly,
    Draft,
    Full,
}

impl Mode {
    pub fn name(self) -> &'static str {
        match self {
            Mode::ReadOnly => "read-only",
            Mode::Draft => "draft",
            Mode::Full => "full",
        }
    }

    fn grants(self, cap: Capability) -> bool {
        match self {
            Mode::ReadOnly => cap == Capability::Read,
            Mode::Draft => matches!(cap, Capability::Read | Capability::Draft),
            Mode::Full => true,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Policy {
    mode: Mode,
    /// Capabilities switched off on top of the mode. Can only narrow it.
    denied: Vec<Capability>,
}

impl Default for Policy {
    fn default() -> Self {
        Self::read_only()
    }
}

impl Policy {
    pub fn read_only() -> Self {
        Self {
            mode: Mode::ReadOnly,
            denied: Vec::new(),
        }
    }

    pub fn mode(&self) -> Mode {
        self.mode
    }

    /// `mode`: `read-only` (default), `draft` or `full`. `deny`: comma
    /// separated capability names removed from the mode.
    pub fn from_settings(mode: Option<&str>, deny: Option<&str>) -> Result<Self, String> {
        let mode = match mode.map(|m| m.trim().to_ascii_lowercase()).as_deref() {
            None | Some("") | Some("read-only") | Some("readonly") => Mode::ReadOnly,
            Some("draft") => Mode::Draft,
            Some("full") => Mode::Full,
            Some(other) => {
                return Err(format!(
                    "unknown OUTLOOK_MCP_MODE '{other}'; use read-only, draft or full"
                ));
            }
        };
        let mut denied = Vec::new();
        for part in deny.unwrap_or("").split(',') {
            let part = part.trim().to_ascii_lowercase();
            if part.is_empty() {
                continue;
            }
            let cap = Capability::from_name(&part)
                .ok_or_else(|| format!("unknown OUTLOOK_MCP_DENY capability '{part}'"))?;
            if cap == Capability::Read {
                return Err("OUTLOOK_MCP_DENY cannot disable read".into());
            }
            denied.push(cap);
        }
        Ok(Self { mode, denied })
    }

    pub fn from_env() -> Result<Self, String> {
        Self::from_settings(
            std::env::var("OUTLOOK_MCP_MODE").ok().as_deref(),
            std::env::var("OUTLOOK_MCP_DENY").ok().as_deref(),
        )
    }

    pub fn allows(&self, cap: Capability) -> bool {
        self.mode.grants(cap) && !self.denied.contains(&cap)
    }

    /// The error names the mode and capability but not how to raise them.
    pub fn check(&self, cap: Capability) -> Result<(), String> {
        if self.allows(cap) {
            Ok(())
        } else {
            Err(format!(
                "permission denied: '{}' operations are not allowed in {} mode. This is a server setting controlled by the user; do not retry or look for a workaround.",
                cap.name(),
                self.mode.name()
            ))
        }
    }

    pub fn describe(&self) -> String {
        let mut text = self.mode.name().to_string();
        for cap in &self.denied {
            text.push_str(&format!(" (deny {})", cap.name()));
        }
        text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_read_only_and_blocks_every_mutation() {
        let policy = Policy::default();
        assert_eq!(policy.mode(), Mode::ReadOnly);
        for cap in Capability::ALL {
            assert_eq!(policy.allows(cap), cap == Capability::Read, "{cap:?}");
        }
        assert!(policy.check(Capability::Send).is_err());
    }

    #[test]
    fn draft_mode_allows_only_drafts_beyond_reads() {
        let policy = Policy::from_settings(Some("draft"), None).unwrap();
        for cap in Capability::ALL {
            let expected = matches!(cap, Capability::Read | Capability::Draft);
            assert_eq!(policy.allows(cap), expected, "{cap:?}");
        }
    }

    #[test]
    fn full_mode_allows_all_and_deny_only_narrows() {
        let full = Policy::from_settings(Some(" FULL "), None).unwrap();
        assert!(Capability::ALL.into_iter().all(|c| full.allows(c)));
        let narrowed = Policy::from_settings(Some("full"), Some("send, Delete")).unwrap();
        assert!(!narrowed.allows(Capability::Send));
        assert!(!narrowed.allows(Capability::Delete));
        assert!(narrowed.allows(Capability::Modify));
        // Deny never widens a lower mode.
        let draft = Policy::from_settings(Some("draft"), Some("events")).unwrap();
        assert!(!draft.allows(Capability::Modify));
    }

    #[test]
    fn invalid_configuration_fails_closed() {
        for mode in ["admin", "write", "true", "fulll"] {
            assert!(Policy::from_settings(Some(mode), None).is_err(), "{mode}");
        }
        assert!(Policy::from_settings(Some("full"), Some("sned")).is_err());
        assert!(Policy::from_settings(None, Some("read")).is_err());
        // Unset or blank means the safest mode.
        assert_eq!(
            Policy::from_settings(None, None).unwrap().mode(),
            Mode::ReadOnly
        );
        assert_eq!(
            Policy::from_settings(Some("  "), None).unwrap().mode(),
            Mode::ReadOnly
        );
    }

    #[test]
    fn only_reads_are_idempotent() {
        for cap in Capability::ALL {
            assert_eq!(cap.idempotent(), cap == Capability::Read, "{cap:?}");
        }
    }

    #[test]
    fn denial_message_does_not_explain_how_to_escalate() {
        let message = Policy::read_only().check(Capability::Delete).unwrap_err();
        assert!(message.contains("read-only"));
        assert!(!message.contains("OUTLOOK_MCP"));
    }
}
