//! Which backend shows the approval prompt (docs/DIALOG-APPROVAL-DESIGN.md).
//! Pure selection logic; nothing here can approve anything by itself.

/// Visible lines in the dialog before the rest goes behind "Show full details".
const VISIBLE_LINES: usize = 12;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ApprovalMode {
    /// Native window on the interactive desktop (default).
    Dialog,
    /// MCP elicitation if the client offers it, else the native window.
    Auto,
    /// MCP elicitation only; an unsupported client means nothing happens.
    Elicitation,
    /// No approval backend: sends save a draft, other actions do nothing.
    None,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Backend {
    Dialog,
    Elicitation,
    Unavailable,
}

impl ApprovalMode {
    pub fn from_setting(value: Option<&str>) -> Result<Self, String> {
        match value.map(|v| v.trim().to_ascii_lowercase()).as_deref() {
            None | Some("") | Some("dialog") => Ok(Self::Dialog),
            Some("auto") => Ok(Self::Auto),
            Some("elicitation") => Ok(Self::Elicitation),
            Some("none") => Ok(Self::None),
            Some(other) => Err(format!(
                "unknown OUTLOOK_MCP_APPROVAL '{other}'; use dialog, auto, elicitation or none"
            )),
        }
    }

    pub fn from_env() -> Result<Self, String> {
        Self::from_setting(std::env::var("OUTLOOK_MCP_APPROVAL").ok().as_deref())
    }

    pub fn name(self) -> &'static str {
        match self {
            Self::Dialog => "dialog",
            Self::Auto => "auto",
            Self::Elicitation => "elicitation",
            Self::None => "none",
        }
    }

    pub fn choose(self, client_elicits: bool) -> Backend {
        match self {
            Self::Dialog => Backend::Dialog,
            Self::Auto if client_elicits => Backend::Elicitation,
            Self::Auto => Backend::Dialog,
            Self::Elicitation => Backend::Elicitation,
            Self::None => Backend::Unavailable,
        }
    }
}

/// Splits the server-built prompt into the always-visible part and the rest.
pub fn split_prompt(prompt: &str) -> (String, String) {
    let lines: Vec<&str> = prompt.trim().lines().collect();
    if lines.len() <= VISIBLE_LINES {
        return (lines.join("\n"), String::new());
    }
    let shown = lines[..VISIBLE_LINES].join("\n");
    (format!("{shown}\n..."), lines.join("\n"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_dialog_and_unknown_is_rejected() {
        assert_eq!(ApprovalMode::from_setting(None), Ok(ApprovalMode::Dialog));
        assert_eq!(
            ApprovalMode::from_setting(Some(" ")),
            Ok(ApprovalMode::Dialog)
        );
        assert!(ApprovalMode::from_setting(Some("yes")).is_err());
    }

    #[test]
    fn each_mode_picks_its_backend() {
        for elicits in [false, true] {
            assert_eq!(ApprovalMode::Dialog.choose(elicits), Backend::Dialog);
            assert_eq!(
                ApprovalMode::Elicitation.choose(elicits),
                Backend::Elicitation
            );
            assert_eq!(ApprovalMode::None.choose(elicits), Backend::Unavailable);
        }
        assert_eq!(ApprovalMode::Auto.choose(true), Backend::Elicitation);
        assert_eq!(ApprovalMode::Auto.choose(false), Backend::Dialog);
    }

    #[test]
    fn short_prompt_has_no_details_and_long_prompt_keeps_everything() {
        assert_eq!(split_prompt("a\nb"), ("a\nb".to_string(), String::new()));
        let long: String = (0..30).map(|i| format!("line {i}\n")).collect();
        let (shown, details) = split_prompt(&long);
        assert!(shown.ends_with("..."));
        assert_eq!(shown.lines().count(), VISIBLE_LINES + 1);
        assert!(details.contains("line 29"));
    }
}
