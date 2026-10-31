//! Pure, platform-independent thread text processing.
//!
//! Never infer that an unrecognized block is quoted history. Remove only a
//! contiguous, terminal >-prefixed block or an explicit original-message
//! delimiter. Report every suppression and offer full raw read by message ID.

#[derive(Debug, PartialEq, Eq)]
pub struct Trimmed {
    pub text: String,
    pub quote_removed: bool,
    pub chars_removed: usize,
}

fn is_original_separator(line: &str) -> bool {
    let value = line.trim().to_ascii_lowercase();
    matches!(value.as_str(),
        "-----original message-----"
        | "----- forwarded message -----"
        | "begin forwarded message:"
    )
}

/// Remove only recognized terminal quoted history. Keep the original body
/// entirely when a separator would be at the first line or the shape is
/// ambiguous. This is not a semantic conversation summarizer.
pub fn trim_quote_tail(body: &str) -> Trimmed {
    let lines: Vec<&str> = body.lines().collect();
    let cut = lines.iter().enumerate().find_map(|(i, line)| {
        if i == 0 || !is_original_separator(line) { return None; }
        let prefix = lines[..i].join("\n");
        (!prefix.trim().is_empty()).then_some(prefix)
    }).or_else(|| {
        // Only terminal blocks of >-prefixed lines count.
        let mut start = lines.len();
        for (i, line) in lines.iter().enumerate().rev() {
            if line.trim().is_empty() && start == lines.len() { continue; }
            if line.trim_start().starts_with('>') || (line.trim().is_empty() && start < lines.len()) {
                start = i;
            } else { break; }
        }
        if start > 0 && start < lines.len()
            && lines[start..].iter().any(|s| s.trim_start().starts_with('>'))
        {
            let prefix = lines[..start].join("\n");
            (!prefix.trim().is_empty()).then_some(prefix)
        } else { None }
    });

    if let Some(prefix) = cut {
        let removed = body.chars().count().saturating_sub(prefix.chars().count());
        if removed > 0 {
            return Trimmed {text:prefix.trim_end().to_owned(),quote_removed:true,chars_removed:removed};
        }
    }
    Trimmed { text:body.to_owned(), quote_removed:false, chars_removed:0 }
}

/// Apply a Unicode-character budget without splitting multibyte characters.
pub fn limit_chars(body: &str, budget: usize) -> (String, bool) {
    let mut iter = body.chars();
    let retained: String = iter.by_ref().take(budget).collect();
    let truncated = iter.next().is_some();
    (retained, truncated)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn removes_only_explicit_original_history_after_new_content() {
        let b = "Please deploy after lunch.\n-----Original Message-----\nFrom: A\nYesterday";
        let r = trim_quote_tail(b);
        assert_eq!(r.text, "Please deploy after lunch.");
        assert!(r.quote_removed);
        assert!(r.chars_removed > 0);
    }
    #[test]
    fn removes_terminal_angle_quote_but_not_inline_new_material() {
        let r=trim_quote_tail("New fix applied\n> old request\n> previous text");
        assert_eq!(r.text,"New fix applied");
        let raw=trim_quote_tail("New fix\n> quote\nAdditional new instruction");
        assert_eq!(raw.text,"New fix\n> quote\nAdditional new instruction");
        assert!(!raw.quote_removed);
    }
    #[test]
    fn never_removes_initial_separator_or_unrecognized_date_line() {
        let body="-----Original Message-----\nOnly content here";
        assert_eq!(trim_quote_tail(body).text, body);
        let ambiguous="New decision\nOn Friday Alex wrote:\nMove to phase two";
        assert_eq!(trim_quote_tail(ambiguous).text, ambiguous);
    }
    #[test]
    fn unicode_budget_and_truncation_are_precise() {
        assert_eq!(limit_chars("🙂éA", 2),("🙂é".to_owned(),true));
        assert_eq!(limit_chars("🙂é", 2),("🙂é".to_owned(),false));
        assert_eq!(limit_chars("🙂é", 0),("".to_owned(),true));
    }
    #[test]
    fn empty_or_unusual_text_is_kept() {
        for body in ["", ">", "\n\n", "Please approve", "Hi\n>quote\nThanks"] {
            assert_eq!(trim_quote_tail(body).text,body);
        }
    }
}
