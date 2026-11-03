//! Validation for files `outlook_save_attachments` may write.
use std::fs::{File, OpenOptions};
use std::io;
use std::path::{Path, PathBuf};

const RESERVED: [&str; 22] = [
    "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8",
    "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
];

/// A single path component that is safe to create on Windows.
pub fn validate_filename(name: &str) -> Result<&str, String> {
    let bad = |why: &str| Err(format!("unsafe attachment file name: {why}"));
    if name.is_empty() || name.chars().count() > 150 {
        return bad("length must be 1..150 characters");
    }
    if name == "." || name == ".." {
        return bad("relative path component");
    }
    if name.chars().any(|c| {
        c.is_control() || matches!(c, '/' | '\\' | ':' | '<' | '>' | '"' | '|' | '?' | '*')
    }) {
        return bad("contains a path separator, stream marker or reserved character");
    }
    if name.ends_with('.') || name.ends_with(' ') {
        return bad("trailing dot or space");
    }
    let stem = name.split('.').next().unwrap_or("").trim_end();
    if RESERVED.iter().any(|r| stem.eq_ignore_ascii_case(r)) {
        return bad("reserved device name");
    }
    Ok(name)
}

/// Extensions that can run code when opened. Checked on the last extension, so
/// `invoice.pdf.exe` is caught.
const EXECUTABLE: [&str; 28] = [
    "exe",
    "dll",
    "js",
    "jse",
    "vbs",
    "vbe",
    "ps1",
    "psm1",
    "bat",
    "cmd",
    "com",
    "lnk",
    "scr",
    "msi",
    "msp",
    "hta",
    "cpl",
    "jar",
    "reg",
    "wsf",
    "wsh",
    "pif",
    "inf",
    "url",
    "msc",
    "gadget",
    "application",
    "appref-ms",
];

/// Why a file must not be saved, naming the file, or `None` when it may be.
pub fn refusal_reason(name: &str) -> Option<String> {
    if let Err(why) = validate_filename(name) {
        return Some(format!("{name} was not saved: {why}"));
    }
    let extension = name.rsplit_once('.').map(|(_, e)| e.to_ascii_lowercase());
    if extension.is_some_and(|e| EXECUTABLE.contains(&e.as_str())) {
        return Some(format!(
            "{name} was not saved: executable and script files can run code when opened"
        ));
    }
    None
}

/// The path as Outlook's `SaveAsFile` wants it: canonicalizing adds a
/// verbatim `\\?\` prefix that Outlook does not accept.
pub fn plain_path(path: &Path) -> String {
    let text = path.to_string_lossy();
    text.strip_prefix(r"\\?\").unwrap_or(&text).to_string()
}

/// Resolve `requested` to a real directory inside one of `allowed_roots`.
/// Canonicalization follows symlinks and junctions, so a link that points
/// outside an allowed root is rejected along with `..` traversal.
pub fn resolve_output_dir(requested: &Path, allowed_roots: &[PathBuf]) -> Result<PathBuf, String> {
    let dir = requested
        .canonicalize()
        .map_err(|_| "output directory does not exist or is not accessible".to_string())?;
    if !dir.is_dir() {
        return Err("output path is not a directory".into());
    }
    let inside = allowed_roots
        .iter()
        .filter_map(|root| root.canonicalize().ok())
        .any(|root| dir.starts_with(root));
    if inside {
        Ok(dir)
    } else {
        Err("output directory is outside the allowed locations".into())
    }
}

/// Create `filename` in `dir` without ever overwriting: a taken name gets
/// ` (1)`, ` (2)`... `create_new` makes the check atomic, so there is no
/// window between choosing a name and creating it.
pub fn create_unique(dir: &Path, filename: &str) -> io::Result<(PathBuf, File)> {
    validate_filename(filename).map_err(|e| io::Error::new(io::ErrorKind::InvalidInput, e))?;
    let (stem, ext) = match filename.rfind('.') {
        Some(i) if i > 0 => (&filename[..i], &filename[i..]),
        _ => (filename, ""),
    };
    for n in 0..100 {
        let name = if n == 0 {
            filename.to_string()
        } else {
            format!("{stem} ({n}){ext}")
        };
        let path = dir.join(name);
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => return Ok((path, file)),
            Err(e) if e.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(e) => return Err(e),
        }
    }
    Err(io::Error::new(
        io::ErrorKind::AlreadyExists,
        "too many files with this name",
    ))
}

/// Default allowlist: the signed-in user's profile only.
pub fn default_roots() -> Vec<PathBuf> {
    std::env::var_os("USERPROFILE")
        .map(PathBuf::from)
        .into_iter()
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn scratch(tag: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!(
            "outlook-mcp-safe-path-{}-{tag}",
            std::process::id()
        ));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn executables_and_scripts_are_refused_with_the_file_named() {
        for bad in [
            "invoice.exe",
            "Run.PS1",
            "a.pdf.exe",
            "x.lnk",
            "setup.MSI",
            "con.txt",
        ] {
            let reason = refusal_reason(bad).expect(bad);
            assert!(reason.starts_with(bad), "{reason}");
        }
        assert!(refusal_reason("invoice.exe").unwrap().contains("run code"));
        for ok in ["report.pdf", "photo.jpeg", "notes", "exe.txt", "macro.docm"] {
            assert!(refusal_reason(ok).is_none(), "{ok}");
        }
    }

    #[test]
    fn plain_path_drops_the_verbatim_prefix() {
        assert_eq!(plain_path(Path::new(r"\\?\C:\Users\a")), r"C:\Users\a");
        assert_eq!(plain_path(Path::new(r"C:\x")), r"C:\x");
    }

    #[test]
    fn accepts_ordinary_names_and_rejects_dangerous_ones() {
        for ok in [
            "report.pdf",
            "a b.txt",
            "archive.tar.gz",
            "CONFIG.txt",
            "COM10.txt",
        ] {
            assert!(validate_filename(ok).is_ok(), "{ok}");
        }
        for bad in [
            "",
            ".",
            "..",
            "../x",
            "a/b",
            "a\\b",
            "C:evil",
            "x.txt:stream",
            "con",
            "NUL.txt",
            "Com1.log",
            "lpt9",
            "aux.tar.gz",
            "trail.",
            "trail ",
            "a<b",
            "a|b",
            "a?b",
            "a*b",
            "tab\there",
            "new\nline",
        ] {
            assert!(validate_filename(bad).is_err(), "{bad:?}");
        }
        assert!(validate_filename(&"x".repeat(151)).is_err());
    }

    #[test]
    fn directory_must_exist_and_sit_inside_an_allowed_root() {
        let root = scratch("root");
        let inner = root.join("inner");
        std::fs::create_dir_all(&inner).unwrap();
        let outside = scratch("outside");
        let roots = vec![root.clone()];
        assert!(resolve_output_dir(&inner, &roots).is_ok());
        assert!(resolve_output_dir(&root, &roots).is_ok());
        assert!(resolve_output_dir(&outside, &roots).is_err());
        assert!(resolve_output_dir(&root.join("missing"), &roots).is_err());
        // Traversal out of the root through `..` resolves outside and fails.
        let traversal = inner
            .join("..")
            .join("..")
            .join(outside.file_name().unwrap());
        assert!(resolve_output_dir(&traversal, &roots).is_err());
        let _ = std::fs::remove_dir_all(root);
        let _ = std::fs::remove_dir_all(outside);
    }

    #[cfg(windows)]
    #[test]
    fn symlink_out_of_root_is_rejected() {
        let root = scratch("link-root");
        let outside = scratch("link-outside");
        let link = root.join("escape");
        // Creating symlinks needs a privilege on Windows; skip when missing.
        if std::os::windows::fs::symlink_dir(&outside, &link).is_ok() {
            assert!(resolve_output_dir(&link, std::slice::from_ref(&root)).is_err());
        }
        let _ = std::fs::remove_dir_all(root);
        let _ = std::fs::remove_dir_all(outside);
    }

    #[test]
    fn never_overwrites_an_existing_file() {
        let dir = scratch("unique");
        let (first, _) = create_unique(&dir, "a.txt").unwrap();
        let (second, _) = create_unique(&dir, "a.txt").unwrap();
        let (third, _) = create_unique(&dir, "a.txt").unwrap();
        assert_eq!(first.file_name().unwrap(), "a.txt");
        assert_eq!(second.file_name().unwrap(), "a (1).txt");
        assert_eq!(third.file_name().unwrap(), "a (2).txt");
        assert!(create_unique(&dir, "../escape.txt").is_err());
        let _ = std::fs::remove_dir_all(dir);
    }
}
