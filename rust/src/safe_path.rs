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

fn is_executable(name: &str) -> bool {
    name.rsplit_once('.')
        .map(|(_, e)| e.to_ascii_lowercase())
        .is_some_and(|e| EXECUTABLE.contains(&e.as_str()))
}

/// Most files one outgoing message may carry.
pub const MAX_ATTACH_FILES: usize = 10;
/// Most bytes (all files together) one outgoing message may carry: 20 MiB.
pub const MAX_ATTACH_TOTAL_BYTES: u64 = 20 * 1024 * 1024;

/// A file that passed the outgoing-attachment checks.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AttachFile {
    /// Canonical path in the form Outlook's `Attachments.Add` accepts.
    pub path: String,
    pub name: String,
    pub size: u64,
}

/// Check every requested attachment path: inside an allowed root (after
/// symlink resolution, so `..` and links cannot escape), an existing regular
/// file, not an executable or script, and within the count and size caps.
/// Each refusal names the file and gives the reason.
pub fn resolve_attachment_files(
    requested: &[String],
    allowed_roots: &[PathBuf],
) -> Result<Vec<AttachFile>, String> {
    if requested.len() > MAX_ATTACH_FILES {
        return Err(format!(
            "too many attachments ({}); at most {MAX_ATTACH_FILES} are allowed",
            requested.len()
        ));
    }
    let roots: Vec<PathBuf> = allowed_roots
        .iter()
        .filter_map(|r| r.canonicalize().ok())
        .collect();
    let mut files: Vec<AttachFile> = Vec::new();
    let mut total: u64 = 0;
    for text in requested {
        let shown: String = text.chars().take(80).collect();
        if text.trim().is_empty() || text.chars().any(char::is_control) {
            return Err("an attachment path is empty or contains control characters".into());
        }
        let given = Path::new(text);
        if given
            .components()
            .any(|c| matches!(c, std::path::Component::ParentDir))
        {
            return Err(format!("{shown} was not attached: path traversal (..)"));
        }
        let real = given.canonicalize().map_err(|_| {
            format!("{shown} was not attached: file does not exist or is not accessible")
        })?;
        let name = real
            .file_name()
            .map(|n| n.to_string_lossy().into_owned())
            .unwrap_or_default();
        if !roots.iter().any(|root| real.starts_with(root)) {
            return Err(format!(
                "{name} was not attached: it is outside the allowed locations (your user profile)"
            ));
        }
        let meta = std::fs::metadata(&real)
            .map_err(|_| format!("{name} was not attached: file is not accessible"))?;
        if !meta.is_file() {
            return Err(format!("{name} was not attached: it is not a regular file"));
        }
        if let Err(why) = validate_filename(&name) {
            return Err(format!("{name} was not attached: {why}"));
        }
        if is_executable(&name) {
            return Err(format!(
                "{name} was not attached: executable and script files can run code when opened"
            ));
        }
        total = total.saturating_add(meta.len());
        if total > MAX_ATTACH_TOTAL_BYTES {
            return Err(format!(
                "attachments are too large (limit {} MiB in total; {name} would exceed it)",
                MAX_ATTACH_TOTAL_BYTES / (1024 * 1024)
            ));
        }
        let path = plain_path(&real);
        if files.iter().any(|f| f.path.eq_ignore_ascii_case(&path)) {
            return Err(format!("{name} is listed more than once"));
        }
        files.push(AttachFile {
            path,
            name,
            size: meta.len(),
        });
    }
    Ok(files)
}

/// Why a file must not be saved, naming the file, or `None` when it may be.
pub fn refusal_reason(name: &str) -> Option<String> {
    if let Err(why) = validate_filename(name) {
        return Some(format!("{name} was not saved: {why}"));
    }
    if is_executable(name) {
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
    fn outgoing_attachments_are_checked_and_refused_by_name() {
        let root = scratch("attach-root");
        let outside = scratch("attach-outside");
        let roots = vec![root.clone()];
        let write = |dir: &Path, name: &str, len: usize| {
            let p = dir.join(name);
            std::fs::write(&p, vec![b'x'; len]).unwrap();
            p.to_string_lossy().into_owned()
        };
        let ok = write(&root, "report.pdf", 10);
        let got = resolve_attachment_files(std::slice::from_ref(&ok), &roots).unwrap();
        assert_eq!((got[0].name.as_str(), got[0].size), ("report.pdf", 10));
        assert!(!got[0].path.starts_with(r"\\?\"));
        assert_eq!(resolve_attachment_files(&[], &roots).unwrap(), vec![]);

        let exe = write(&root, "setup.exe", 1);
        let e = resolve_attachment_files(&[exe], &roots).unwrap_err();
        assert!(
            e.starts_with("setup.exe was not attached") && e.contains("run code"),
            "{e}"
        );
        let far = write(&outside, "a.txt", 1);
        let e = resolve_attachment_files(&[far], &roots).unwrap_err();
        assert!(e.contains("outside"), "{e}");
        let e = resolve_attachment_files(&[format!("{}/../x.txt", root.display())], &roots)
            .unwrap_err();
        assert!(e.contains("traversal"), "{e}");
        let missing = root.join("nope.txt").to_string_lossy().into_owned();
        let e = resolve_attachment_files(&[missing], &roots).unwrap_err();
        assert!(e.contains("does not exist"), "{e}");
        let e = resolve_attachment_files(&[root.to_string_lossy().into()], &roots).unwrap_err();
        assert!(e.contains("not a regular file"), "{e}");
        let e = resolve_attachment_files(&[ok.clone(), ok.clone()], &roots).unwrap_err();
        assert!(e.contains("more than once"), "{e}");
        let e =
            resolve_attachment_files(&vec![ok.clone(); MAX_ATTACH_FILES + 1], &roots).unwrap_err();
        assert!(e.contains("too many"), "{e}");
        let _ = std::fs::remove_dir_all(root);
        let _ = std::fs::remove_dir_all(outside);
    }

    #[test]
    fn outgoing_attachments_total_size_is_capped() {
        let root = scratch("attach-size");
        let big = root.join("big.bin");
        let f = std::fs::File::create(&big).unwrap();
        f.set_len(MAX_ATTACH_TOTAL_BYTES + 1).unwrap();
        let e =
            resolve_attachment_files(&[big.to_string_lossy().into()], std::slice::from_ref(&root))
                .unwrap_err();
        assert!(e.contains("too large") && e.contains("big.bin"), "{e}");
        let _ = std::fs::remove_dir_all(root);
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
