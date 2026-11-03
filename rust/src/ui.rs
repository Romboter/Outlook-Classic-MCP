//! MCP Apps (SEP-1865) views, served as `ui://` resources and linked from the read tools
//! through `_meta.ui.resourceUri`. The HTML is the Python server's own single copy
//! (`src/outlook_mcp/ui/`, MIT, upstream credit kept); hosts without MCP Apps support
//! ignore the metadata and show the plain text result.

use rmcp::model::{MetaObject, ReadResourceResult, Resource, ResourceContents};
use serde_json::json;

pub const MIME_TYPE: &str = "text/html;profile=mcp-app";

const COMMON: &str = include_str!("../../src/outlook_mcp/ui/_common.html");

struct View {
    name: &'static str,
    title: &'static str,
    description: &'static str,
    html: &'static str,
    /// Extra `ui` resource metadata (the contacts view copies to the clipboard).
    clipboard: bool,
}

const VIEWS: [View; 5] = [
    View {
        name: "mail-list",
        title: "Outlook mail list",
        description: "Interactive inbox view: open, mark, flag, and delete mails.",
        html: include_str!("../../src/outlook_mcp/ui/mail-list.html"),
        clipboard: false,
    },
    View {
        name: "mail-view",
        title: "Outlook mail reader",
        description: "Full reading view for a single mail with attachment list.",
        html: include_str!("../../src/outlook_mcp/ui/mail-view.html"),
        clipboard: false,
    },
    View {
        name: "calendar",
        title: "Outlook calendar agenda",
        description: "Agenda view of upcoming events grouped by day.",
        html: include_str!("../../src/outlook_mcp/ui/calendar.html"),
        clipboard: false,
    },
    View {
        name: "contacts",
        title: "Outlook contacts",
        description: "Searchable contact cards (saved contacts + org directory).",
        html: include_str!("../../src/outlook_mcp/ui/contacts.html"),
        clipboard: true,
    },
    View {
        name: "tasks",
        title: "Outlook tasks",
        description: "Task checklist: complete tasks and add new ones.",
        html: include_str!("../../src/outlook_mcp/ui/tasks.html"),
        clipboard: false,
    },
];

/// Python's mail items carry an ISO `received`; Rust returns Outlook's text as `received_local`.
/// The views call `new Date(received)` and fall back to showing the text as it is, so add
/// `received`: ISO local time when the text has the `yyyy/mm/dd:hh:mm:ss` shape, else the
/// text unchanged. Only the structured copy changes; the text for the model does not.
pub fn add_received(value: &mut serde_json::Value) {
    fn iso(text: &str) -> Option<String> {
        let b = text.as_bytes();
        let digits =
            |r: std::ops::Range<usize>| b.get(r).is_some_and(|s| s.iter().all(u8::is_ascii_digit));
        let shaped = b.len() >= 19
            && digits(0..4)
            && b[4] == b'/'
            && digits(5..7)
            && b[7] == b'/'
            && digits(8..10)
            && b[10] == b':'
            && digits(11..13)
            && b[13] == b':'
            && digits(14..16)
            && b[16] == b':'
            && digits(17..19);
        shaped.then(|| {
            format!(
                "{}-{}-{}T{}",
                &text[0..4],
                &text[5..7],
                &text[8..10],
                &text[11..19]
            )
        })
    }
    fn alias(item: &mut serde_json::Value) {
        let Some(local) = item.get("received_local").and_then(|v| v.as_str()) else {
            return;
        };
        let shown = iso(local).unwrap_or_else(|| local.to_string());
        item["received"] = serde_json::Value::String(shown);
    }
    alias(value);
    if let Some(items) = value.get_mut("items").and_then(|v| v.as_array_mut()) {
        items.iter_mut().for_each(alias);
    }
}

pub fn view_uri(name: &str) -> String {
    format!("ui://outlook/{name}.html")
}

fn meta_object(value: serde_json::Value) -> MetaObject {
    match value {
        serde_json::Value::Object(map) => MetaObject(map),
        _ => MetaObject::new(),
    }
}

/// `_meta` for a tool that renders `view` when called.
pub fn tool_meta(name: &str) -> MetaObject {
    debug_assert!(VIEWS.iter().any(|v| v.name == name), "unknown view {name}");
    meta_object(json!({ "ui": { "resourceUri": view_uri(name) } }))
}

fn resource_meta(view: &View) -> MetaObject {
    let mut ui = json!({ "prefersBorder": true });
    if view.clipboard {
        ui["permissions"] = json!({ "clipboardWrite": {} });
    }
    meta_object(json!({ "ui": ui }))
}

pub fn resources() -> Vec<Resource> {
    VIEWS
        .iter()
        .map(|v| {
            Resource::new(view_uri(v.name), format!("outlook-ui-{}", v.name))
                .with_title(v.title)
                .with_description(v.description)
                .with_mime_type(MIME_TYPE)
                .with_meta(resource_meta(v))
        })
        .collect()
}

fn render(view: &View) -> String {
    view.html.replace("<!--%COMMON%-->", COMMON)
}

pub fn read(uri: &str) -> Option<ReadResourceResult> {
    let view = VIEWS.iter().find(|v| view_uri(v.name) == uri)?;
    let contents = ResourceContents::text(render(view), uri)
        .with_mime_type(MIME_TYPE)
        .with_meta(resource_meta(view));
    Some(ReadResourceResult::new(vec![contents]))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_view_is_listed_readable_and_has_the_shared_bridge_injected() {
        let listed = resources();
        assert_eq!(listed.len(), 5);
        for resource in &listed {
            assert_eq!(resource.mime_type.as_deref(), Some(MIME_TYPE));
            let result = read(&resource.uri).expect("listed view must be readable");
            let ResourceContents::TextResourceContents { text, .. } = &result.contents[0] else {
                panic!("expected text contents");
            };
            assert!(
                !text.contains("<!--%COMMON%-->"),
                "{} kept its placeholder",
                resource.uri
            );
            assert!(
                text.contains("appInfo"),
                "{} lacks the fixed ui/initialize bridge",
                resource.uri
            );
        }
        assert!(read("ui://outlook/nope.html").is_none());
    }

    #[test]
    fn received_alias_is_iso_when_shaped_and_raw_text_otherwise() {
        let mut v = json!({"received_local": "2026/11/02:18:08:10.000", "items": [
            {"received_local": "2026/11/02:18:08:10.000"}, {"received_local": "02.11.2026 18:08"}, {"subject": "x"}]});
        add_received(&mut v);
        assert_eq!(v["received"], "2026-11-02T18:08:10");
        assert_eq!(v["items"][0]["received"], "2026-11-02T18:08:10");
        assert_eq!(v["items"][1]["received"], "02.11.2026 18:08");
        assert!(v["items"][2].get("received").is_none());
        assert_eq!(v["items"][0]["received_local"], "2026/11/02:18:08:10.000");
    }

    #[test]
    fn tool_meta_points_at_the_view_uri() {
        let meta = tool_meta("mail-list");
        assert_eq!(meta.0["ui"]["resourceUri"], "ui://outlook/mail-list.html");
    }
}
