//! Session-scoped, bounded opaque aliases for Outlook EntryID + StoreID.
//! An MCP tool response must not include the long raw identifiers, but the
//! COM worker keeps them to resolve an exact selected message later.
//!
//! These are convenience references, NOT authentication tokens or security
//! boundaries: the server already has the permissions of the Outlook user.
use crate::contracts::GetMailParams;
use anyhow::{Result, bail};
use serde_json::{Map, Value};
use std::collections::{HashMap, VecDeque};

const DEFAULT_CAPACITY: usize = 2048;

#[derive(Default)]
pub struct MailRefs {
    next: u64,
    by_handle: HashMap<String, (String, String)>,
    by_identity: HashMap<(String, String), String>,
    insertion_order: VecDeque<String>,
    max_entries: usize,
}

impl MailRefs {
    pub fn new() -> Self {
        Self::with_capacity(DEFAULT_CAPACITY)
    }

    fn with_capacity(max_entries: usize) -> Self {
        Self {
            next: 0,
            by_handle: HashMap::new(),
            by_identity: HashMap::new(),
            insertion_order: VecDeque::new(),
            max_entries: max_entries.max(1),
        }
    }

    fn remember(&mut self, entry_id: String, store_id: String) -> String {
        let identity = (entry_id, store_id);
        if let Some(existing) = self.by_identity.get(&identity) {
            return existing.clone();
        }
        if self.by_handle.len() >= self.max_entries
            && let Some(oldest) = self.insertion_order.pop_front()
            && let Some(old_identity) = self.by_handle.remove(&oldest)
        {
            self.by_identity.remove(&old_identity);
        }

        self.next += 1;
        let handle = format!("m{}", self.next);
        self.by_identity.insert(identity.clone(), handle.clone());
        self.by_handle.insert(handle.clone(), identity);
        self.insertion_order.push_back(handle.clone());
        handle
    }

    fn compact_item(&mut self, item: &mut Map<String, Value>) {
        let Some(entry_id) = item
            .get("entry_id")
            .and_then(Value::as_str)
            .map(str::to_owned)
        else {
            return;
        };
        // Never publish the full StoreID even for items that have no EntryID.
        let store_id = item
            .remove("store_id")
            .and_then(|v| v.as_str().map(str::to_owned))
            .unwrap_or_default();

        if entry_id.is_empty() {
            item.remove("entry_id");
            return;
        }
        let handle = self.remember(entry_id, store_id);
        // Preserve the existing argument name 'entry_id' for caller
        // compatibility while replacing its value with a short session handle.
        item.insert("entry_id".into(), Value::String(handle));
    }

    pub fn compact_response(&mut self, response: &mut Value) {
        let Value::Object(root) = response else {
            return;
        };
        // Search/list use "items"; thread retrieval uses "messages".
        // Both MUST be sanitized before any MCP response is serialized.
        for key in ["items", "messages"] {
            if let Some(Value::Array(items)) = root.get_mut(key) {
                for item in items {
                    if let Value::Object(item) = item {
                        self.compact_item(item);
                    }
                }
                return;
            }
        }
        self.compact_item(root);
    }

    pub fn resolve_input(&self, params: &mut GetMailParams) -> Result<()> {
        let id = &params.entry_id;
        if let Some((entry, store)) = self.by_handle.get(id) {
            params.entry_id = entry.clone();
            params.store_id = if store.is_empty() {
                None
            } else {
                Some(store.clone())
            };
            return Ok(());
        }
        // Reject stale server-generated handles. Raw Outlook EntryIDs are
        // still accepted for compatibility with the original Python server.
        if is_handle(id) {
            bail!("Unknown or expired message handle; run search/list again in this MCP session");
        }
        Ok(())
    }
}

fn is_handle(s: &str) -> bool {
    s.strip_prefix('m')
        .is_some_and(|rest| !rest.is_empty() && rest.bytes().all(|b| b.is_ascii_digit()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn selected(id: &str, store: Option<&str>) -> GetMailParams {
        GetMailParams {
            entry_id: id.into(),
            store_id: store.map(str::to_owned),
            include_body: Some(false),
            include_html: None,
            max_body_chars: None,
        }
    }

    #[test]
    fn compact_output_never_exposes_long_identifiers() {
        let mut refs = MailRefs::new();
        let mut result = json!({"items": [
            {"entry_id": "private-entry-id", "store_id": "private-store-id", "subject": "Test"},
            {"entry_id": "private-entry-id-2", "store_id": "private-store-id", "subject": "Other"}
        ]});
        refs.compact_response(&mut result);
        let output = result.to_string();
        assert!(!output.contains("private-entry-id"));
        assert!(!output.contains("private-store-id"));
        assert_eq!(result["items"][0]["entry_id"], "m1");
        assert_eq!(result["items"][1]["entry_id"], "m2");
        assert!(result["items"][0].get("store_id").is_none());
    }

    #[test]
    fn resolves_exact_entry_and_store_then_preserves_handle() {
        let mut refs = MailRefs::new();
        let mut result = json!({"entry_id": "long-1", "store_id": "store-2"});
        refs.compact_response(&mut result);
        assert_eq!(result["entry_id"], "m1");
        let mut params = selected("m1", Some("wrong-store"));
        refs.resolve_input(&mut params).unwrap();
        assert_eq!(params.entry_id, "long-1");
        assert_eq!(params.store_id.as_deref(), Some("store-2"));
        let mut duplicate = json!({"entry_id":"long-1","store_id":"store-2"});
        refs.compact_response(&mut duplicate);
        assert_eq!(duplicate["entry_id"], "m1");
    }

    #[test]
    fn distinct_stores_do_not_share_handles() {
        let mut refs = MailRefs::new();
        let mut a = json!({"entry_id":"same","store_id":"store-a"});
        let mut b = json!({"entry_id":"same","store_id":"store-b"});
        refs.compact_response(&mut a);
        refs.compact_response(&mut b);
        assert_ne!(a["entry_id"], b["entry_id"]);
    }

    #[test]
    fn eviction_and_stale_handles_fail_closed() {
        let mut refs = MailRefs::with_capacity(2);
        for n in 1..=3 {
            let mut item = json!({"entry_id":format!("entry-{n}"),"store_id":"store"});
            refs.compact_response(&mut item);
        }
        assert!(refs.resolve_input(&mut selected("m1", None)).is_err());
        let mut params = selected("m3", None);
        refs.resolve_input(&mut params).unwrap();
        assert_eq!(params.entry_id, "entry-3");
        assert!(refs.resolve_input(&mut selected("m999", None)).is_err());
    }

    #[test]
    fn thread_messages_never_leak_entry_or_store_id() {
        let mut refs = MailRefs::new();
        let mut output = json!({"messages":[
            {"entry_id":"super-long-secret-entry","store_id":"super-long-secret-store","body":"Hello"},
            {"entry_id":"second-entry","store_id":"second-store","body":"Reply"}
        ],"count":2});
        refs.compact_response(&mut output);
        assert_eq!(output["messages"][0]["entry_id"], "m1");
        assert_eq!(output["messages"][1]["entry_id"], "m2");
        assert!(!output.to_string().contains("secret"));
        assert!(output["messages"][0].get("store_id").is_none());
        let mut request = selected("m1", None);
        refs.resolve_input(&mut request).unwrap();
        assert_eq!(request.entry_id, "super-long-secret-entry");
        assert_eq!(request.store_id.as_deref(), Some("super-long-secret-store"));
    }

    #[test]
    fn raw_entry_id_remains_compatible() {
        let refs = MailRefs::new();
        let mut input = selected("AAAA-real-outlook-entry-id", None);
        refs.resolve_input(&mut input).unwrap();
        assert_eq!(input.entry_id, "AAAA-real-outlook-entry-id");
        assert!(!is_handle("msg42"));
        assert!(!is_handle("m"));
        assert!(is_handle("m42"));
    }
}
