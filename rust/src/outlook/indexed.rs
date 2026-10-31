//! Event-driven AdvancedSearch + Search.GetTable backend.
//!
//! The event sink is a small private COM IDispatch implementation living only
//! on the same STA as Outlook. Never use Search.Results before completion.
//! Opt-in backend only until validated against a real Outlook Classic profile.
use super::{table, win};
use crate::contracts::{SearchMailsParams, checked_limit, escape_dasl};
use anyhow::{Context, Result, bail};
use serde_json::{Value, json};
use std::{
    cell::Cell,
    ffi::c_void,
    ptr,
    rc::Rc,
    sync::atomic::{AtomicU32, AtomicU64, Ordering},
    thread,
    time::{Duration, Instant},
};
use windows::{
    core::{GUID, IUnknown, Interface},
    Win32::System::{
        Com::{DISPPARAMS, IConnectionPoint, IConnectionPointContainer, IDispatch, DISPATCH_METHOD},
        Variant::VARIANT,
    },
};

// Outlook ApplicationEvents_11 (dispatch) interface and the documented
// AdvancedSearchComplete dispid. Pin to this source interface, not IDispatch
// alone: ConnectionPoint.Advise queries for the outgoing interface IID.
const EVENT_IID: GUID = GUID::from_u128(0x0006302c_0000_0000_c000_000000000046);
const IID_IUNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_c000_000000000046);
const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
const ADVANCED_SEARCH_COMPLETE: i32 = 64106;
const SEARCH_TIMEOUT: Duration = Duration::from_secs(18);
static NEXT_TAG: AtomicU64 = AtomicU64::new(0);

const S_OK: i32 = 0;
const E_POINTER: i32 = 0x80004003_u32 as i32;
const E_NOTIMPL: i32 = 0x80004001_u32 as i32;
const E_NOINTERFACE: i32 = 0x80004002_u32 as i32;

// IDispatch is a pure vtable contract. For the Outlook event dispinterface,
// QueryInterface must support the _source IID_ as well as IDispatch.
#[repr(C)]
struct SinkVtbl {
    query_interface: unsafe extern "system" fn(*mut Sink, *const GUID, *mut *mut c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut Sink) -> u32,
    release: unsafe extern "system" fn(*mut Sink) -> u32,
    get_type_info_count: unsafe extern "system" fn(*mut Sink, *mut u32) -> i32,
    get_type_info: unsafe extern "system" fn(*mut Sink, u32, u32, *mut *mut c_void) -> i32,
    get_ids_of_names:
        unsafe extern "system" fn(*mut Sink, *const GUID, *const *const u16, u32, u32, *mut i32) -> i32,
    invoke: unsafe extern "system" fn(
        *mut Sink, i32, *const GUID, u32, u16, *mut DISPPARAMS, *mut VARIANT,
        *mut c_void, *mut u32,
    ) -> i32,
}

#[repr(C)]
struct Sink {
    vtable: *const SinkVtbl,
    references: AtomicU32,
    expected_tag: String,
    completed: Rc<Cell<bool>>,
}

unsafe extern "system" fn query_interface(
    this: *mut Sink, iid: *const GUID, out: *mut *mut c_void,
) -> i32 {
    if iid.is_null() || out.is_null() { return E_POINTER; }
    // SAFETY: Called by COM with a valid live Sink. Return one additional
    // COM ref for every successfully queried supported interface.
    unsafe {
        *out = ptr::null_mut();
        if *iid == EVENT_IID || *iid == IID_IUNKNOWN || *iid == IID_IDISPATCH {
            *out = this.cast();
            add_ref(this);
            S_OK
        } else {
            E_NOINTERFACE
        }
    }
}
unsafe extern "system" fn add_ref(this: *mut Sink) -> u32 {
    unsafe { (*this).references.fetch_add(1, Ordering::Relaxed) + 1 }
}
unsafe extern "system" fn release(this: *mut Sink) -> u32 {
    let remaining = unsafe { (*this).references.fetch_sub(1, Ordering::Release) - 1 };
    if remaining == 0 {
        std::sync::atomic::fence(Ordering::Acquire);
        unsafe { drop(Box::from_raw(this)); }
    }
    remaining
}
unsafe extern "system" fn get_type_info_count(_this: *mut Sink, count: *mut u32) -> i32 {
    if count.is_null() { return E_POINTER; }
    unsafe { *count = 0; }
    S_OK
}
unsafe extern "system" fn get_type_info(
    _this: *mut Sink, _index: u32, _lcid: u32, _info: *mut *mut c_void,
) -> i32 { E_NOTIMPL }
unsafe extern "system" fn get_ids_of_names(
    _this: *mut Sink, _iid: *const GUID, _names: *const *const u16,
    _count: u32, _lcid: u32, _dispids: *mut i32,
) -> i32 { E_NOTIMPL }

unsafe extern "system" fn invoke_event(
    this: *mut Sink, dispid: i32, _iid: *const GUID, _lcid: u32, _flags: u16,
    args: *mut DISPPARAMS, _result: *mut VARIANT, _exception: *mut c_void,
    _argerr: *mut u32,
) -> i32 {
    if dispid != ADVANCED_SEARCH_COMPLETE || args.is_null() { return S_OK; }
    // COM calls us on our registered STA. The event argument is the Search
    // dispatch interface; examine its Tag to avoid accepting someone else's
    // concurrent AdvancedSearchComplete notification.
    unsafe {
        let params = &*args;
        if params.cArgs < 1 || params.rgvarg.is_null() { return S_OK; }
        let search_value = &*params.rgvarg;
        if let Ok(search) = IDispatch::try_from(search_value)
            && let Ok(tag) = win::string_property(&search, "Tag")
            && tag == (*this).expected_tag
        {
            (*this).completed.set(true);
        }
    }
    S_OK
}

static SINK_VTBL: SinkVtbl = SinkVtbl {
    query_interface, add_ref, release, get_type_info_count, get_type_info,
    get_ids_of_names, invoke: invoke_event,
};

struct Subscription {
    point: IConnectionPoint,
    cookie: u32,
    _sink: IUnknown,
    completed: Rc<Cell<bool>>,
}
impl Subscription {
    fn attach(application: &IDispatch, tag: &str) -> Result<Self> {
        let source: IConnectionPointContainer = application.cast()
            .context("Outlook Application doesn't expose connection points")?;
        let point = unsafe { source.FindConnectionPoint(&EVENT_IID) }
            .context("Outlook ApplicationEvents_11 connection point unavailable")?;
        let completed = Rc::new(Cell::new(false));
        let raw = Box::into_raw(Box::new(Sink {
            vtable: &SINK_VTBL, references: AtomicU32::new(1),
            expected_tag: tag.to_owned(), completed: completed.clone(),
        }));
        // SAFETY: Box begins with the exact IDispatch vtable pointer and is
        // created with a single COM reference owned by IUnknown. Advise takes
        // its own reference, dropped by Unadvise.
        let sink = unsafe { IUnknown::from_raw(raw.cast()) };
        let cookie = unsafe { point.Advise(&sink) }
            .context("Unable to subscribe to Outlook AdvancedSearchComplete")?;
        Ok(Self { point, cookie, _sink: sink, completed })
    }
}
impl Drop for Subscription {
    fn drop(&mut self) {
        // Always unadvise before releasing the client-owned sink. Never leave
        // a COM callback pointing at a released Rust object.
        let _ = unsafe { self.point.Unadvise(self.cookie) };
    }
}

pub fn filter(query: &str, scope: &str) -> String {
    let q = escape_dasl(query);
    let matching = |property: &str| format!("\"{property}\" ci_phrasematch '{q}'");
    let subject = matching("urn:schemas:httpmail:subject");
    let from_name = matching("urn:schemas:httpmail:fromname");
    let from_email = matching("urn:schemas:httpmail:fromemail");
    let body = matching("urn:schemas:httpmail:textdescription");
    match scope {
        "subject" => subject,
        "from" => format!("({from_name} OR {from_email})"),
        _ => format!("({subject} OR {body})"),
    }
}

pub fn search(
    application: &IDispatch, folder: &IDispatch, params: &SearchMailsParams, store_id: &str,
) -> Result<Value> {
    let started = Instant::now();
    let index = table::store_instant_search(folder);
    if index != Some(true) {
        bail!("Indexed AdvancedSearch requires an Outlook store with Instant Search enabled; requested store did not confirm indexing");
    }
    let limit = checked_limit(params.limit).map_err(anyhow::Error::msg)?;
    let scope_name = win::string_property(folder, "FolderPath")
        .context("Unable to read Outlook FolderPath for AdvancedSearch scope")?;
    // Microsoft specifies single-quoted folder paths. A literal apostrophe
    // must be doubled to remain part of the folder name.
    let scope = format!("'{}'", scope_name.replace('\'', "''"));
    let search_filter = filter(
        params.query.trim(), params.scope.as_deref().unwrap_or("subject_body")
    );
    let tag = format!(
        "outlook-mcp-rs-{}-{}",
        std::process::id(), NEXT_TAG.fetch_add(1, Ordering::Relaxed)
    );
    // The event sink must be advised BEFORE triggering AdvancedSearch. A
    // completion event may arrive immediately on cached/short queries.
    let subscription = Subscription::attach(application, &tag)?;
    let operation = win::method_object(application, "AdvancedSearch", vec![
        VARIANT::from(scope.as_str()),
        VARIANT::from(search_filter.as_str()),
        VARIANT::from(params.search_subfolders.unwrap_or(false)),
        VARIANT::from(tag.as_str()),
    ]).context("Application.AdvancedSearch failed")?;

    let synchronous = win::property(&operation, "IsSynchronous")
        .ok().and_then(|v| bool::try_from(&v).ok()).unwrap_or(false);
    if !synchronous {
        while !subscription.completed.get() {
            if started.elapsed() >= SEARCH_TIMEOUT {
                // Best-effort cancellation of *only this* search; do not
                // clear or interfere with another Outlook UI search.
                let _ = win::invoke(&operation, "Stop", DISPATCH_METHOD, vec![]);
                bail!("Outlook indexed search did not signal AdvancedSearchComplete within 18 seconds");
            }
            win::pump_messages();
            thread::sleep(Duration::from_millis(10));
        }
    }

    // Search.GetTable never triggers a new search and MUST be called only
    // once the associated operation is synchronous or has completed.
    let search_table = win::method_object(&operation, "GetTable", vec![])
        .context("Search.GetTable failed after completion")?;
    let page = table::enumerate_existing(
        &search_table, store_id,
        table::RowOptions {filter: "", limit, offset: 0, unread_only: false, sender_substring: None}
    )?;
    drop(subscription);
    Ok(json!({
        "query": params.query, "scope": params.scope.as_deref().unwrap_or("subject_body"),
        "backend": "advanced_search_indexed",
        "folder": win::string_property(folder, "Name").unwrap_or_default(),
        "search_subfolders": params.search_subfolders.unwrap_or(false),
        "instant_search_enabled": true,
        "count": page.items.len(), "has_more": page.has_more,
        "scanned": page.scanned, "scan_truncated": page.scan_truncated,
        "elapsed_ms": started.elapsed().as_millis(),
        "items": page.items
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn advanced_search_dasl_never_uses_at_sql_prefix() {
        assert_eq!(
            filter("O'Brien", "subject"),
            "\"urn:schemas:httpmail:subject\" ci_phrasematch 'O''Brien'"
        );
        assert!(!filter("hello", "subject_body").starts_with("@SQL="));
    }
    #[test]
    fn phrase_and_sender_search_have_explicit_scopes() {
        assert!(filter("world", "from").contains("urn:schemas:httpmail:fromemail"));
        assert!(filter("world", "subject_body").contains("urn:schemas:httpmail:textdescription"));
    }
    #[test]
    fn event_contract_constants_match_outlook_pia() {
        assert_eq!(ADVANCED_SEARCH_COMPLETE, 64106);
        assert_eq!(EVENT_IID, GUID::from_u128(0x0006302c_0000_0000_c000_000000000046));
    }
}
