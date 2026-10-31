//! Read-only Outlook Classic COM discovery through a dedicated STA worker.
//!
//! This is intentionally separate from the MCP tools until issue #4. No mail
//! mutation operations or live mailbox tools are exposed by this module.
use crate::contracts::{MailCommand, validate_command};
use anyhow::{Context, Result, bail};
use serde_json::Value;
#[cfg(windows)]
#[path = "outlook/indexed.rs"]
mod indexed;
#[cfg(windows)]
#[path = "outlook/mailbox.rs"]
mod mailbox;
#[cfg(windows)]
#[path = "outlook/refs.rs"]
mod refs;
#[cfg(windows)]
#[path = "outlook/table.rs"]
mod table;
#[cfg(windows)]
#[path = "outlook/thread.rs"]
mod thread;
use serde::Serialize;
use std::{
    sync::mpsc::{self, Receiver, Sender},
    thread as std_thread,
    time::Duration,
};

#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub struct OutlookProbe {
    pub current_user: String,
    pub inbox_name: String,
    pub inbox_items: i32,
    pub inbox_subfolders: i32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub first_item_subject: Option<String>,
}

enum Request {
    Probe {
        include_first_subject: bool,
        response: Sender<Result<OutlookProbe, String>>,
    },
    Call {
        command: MailCommand,
        response: Sender<Result<Value, String>>,
    },
    Shutdown,
}

/// Send plain Rust values across threads. COM interfaces live exclusively in
/// the worker, and never cross the channel into Tokio/MCP handler threads.
pub struct OutlookBridge {
    sender: Sender<Request>,
    // An intentionally detached thread on Drop avoids indefinitely blocking
    // shutdown if Outlook is displaying a modal COM/security dialog.
    _thread: std_thread::JoinHandle<()>,
}

impl OutlookBridge {
    pub fn start() -> Self {
        let (sender, receiver) = mpsc::channel();
        let handle = std_thread::Builder::new()
            .name("outlook-com-sta".into())
            .spawn(move || run_worker(receiver))
            .expect("unable to start Outlook COM worker");
        Self {
            sender,
            _thread: handle,
        }
    }

    /// Probe the currently running Outlook profile. The optional message
    /// subject is for explicit **local** diagnostics; it is never logged.
    pub fn probe(&self, include_first_subject: bool) -> Result<OutlookProbe> {
        let (tx, rx) = mpsc::channel();
        self.sender
            .send(Request::Probe {
                include_first_subject,
                response: tx,
            })
            .context("Outlook worker has exited")?;

        // Timeout is an unknown outcome (not cancellation of a COM call).
        // This operation is read-only; the worker may finish after timeout.
        match rx.recv_timeout(Duration::from_secs(45)) {
            Ok(Ok(value)) => Ok(value),
            Ok(Err(error)) => bail!("{error}"),
            Err(mpsc::RecvTimeoutError::Timeout) => bail!(
                "Outlook COM probe timed out. Check Outlook for a modal dialog; the underlying read may still be running."
            ),
            Err(mpsc::RecvTimeoutError::Disconnected) => bail!("Outlook COM worker disconnected"),
        }
    }
}

impl OutlookBridge {
    pub fn execute(&self, command: MailCommand) -> Result<Value> {
        validate_command(&command).map_err(anyhow::Error::msg)?;
        let (tx, rx) = mpsc::channel();
        self.sender
            .send(Request::Call {
                command,
                response: tx,
            })
            .context("Outlook COM worker has exited")?;
        match rx.recv_timeout(Duration::from_secs(45)) {
            Ok(Ok(value)) => Ok(value),
            Ok(Err(error)) => bail!("{error}"),
            Err(mpsc::RecvTimeoutError::Timeout) => bail!(
                "Read-only Outlook request timed out; Outlook may be showing a dialog. Request may finish later."
            ),
            Err(mpsc::RecvTimeoutError::Disconnected) => bail!("Outlook COM worker disconnected"),
        }
    }
}

impl Drop for OutlookBridge {
    fn drop(&mut self) {
        let _ = self.sender.send(Request::Shutdown);
    }
}

#[cfg(windows)]
fn run_worker(receiver: Receiver<Request>) {
    use std::sync::mpsc::RecvTimeoutError;
    let initialization = win::ComApartment::initialize();
    let mut session: Option<win::Session> = None;

    loop {
        let command = match receiver.recv_timeout(Duration::from_millis(20)) {
            Ok(command) => command,
            Err(RecvTimeoutError::Timeout) => {
                if initialization.is_ok() {
                    win::pump_messages();
                }
                continue;
            }
            Err(RecvTimeoutError::Disconnected) => break,
        };

        match command {
            Request::Shutdown => break,
            Request::Call { command, response } => {
                let result = match &initialization {
                    Ok(_) => {
                        if session.is_none() {
                            session = win::Session::attach().ok();
                        }
                        match &mut session {
                            Some(existing) => match existing.execute(command.clone()) {
                                Ok(value) => Ok(value),
                                Err(error) if win::is_disconnect(&error) => {
                                    session = None;
                                    match win::Session::attach() {
                                        Ok(mut refreshed) => {
                                            let result = refreshed.execute(command);
                                            session = Some(refreshed);
                                            result
                                        }
                                        Err(error) => Err(error),
                                    }
                                }
                                Err(error) => Err(error),
                            },
                            None => Err(anyhow::anyhow!(
                                "Outlook Classic is not running or its COM session is unavailable"
                            )),
                        }
                    }
                    Err(error) => Err(anyhow::anyhow!("{error}")),
                };
                let _ = response.send(result.map_err(|error| format!("{error:#}")));
            }
            Request::Probe {
                include_first_subject,
                response,
            } => {
                let result = match &initialization {
                    Ok(_) => {
                        if session.is_none() {
                            session = win::Session::attach().ok();
                        }
                        // An absent COM instance is an error, not a reason to
                        // auto-launch Outlook or fabricate mailbox results.
                        match &mut session {
                            Some(existing) => match existing.probe(include_first_subject) {
                                Ok(value) => Ok(value),
                                Err(error) if win::is_disconnect(&error) => {
                                    session = None;
                                    match win::Session::attach() {
                                        Ok(mut new_session) => {
                                            let result = new_session.probe(include_first_subject);
                                            session = Some(new_session);
                                            result
                                        }
                                        Err(error) => Err(error),
                                    }
                                }
                                Err(error) => Err(error),
                            },
                            None => Err(anyhow::anyhow!(
                                "Cannot attach to Outlook Classic. Open Outlook under this Windows account and ensure its profile is ready."
                            )),
                        }
                    }
                    Err(error) => Err(anyhow::anyhow!("{error}")),
                };
                let _ = response.send(result.map_err(|error| format!("{error:#}")));
            }
        }
    }
    // session drops before COM initialization guard when this scope ends.
    drop(session);
}

#[cfg(not(windows))]
fn run_worker(receiver: Receiver<Request>) {
    while let Ok(request) = receiver.recv() {
        match request {
            Request::Shutdown => break,
            Request::Call { response, .. } => {
                let _ = response.send(Err("Outlook Classic COM requires Windows".into()));
            }
            Request::Probe { response, .. } => {
                let _ = response.send(Err("Outlook Classic COM requires Windows".into()));
            }
        }
    }
}

#[cfg(windows)]
mod win {
    use super::OutlookProbe;
    use super::refs::MailRefs;
    use anyhow::{Context, Result};
    use std::ptr;
    use windows::{
        Win32::{
            System::{
                Com::{
                    CLSIDFromProgID, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
                    CoInitializeEx, CoUninitialize, DISPATCH_FLAGS, DISPATCH_METHOD,
                    DISPATCH_PROPERTYGET, DISPPARAMS, IDispatch,
                },
                Ole::GetActiveObject,
                Variant::VARIANT,
            },
            UI::WindowsAndMessaging::{
                DispatchMessageW, MSG, PM_REMOVE, PeekMessageW, TranslateMessage, WM_QUIT,
            },
        },
        core::{BSTR, GUID, IUnknown, Interface, PCWSTR},
    };

    pub struct ComApartment;
    impl ComApartment {
        pub fn initialize() -> Result<Self> {
            // SAFETY: This is the dedicated thread; balanced by Drop on
            // that same thread. Never move COM interfaces between threads.
            unsafe {
                CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)
                    .ok()
                    .context("CoInitializeEx(STA) failed")?;
            }
            Ok(Self)
        }
    }
    impl Drop for ComApartment {
        fn drop(&mut self) {
            // SAFETY: Runs on the same worker thread that initialized COM.
            unsafe {
                CoUninitialize();
            }
        }
    }

    pub fn pump_messages() {
        // SAFETY: The STA owns this thread's message queue. Only remove
        // pending messages here; do not block the worker waiting for them.
        unsafe {
            let mut message = MSG::default();
            while PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() {
                if message.message == WM_QUIT {
                    break;
                }
                let _ = TranslateMessage(&message);
                DispatchMessageW(&message);
            }
        }
    }

    pub fn is_disconnect(err: &anyhow::Error) -> bool {
        err.chain()
            .filter_map(|cause| cause.downcast_ref::<windows::core::Error>())
            .any(|error| {
                matches!(
                    error.code().0 as u32,
                    0x800706BA
                        | 0x800706BE
                        | 0x800706BF
                        | 0x80010108
                        | 0x80010007
                        | 0x80010012
                        | 0x800401FD
                )
            })
    }

    pub(super) fn invoke(
        object: &IDispatch,
        member: &str,
        flags: DISPATCH_FLAGS,
        mut args: Vec<VARIANT>,
    ) -> Result<VARIANT> {
        let name: Vec<u16> = member.encode_utf16().chain(std::iter::once(0)).collect();
        let names = [PCWSTR(name.as_ptr())];
        let iid_null = GUID::from_u128(0);
        let mut id = 0i32;
        // SAFETY: All pointers reference live stack data, and IDispatch is
        // invoked on the same STA that acquired it.
        unsafe {
            object
                .GetIDsOfNames(&iid_null, names.as_ptr(), 1, 0, &mut id)
                .with_context(|| format!("Outlook member lookup: {member}"))?;
            args.reverse(); // IDispatch positional arguments are reversed.
            let parameters = DISPPARAMS {
                rgvarg: if args.is_empty() {
                    ptr::null_mut()
                } else {
                    args.as_mut_ptr()
                },
                rgdispidNamedArgs: ptr::null_mut(),
                cArgs: args.len() as u32,
                cNamedArgs: 0,
            };
            let mut result = VARIANT::default();
            object
                .Invoke(
                    id,
                    &iid_null,
                    0,
                    flags,
                    &parameters,
                    Some(&mut result),
                    None,
                    None,
                )
                .with_context(|| format!("Outlook invoke: {member}"))?;
            Ok(result)
        }
    }

    pub(super) fn property(object: &IDispatch, member: &str) -> Result<VARIANT> {
        invoke(object, member, DISPATCH_PROPERTYGET, vec![])
    }

    pub(super) fn object_property(object: &IDispatch, name: &str) -> Result<IDispatch> {
        IDispatch::try_from(&property(object, name)?)
            .with_context(|| format!("Outlook {name} object conversion"))
    }

    pub(super) fn method_object(
        object: &IDispatch,
        name: &str,
        args: Vec<VARIANT>,
    ) -> Result<IDispatch> {
        IDispatch::try_from(&invoke(object, name, DISPATCH_METHOD, args)?)
            .with_context(|| format!("Outlook {name} returned non-object"))
    }

    pub(super) fn string_property(object: &IDispatch, name: &str) -> Result<String> {
        let value = property(object, name)?;
        let str_value = BSTR::try_from(&value)
            .with_context(|| format!("Outlook {name} returned non-string"))?;
        Ok(str_value.to_string())
    }

    pub(super) fn int_property(object: &IDispatch, name: &str) -> Result<i32> {
        let value = property(object, name)?;
        i32::try_from(&value).with_context(|| format!("Outlook {name} returned non-integer"))
    }

    pub struct Session {
        // Both COM pointers live and drop on this worker apartment.
        pub(super) _application: IDispatch,
        pub(super) namespace: IDispatch,
        pub(super) refs: MailRefs,
    }

    impl Session {
        pub fn attach() -> Result<Self> {
            // Outlook is expected to be already running; intentionally
            // avoid CoCreateInstance, which would start OUTLOOK.EXE.
            let clsid = unsafe { CLSIDFromProgID(windows::core::w!("Outlook.Application")) }
                .context("Outlook Classic is not registered")?;
            let mut unknown: Option<IUnknown> = None;
            unsafe { GetActiveObject(&clsid, None, &mut unknown) }
                .context("Outlook Classic is not running or is not registered in the COM Running Object Table")?;
            let application: IDispatch =
                unknown
                    .context("GetActiveObject returned null")?
                    .cast()
                    .context("Active Outlook instance does not expose IDispatch")?;
            let namespace =
                method_object(&application, "GetNamespace", vec![VARIANT::from("MAPI")])?;
            Ok(Self {
                _application: application,
                namespace,
                refs: MailRefs::new(),
            })
        }

        pub fn probe(&mut self, first_subject: bool) -> Result<OutlookProbe> {
            let current = object_property(&self.namespace, "CurrentUser")?;
            let current_user = string_property(&current, "Name")?;
            let inbox = method_object(
                &self.namespace,
                "GetDefaultFolder",
                vec![VARIANT::from(6i32)],
            )?;
            let inbox_name = string_property(&inbox, "Name")?;
            let folders = object_property(&inbox, "Folders")?;
            let inbox_subfolders = int_property(&folders, "Count")?;
            let items = object_property(&inbox, "Items")?;
            let inbox_items = int_property(&items, "Count")?;
            let first_item_subject = if first_subject && inbox_items > 0 {
                let first = method_object(&items, "Item", vec![VARIANT::from(1i32)])?;
                Some(string_property(&first, "Subject")?)
            } else {
                None
            };
            Ok(OutlookProbe {
                current_user,
                inbox_name,
                inbox_items,
                inbox_subfolders,
                first_item_subject,
            })
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn serialized_probe_is_bounded_and_omits_optional_subject() {
        let result = OutlookProbe {
            current_user: "User".into(),
            inbox_name: "Inbox".into(),
            inbox_items: 100,
            inbox_subfolders: 3,
            first_item_subject: None,
        };
        let json = serde_json::to_string(&result).unwrap();
        assert!(!json.contains("first_item_subject"));
        assert!(json.len() < 160);
    }
    #[cfg(not(windows))]
    #[test]
    fn platform_returns_explicit_error() {
        let bridge = OutlookBridge::start();
        let error = bridge.probe(false).unwrap_err().to_string();
        assert!(error.contains("requires Windows"));
    }
}
