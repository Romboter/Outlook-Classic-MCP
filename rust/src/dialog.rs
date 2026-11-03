//! Spike: native approval window (see docs/DIALOG-APPROVAL-DESIGN.md). Not wired into any tool.
//! Fails closed: anything other than a click on Approve is not an approval.

use std::time::Duration;

use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::Controls::{
    TASKDIALOG_BUTTON, TASKDIALOG_NOTIFICATIONS, TASKDIALOGCONFIG, TDE_FOOTER,
    TDF_ALLOW_DIALOG_CANCELLATION, TDF_CALLBACK_TIMER, TDF_EXPAND_FOOTER_AREA, TDF_SIZE_TO_CONTENT,
    TDM_CLICK_BUTTON, TDM_SET_ELEMENT_TEXT, TDN_CREATED, TDN_TIMER, TaskDialogIndirect,
};
use windows::Win32::UI::WindowsAndMessaging::{
    HWND_TOPMOST, SWP_NOMOVE, SWP_NOSIZE, SendMessageW, SetForegroundWindow, SetWindowPos,
};
use windows::core::{HRESULT, PCWSTR};

const APPROVE: i32 = 1001;
const DECLINE: i32 = 1002;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Decision {
    Approved,
    /// Declined, closed, or any failure to show the window.
    Declined,
    TimedOut,
}

pub struct Request {
    pub title: String,
    /// One-line statement of what will happen.
    pub headline: String,
    /// Short facts always visible.
    pub summary: String,
    /// Long text (full attendee or file list), shown in the expandable area.
    pub details: String,
    pub timeout: Duration,
}

struct CallbackData {
    timeout_ms: u32,
    timed_out: std::sync::atomic::AtomicBool,
}

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

unsafe extern "system" fn callback(
    hwnd: HWND,
    msg: TASKDIALOG_NOTIFICATIONS,
    wparam: WPARAM,
    _lparam: LPARAM,
    data: isize,
) -> HRESULT {
    let data = unsafe { &*(data as *const CallbackData) };
    if msg == TDN_CREATED {
        // Visible even if Outlook or the client is in front.
        unsafe {
            let _ = SetWindowPos(
                hwnd,
                Some(HWND_TOPMOST),
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE,
            );
            let _ = SetForegroundWindow(hwnd);
        }
    } else if msg == TDN_TIMER {
        let elapsed = wparam.0 as u32;
        if elapsed >= data.timeout_ms {
            data.timed_out
                .store(true, std::sync::atomic::Ordering::SeqCst);
            unsafe {
                SendMessageW(
                    hwnd,
                    TDM_CLICK_BUTTON.0 as u32,
                    Some(WPARAM(DECLINE as usize)),
                    None,
                )
            };
        } else {
            let left = (data.timeout_ms - elapsed).div_ceil(1000);
            let text = wide(&format!("Declines automatically in {left} s"));
            unsafe {
                SendMessageW(
                    hwnd,
                    TDM_SET_ELEMENT_TEXT.0 as u32,
                    Some(WPARAM(TDE_FOOTER.0 as usize)),
                    Some(LPARAM(text.as_ptr() as isize)),
                )
            };
        }
    }
    HRESULT(0)
}

/// Shows the window on its own thread and waits. Never call from the COM worker.
pub fn approve(request: Request) -> Decision {
    std::thread::spawn(move || show(&request))
        .join()
        .unwrap_or(Decision::Declined)
}

fn show(request: &Request) -> Decision {
    let title = wide(&request.title);
    let headline = wide(&request.headline);
    let summary = wide(&request.summary);
    let details = wide(&request.details);
    let approve_text = wide("Approve");
    let decline_text = wide("Decline");
    let footer = wide("Declines automatically");
    let expand = wide("Hide full details");
    let collapse = wide("Show full details");
    let buttons = [
        TASKDIALOG_BUTTON {
            nButtonID: APPROVE,
            pszButtonText: PCWSTR(approve_text.as_ptr()),
        },
        TASKDIALOG_BUTTON {
            nButtonID: DECLINE,
            pszButtonText: PCWSTR(decline_text.as_ptr()),
        },
    ];
    let data = CallbackData {
        timeout_ms: request.timeout.as_millis().min(u32::MAX as u128) as u32,
        timed_out: std::sync::atomic::AtomicBool::new(false),
    };
    let mut config = TASKDIALOGCONFIG {
        cbSize: std::mem::size_of::<TASKDIALOGCONFIG>() as u32,
        dwFlags: TDF_ALLOW_DIALOG_CANCELLATION
            | TDF_CALLBACK_TIMER
            | TDF_EXPAND_FOOTER_AREA
            | TDF_SIZE_TO_CONTENT,
        pszWindowTitle: PCWSTR(title.as_ptr()),
        pszMainInstruction: PCWSTR(headline.as_ptr()),
        pszContent: PCWSTR(summary.as_ptr()),
        pszExpandedInformation: PCWSTR(details.as_ptr()),
        pszExpandedControlText: PCWSTR(expand.as_ptr()),
        pszCollapsedControlText: PCWSTR(collapse.as_ptr()),
        pszFooter: PCWSTR(footer.as_ptr()),
        cButtons: buttons.len() as u32,
        pButtons: buttons.as_ptr(),
        nDefaultButton: DECLINE,
        pfCallback: Some(callback),
        lpCallbackData: &data as *const CallbackData as isize,
        ..Default::default()
    };
    config.cxWidth = 0;
    let mut pressed = 0i32;
    if unsafe { TaskDialogIndirect(&config, Some(&mut pressed), None, None) }.is_err() {
        return Decision::Declined;
    }
    if data.timed_out.load(std::sync::atomic::Ordering::SeqCst) {
        Decision::TimedOut
    } else if pressed == APPROVE {
        Decision::Approved
    } else {
        Decision::Declined
    }
}
