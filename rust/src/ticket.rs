//! Lifecycle of one queued Outlook request, so a caller that gave up on a
//! queued request cannot have it execute later, and a request that was
//! already running is reported as an unknown outcome instead of a failure.
use std::sync::atomic::{AtomicU8, Ordering};

const QUEUED: u8 = 0;
const RUNNING: u8 = 1;
const DONE: u8 = 2;
const CANCELLED: u8 = 3;

/// Prefix on error strings that mean "may or may not have happened".
pub const UNKNOWN_OUTCOME: &str = "unknown_outcome: ";

/// Prefix on error strings that mean "provably did not happen".
pub const NOT_SENT: &str = "not_sent: ";

#[derive(Debug, Default)]
pub struct Ticket(AtomicU8);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Cancel {
    /// Never started and never will: the operation did not happen.
    NotExecuted,
    /// The worker is executing it right now.
    Running,
    /// It already finished.
    Completed,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeoutVerdict {
    /// Cancelled while queued; safe to report as failed.
    NotExecuted,
    /// A read is still running; it may finish later.
    ReadStillRunning,
    /// A write was in flight; verify in Outlook before any retry.
    UnknownOutcome,
    /// The result raced the timeout; fetch it.
    Completed,
}

impl Ticket {
    pub fn new() -> Self {
        Self(AtomicU8::new(QUEUED))
    }

    /// Worker side: true when it may execute. False means it was cancelled.
    pub fn start(&self) -> bool {
        self.0
            .compare_exchange(QUEUED, RUNNING, Ordering::AcqRel, Ordering::Acquire)
            .is_ok()
    }

    pub fn finish(&self) {
        self.0.store(DONE, Ordering::Release);
    }

    /// Caller side, after giving up. Cancels only a request that has not started.
    pub fn cancel(&self) -> Cancel {
        match self
            .0
            .compare_exchange(QUEUED, CANCELLED, Ordering::AcqRel, Ordering::Acquire)
        {
            Ok(_) => Cancel::NotExecuted,
            Err(RUNNING) => Cancel::Running,
            Err(_) => Cancel::Completed,
        }
    }
}

pub fn timeout_verdict(cancel: Cancel, idempotent: bool) -> TimeoutVerdict {
    match (cancel, idempotent) {
        (Cancel::NotExecuted, _) => TimeoutVerdict::NotExecuted,
        (Cancel::Completed, _) => TimeoutVerdict::Completed,
        (Cancel::Running, true) => TimeoutVerdict::ReadStillRunning,
        (Cancel::Running, false) => TimeoutVerdict::UnknownOutcome,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn queued_request_cancelled_by_timeout_never_runs() {
        let ticket = Ticket::new();
        assert_eq!(ticket.cancel(), Cancel::NotExecuted);
        // The worker reaches it later and must refuse to execute.
        assert!(!ticket.start());
    }

    #[test]
    fn running_request_cannot_be_cancelled_and_is_unknown_for_writes() {
        let ticket = Ticket::new();
        assert!(ticket.start());
        assert_eq!(ticket.cancel(), Cancel::Running);
        assert_eq!(
            timeout_verdict(Cancel::Running, false),
            TimeoutVerdict::UnknownOutcome
        );
        assert_eq!(
            timeout_verdict(Cancel::Running, true),
            TimeoutVerdict::ReadStillRunning
        );
    }

    #[test]
    fn finished_request_reports_completed_and_start_is_single_shot() {
        let ticket = Ticket::new();
        assert!(ticket.start());
        assert!(!ticket.start());
        ticket.finish();
        assert_eq!(ticket.cancel(), Cancel::Completed);
        assert_eq!(
            timeout_verdict(Cancel::Completed, false),
            TimeoutVerdict::Completed
        );
    }

    #[test]
    fn queued_cancellation_is_not_unknown_even_for_writes() {
        assert_eq!(
            timeout_verdict(Cancel::NotExecuted, false),
            TimeoutVerdict::NotExecuted
        );
    }
}
