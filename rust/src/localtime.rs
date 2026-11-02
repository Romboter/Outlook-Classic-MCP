//! Minimal local wall-clock date/time for calendar queries.
//!
//! Outlook stores local wall-clock times, so this deliberately has no time
//! zone: ISO strings with an offset or `Z` are rejected rather than guessed.

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LocalTime {
    pub year: i32,
    pub month: u32,
    pub day: u32,
    pub hour: u32,
    pub minute: u32,
    pub second: u32,
}

fn is_leap(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        _ if is_leap(year) => 29,
        _ => 28,
    }
}

// Howard Hinnant's civil-date algorithms (days since 1970-01-01).
fn days_from_civil(y: i32, m: u32, d: u32) -> i64 {
    let y = i64::from(if m <= 2 { y - 1 } else { y });
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = i64::from((m + 9) % 12);
    let doy = (153 * mp + 2) / 5 + i64::from(d) - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn civil_from_days(z: i64) -> (i32, u32, u32) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    let y = (yoe + era * 400) as i32 + i32::from(m <= 2);
    (y, m, d)
}

fn digits(s: &str, min: usize, max: usize) -> Option<u32> {
    if !(min..=max).contains(&s.len()) || !s.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    s.parse().ok()
}

impl LocalTime {
    /// Accepts `YYYY-MM-DD`, optionally followed by `T` or a space and
    /// `HH:MM[:SS[.fraction]]`. Offsets and `Z` are rejected.
    pub fn parse(input: &str) -> Result<Self, String> {
        let bad = || {
            format!(
                "invalid datetime '{input}'. Use local ISO-8601 like 2026-04-25T14:30:00 (no time zone offset)"
            )
        };
        let text = input.trim();
        let (date, time) = match text.find(['T', ' ']) {
            Some(i) => (&text[..i], Some(&text[i + 1..])),
            None => (text, None),
        };
        let mut d = date.split('-');
        let (Some(y), Some(m), Some(day), None) = (d.next(), d.next(), d.next(), d.next()) else {
            return Err(bad());
        };
        let (Some(year), Some(month), Some(day)) =
            (digits(y, 4, 4), digits(m, 2, 2), digits(day, 2, 2))
        else {
            return Err(bad());
        };
        let (mut hour, mut minute, mut second) = (0, 0, 0);
        if let Some(time) = time {
            let time = time.split('.').next().unwrap_or("");
            let mut t = time.split(':');
            let (Some(h), Some(mi)) = (t.next(), t.next()) else {
                return Err(bad());
            };
            let (Some(h), Some(mi)) = (digits(h, 2, 2), digits(mi, 2, 2)) else {
                return Err(bad());
            };
            hour = h;
            minute = mi;
            if let Some(s) = t.next() {
                second = digits(s, 2, 2).ok_or_else(bad)?;
            }
            if t.next().is_some() {
                return Err(bad());
            }
        }
        let year = year as i32;
        if !(1..=12).contains(&month)
            || day < 1
            || day > days_in_month(year, month)
            || hour > 23
            || minute > 59
            || second > 59
        {
            return Err(bad());
        }
        Ok(Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
        })
    }

    pub fn from_parts(
        year: i32,
        month: u32,
        day: u32,
        hour: u32,
        minute: u32,
        second: u32,
    ) -> Self {
        Self {
            year,
            month,
            day,
            hour,
            minute,
            second,
        }
    }

    pub fn add_days(self, days: i64) -> Self {
        let (year, month, day) =
            civil_from_days(days_from_civil(self.year, self.month, self.day) + days);
        Self {
            year,
            month,
            day,
            ..self
        }
    }

    /// Minutes since 1970-01-01 00:00 on the same wall clock (ordering and spans).
    pub fn minutes(&self) -> i64 {
        days_from_civil(self.year, self.month, self.day) * 1440
            + i64::from(self.hour) * 60
            + i64::from(self.minute)
    }

    pub fn iso(&self) -> String {
        format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            self.year, self.month, self.day, self.hour, self.minute, self.second
        )
    }

    /// Outlook Restrict date literal: 12-hour clock with AM/PM.
    pub fn jet(&self) -> String {
        let (h12, ampm) = match self.hour {
            0 => (12, "AM"),
            1..=11 => (self.hour, "AM"),
            12 => (12, "PM"),
            h => (h - 12, "PM"),
        };
        format!(
            "{:02}/{:02}/{:04} {:02}:{:02} {}",
            self.month, self.day, self.year, h12, self.minute, ampm
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_date_and_time_forms() {
        let t = LocalTime::parse("2026-04-25T14:30:05").unwrap();
        assert_eq!(t.iso(), "2026-04-25T14:30:05");
        assert_eq!(
            LocalTime::parse("2026-04-25").unwrap().iso(),
            "2026-04-25T00:00:00"
        );
        assert_eq!(LocalTime::parse("2026-04-25 09:05").unwrap().minute, 5);
        assert_eq!(
            LocalTime::parse("2026-04-25T14:30:05.123").unwrap().second,
            5
        );
    }

    #[test]
    fn rejects_offsets_and_invalid_values() {
        for bad in [
            "2026-04-25T14:30:00Z",
            "2026-04-25T14:30:00+02:00",
            "2026-04-25T14:30:00-05:00",
            "2026-02-30",
            "2026-13-01",
            "2026-04-25T24:00",
            "26-04-25",
            "garbage",
            "",
        ] {
            assert!(LocalTime::parse(bad).is_err(), "{bad}");
        }
        assert!(LocalTime::parse("2028-02-29").is_ok());
        assert!(LocalTime::parse("2027-02-29").is_err());
    }

    #[test]
    fn adds_days_across_month_year_and_leap_day() {
        let t = LocalTime::parse("2026-12-25T08:00").unwrap();
        assert_eq!(t.add_days(14).iso(), "2027-01-08T08:00:00");
        assert_eq!(LocalTime::parse("2028-02-28").unwrap().add_days(1).day, 29);
        assert_eq!(LocalTime::parse("2027-02-28").unwrap().add_days(1).month, 3);
        assert_eq!(t.add_days(-25).iso(), "2026-11-30T08:00:00");
    }

    #[test]
    fn minutes_orders_and_spans() {
        let a = LocalTime::parse("2026-01-01T00:00").unwrap();
        let b = LocalTime::parse("2027-01-02T00:00").unwrap();
        assert_eq!(b.minutes() - a.minutes(), 366 * 1440);
    }

    #[test]
    fn jet_uses_twelve_hour_clock() {
        let jet = |h, m| LocalTime::from_parts(2026, 4, 5, h, m, 0).jet();
        assert_eq!(jet(0, 30), "04/05/2026 12:30 AM");
        assert_eq!(jet(9, 5), "04/05/2026 09:05 AM");
        assert_eq!(jet(12, 0), "04/05/2026 12:00 PM");
        assert_eq!(jet(14, 30), "04/05/2026 02:30 PM");
    }
}
