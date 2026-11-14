//! JavaScript dates, and the number formats that make Excel display a datetime
//! as a date rather than as its serial number.

use std::cell::OnceCell;

use rust_xlsxwriter::{ExcelDateTime, Format, XlsxError};

use crate::format::RuscFormat;

/// Milliseconds in a day, the unit of an Excel serial datetime.
const MS_PER_DAY: f64 = 86_400_000.0;

/// Excel's dates, 1900-01-01T00:00:00Z up to but excluding
/// 10000-01-01T00:00:00Z, as milliseconds since the Unix epoch.
const EXCEL_TIME_RANGE: std::ops::Range<f64> = -2_208_988_800_000.0..253_402_300_800_000.0;

/// The Excel datetime of a JavaScript `Date`, from its milliseconds since the
/// Unix epoch (`Date.prototype.getTime()`), read as UTC.
pub(crate) fn from_js_time(time: f64) -> Result<ExcelDateTime, XlsxError> {
    if !EXCEL_TIME_RANGE.contains(&time) {
        return Err(XlsxError::DateTimeRangeError(
            "Date outside Excel's range of 1900-01-01 to 9999-12-31 (UTC)".to_string(),
        ));
    }
    let seconds = (time / 1000.0).floor();
    // `from_timestamp` checks the 1900-9999 range and handles Excel's 1900 leap
    // year; the milliseconds are then added to the serial value.
    let datetime = ExcelDateTime::from_timestamp(seconds as i64)?;
    let millis = time - seconds * 1000.0;
    if millis == 0.0 {
        return Ok(datetime);
    }
    ExcelDateTime::from_serial_datetime(datetime.to_excel() + millis / MS_PER_DAY)
}

/// What a datetime holds, which picks the number format it is shown with.
#[derive(Clone, Copy)]
enum DateTimeKind {
    Date,
    DateTime,
    Time,
}

impl DateTimeKind {
    fn of(datetime: &ExcelDateTime) -> Self {
        let serial = datetime.to_excel();
        if serial < 1.0 {
            Self::Time
        } else if serial.fract() == 0.0 {
            Self::Date
        } else {
            Self::DateTime
        }
    }

    /// Excel's built-in number format for a typed value of this kind, which
    /// Excel shows in the viewer's locale: `m/d/yyyy`, `m/d/yyyy h:mm` and
    /// `h:mm:ss` in en-US.
    fn num_format_index(self) -> u8 {
        match self {
            Self::Date => 14,
            Self::DateTime => 22,
            Self::Time => 21,
        }
    }

    fn slot(self) -> usize {
        self as usize
    }
}

/// The formats for the datetimes of one write call, built on first use.
///
/// A datetime needs a number format to display as a date. A cell format with
/// its own number format is used as it is; otherwise the datetime gets the cell
/// format with Excel's built-in number format for its kind.
pub(crate) struct DateTimeFormats<'a> {
    format: Option<&'a RuscFormat>,
    formats: [OnceCell<Format>; 3],
}

impl<'a> DateTimeFormats<'a> {
    pub(crate) fn new(format: Option<&'a RuscFormat>) -> Self {
        Self {
            format,
            formats: Default::default(),
        }
    }

    /// The format to write `datetime` with.
    pub(crate) fn get(&self, datetime: &ExcelDateTime) -> &Format {
        let kind = DateTimeKind::of(datetime);
        self.formats[kind.slot()].get_or_init(|| {
            let format = self.format.map(RuscFormat::to_format).unwrap_or_default();
            if self.format.is_some_and(RuscFormat::has_num_format) {
                format
            } else {
                format.set_num_format_index(kind.num_format_index())
            }
        })
    }
}
