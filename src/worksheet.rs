use napi::bindgen_prelude::*;
use napi_derive::napi;
use parking_lot::Mutex;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatAverage,
    ConditionalFormatAverageRule, ConditionalFormatBlank, ConditionalFormatCell,
    ConditionalFormatCellRule, ConditionalFormatDataBar, ConditionalFormatDate,
    ConditionalFormatDateRule, ConditionalFormatDuplicate, ConditionalFormatError,
    ConditionalFormatFormula, ConditionalFormatIconSet, ConditionalFormatIconType,
    ConditionalFormatText, ConditionalFormatTextRule, ConditionalFormatTop,
    ConditionalFormatTopRule, ExcelDateTime, Format, Url, Workbook,
};
use std::sync::Arc;

use crate::button::RuscButton;
use crate::chart::RuscChart;
use crate::datetime::{DateTimeFormats, from_js_time};
use crate::format::RuscFormat;
use crate::formula::{RuscFormula, RuscFormulaData, to_formula};
use crate::image::RuscImage;
use crate::note::RuscNote;
use crate::protection::RuscProtectionOptions;
use crate::richtext::{RuscRichText, rich_string_segments};
use crate::shape::RuscShape;
use crate::sparkline::RuscSparkline;
use crate::table::RuscTable;
use crate::utils::{parse_color, parse_ignore_error, parse_rgb_color};
use crate::validation::RuscDataValidation;

/// Worksheet wrapper for Excel sheet manipulation
#[napi]
pub struct RuscWorksheet {
    pub(crate) workbook: Arc<Mutex<Workbook>>,
    pub(crate) index: usize,
}

/// A value written to one cell.
enum CellValue {
    String(String),
    Number(f64),
    Boolean(bool),
    /// A JavaScript `Date`, as milliseconds since the Unix epoch.
    DateTime(f64),
}

impl CellValue {
    /// Reads a JavaScript cell value; `null` and `undefined` are no value, as
    /// `None` is for `rust_xlsxwriter`'s `Option` cell data.
    fn from_js(value: &Unknown<'_>) -> Result<Option<Self>> {
        // SAFETY: each `cast` follows the `get_type()` check for its type.
        let cell_value = match value.get_type()? {
            ValueType::String => Self::String(unsafe { value.cast()? }),
            ValueType::Number => Self::Number(unsafe { value.cast()? }),
            ValueType::Boolean => Self::Boolean(unsafe { value.cast()? }),
            ValueType::Object if value.is_date()? => {
                // SAFETY: `is_date()` established that this is a JavaScript Date.
                let time = unsafe { value.cast::<napi::JsDate>()? }.value_of()?;
                if time.is_nan() {
                    return Err(Error::new(
                        Status::InvalidArg,
                        "Unsupported cell value: Invalid Date",
                    ));
                }
                Self::DateTime(time)
            }
            ValueType::Null | ValueType::Undefined => return Ok(None),
            value_type => return Err(unsupported_cell_value(value, value_type)?),
        };
        Ok(Some(cell_value))
    }
}

/// The error for a JavaScript value that cannot be written to a cell.
fn unsupported_cell_value(value: &Unknown<'_>, value_type: ValueType) -> Result<Error> {
    let kind = match value_type {
        ValueType::Object if value.is_array()? => "array",
        ValueType::Object => "object",
        ValueType::Function => "function",
        ValueType::Symbol => "symbol",
        ValueType::BigInt => "bigint",
        _ => "unknown",
    };
    Ok(Error::new(
        Status::InvalidArg,
        format!("Unsupported cell value type: {kind}. Use a string, number, boolean, Date or null"),
    ))
}

/// Converts the result of a fluent `rust_xlsxwriter` call into a JavaScript
/// result, naming the failed action in the error.
trait OrJsError {
    fn or_js_error(self, action: &str) -> Result<()>;
}

impl<T> OrJsError for std::result::Result<T, rust_xlsxwriter::XlsxError> {
    fn or_js_error(self, action: &str) -> Result<()> {
        self.map(drop)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to {action}: {e}")))
    }
}

/// Writes `value` like `rust_xlsxwriter` writes `Option` cell data: no value
/// leaves the cell alone, or writes a blank cell when there is a format. A date
/// is written with its format from `dates`, so that it displays as a date.
fn write_cell(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    value: Option<CellValue>,
    format: Option<&Format>,
    dates: &DateTimeFormats<'_>,
) -> std::result::Result<(), rust_xlsxwriter::XlsxError> {
    match (value, format) {
        (Some(CellValue::DateTime(time)), _) => {
            let datetime = from_js_time(time)?;
            worksheet.write_datetime_with_format(row, col, &datetime, dates.get(&datetime))
        }
        (Some(CellValue::String(s)), None) => worksheet.write_string(row, col, &s),
        (Some(CellValue::Number(n)), None) => worksheet.write_number(row, col, n),
        (Some(CellValue::Boolean(b)), None) => worksheet.write_boolean(row, col, b),
        (Some(CellValue::String(s)), Some(f)) => {
            worksheet.write_string_with_format(row, col, &s, f)
        }
        (Some(CellValue::Number(n)), Some(f)) => worksheet.write_number_with_format(row, col, n, f),
        (Some(CellValue::Boolean(b)), Some(f)) => {
            worksheet.write_boolean_with_format(row, col, b, f)
        }
        (None, Some(f)) => worksheet.write_blank(row, col, f),
        (None, None) => return Ok(()),
    }
    .map(|_| ())
}

/// Writes `values` to consecutive cells of `row` the way `write` writes one
/// value, so missing values leave a cell empty or write a formatted blank.
fn write_array_values(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    start_col: u16,
    values: &Array<'_>,
    format: Option<&Format>,
    dates: &DateTimeFormats<'_>,
) -> Result<()> {
    for index in 0..values.len() {
        let column = start_col + index as u16;
        let value = match values.get::<Unknown<'_>>(index)? {
            Some(value) => CellValue::from_js(&value).map_err(|error| {
                Error::new(
                    error.status,
                    format!("Invalid value at row {row}, col {column}: {}", error.reason),
                )
            })?,
            None => None,
        };

        write_cell(worksheet, row, column, value, format, dates).map_err(|error| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write at row {row}, col {column}: {error}"),
            )
        })?;
    }

    Ok(())
}

#[napi]
impl RuscWorksheet {
    /// Write a value to a cell (auto-detects type). `null` or `undefined`
    /// writes nothing, or a blank cell with the format when one is given. A
    /// `Date` is written as a UTC datetime; without a number format of its own it
    /// gets Excel's built-in date or date-time format.
    #[napi]
    pub fn write(
        &self,
        row: u32,
        col: u16,
        #[napi(ts_arg_type = "string | number | boolean | Date | null | undefined")] value: Unknown,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let value = CellValue::from_js(&value)?;
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let dates = DateTimeFormats::new(format);
        let format = format.map(RuscFormat::to_format);
        write_cell(ws, row, col, value, format.as_ref(), &dates)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write: {}", e)))
    }

    /// Write an array of values as a row starting at the given column.
    /// Uses lazy Array iteration — reads one cell at a time from the JS array
    /// instead of deserializing the entire row into Rust memory.
    /// Values are written as `write` writes them: `null`, `undefined` and
    /// holes leave the cell empty, or write a blank cell with the format.
    #[napi(strict)]
    pub fn write_row(
        &self,
        row: u32,
        start_col: u16,
        values: Array,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;
        let dates = DateTimeFormats::new(format);
        let format = format.map(RuscFormat::to_format);

        write_array_values(ws, row, start_col, &values, format.as_ref(), &dates)
    }

    /// Write a 2D array of values starting at the given position.
    /// Uses lazy Array iteration — reads one row/cell at a time from JS,
    /// never materializing the entire grid in Rust memory.
    /// Values are written as `write` writes them: `null`, `undefined` and
    /// holes leave the cell empty, or write a blank cell with the format.
    #[napi(strict)]
    pub fn write_rows(
        &self,
        start_row: u32,
        start_col: u16,
        rows: Array,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;
        let dates = DateTimeFormats::new(format);
        let format = format.map(RuscFormat::to_format);

        for row_offset in 0..rows.len() {
            let row = start_row + row_offset;
            let values = match rows.get::<Unknown<'_>>(row_offset)? {
                // SAFETY: `is_array()` established that this is a JavaScript array.
                Some(values) if values.is_array()? => unsafe { values.cast::<Array>()? },
                _ => {
                    return Err(Error::new(
                        Status::InvalidArg,
                        format!(
                            "Invalid row at index {row_offset}: expected an array of cell values"
                        ),
                    ));
                }
            };
            write_array_values(ws, row, start_col, &values, format.as_ref(), &dates)?;
        }

        Ok(())
    }

    /// Write a string to a cell
    #[napi]
    pub fn write_string(&self, row: u32, col: u16, value: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.write_string(row, col, &value).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write string: {}", e),
            )
        })?;
        Ok(())
    }

    /// Write a number to a cell
    #[napi]
    pub fn write_number(&self, row: u32, col: u16, value: f64) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.write_number(row, col, value).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write number: {}", e),
            )
        })?;
        Ok(())
    }

    /// Write a boolean to a cell
    #[napi]
    pub fn write_boolean(&self, row: u32, col: u16, value: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.write_boolean(row, col, value).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write boolean: {}", e),
            )
        })?;
        Ok(())
    }

    /// Write a blank cell with formatting
    #[napi]
    pub fn write_blank(&self, row: u32, col: u16, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.write_blank(row, col, &format.to_format()).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write blank cell: {}", e),
            )
        })?;
        Ok(())
    }

    /// Write a rich text string (multi-format text) to a cell.
    ///
    /// `richText` is `[text, format]` segments, or a deprecated `RuscRichText`; a
    /// segment without a format uses the default cell font. Segments are plain data,
    /// so nothing outlives the call. Empty segments are skipped, and a rich string
    /// whose segments are all empty leaves the cell untouched.
    #[napi]
    pub fn write_rich_string(
        &self,
        row: u32,
        col: u16,
        rich_text: Either<&RuscRichText, Vec<(String, Option<&RuscFormat>)>>,
    ) -> Result<()> {
        self.write_rich_text(row, col, &rich_text, None)
    }

    /// Write a rich text string with an additional cell format.
    ///
    /// `richText` is as for `writeRichString`. A rich string whose segments are all
    /// empty writes a blank cell with the cell format.
    #[napi]
    pub fn write_rich_string_with_format(
        &self,
        row: u32,
        col: u16,
        rich_text: Either<&RuscRichText, Vec<(String, Option<&RuscFormat>)>>,
        cell_format: &RuscFormat,
    ) -> Result<()> {
        self.write_rich_text(row, col, &rich_text, Some(cell_format))
    }

    /// Write a datetime from a Unix timestamp (seconds since 1970-01-01 UTC)
    #[napi]
    pub fn write_datetime_from_timestamp(
        &self,
        row: u32,
        col: u16,
        timestamp: f64,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let datetime = ExcelDateTime::from_timestamp(timestamp as i64).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to create datetime from timestamp: {}", e),
            )
        })?;

        ws.write_datetime_with_format(
            row,
            col,
            &datetime,
            DateTimeFormats::new(format).get(&datetime),
        )
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write datetime: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a datetime from year, month, day components
    #[napi]
    pub fn write_datetime_from_ymd(
        &self,
        row: u32,
        col: u16,
        year: u16,
        month: u8,
        day: u8,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let datetime = ExcelDateTime::from_ymd(year, month, day).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to create datetime from ymd: {}", e),
            )
        })?;

        ws.write_datetime_with_format(
            row,
            col,
            &datetime,
            DateTimeFormats::new(format).get(&datetime),
        )
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write datetime: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a datetime from year, month, day, hour, minute, second components
    #[napi]
    // The arguments mirror the positional parameters of the public JS API.
    #[allow(clippy::too_many_arguments)]
    pub fn write_datetime_from_ymd_hms(
        &self,
        row: u32,
        col: u16,
        year: u16,
        month: u8,
        day: u8,
        hour: u16,
        min: u8,
        sec: f64,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let datetime = ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to create date: {}", e),
                )
            })?
            .and_hms(hour, min, sec)
            .map_err(|e| {
                Error::new(Status::GenericFailure, format!("Failed to add time: {}", e))
            })?;

        ws.write_datetime_with_format(
            row,
            col,
            &datetime,
            DateTimeFormats::new(format).get(&datetime),
        )
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write datetime: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a datetime from an ISO 8601 string (e.g., "2023-01-25T12:30:00")
    #[napi]
    pub fn write_datetime_from_string(
        &self,
        row: u32,
        col: u16,
        datetime_str: String,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let datetime = ExcelDateTime::parse_from_str(&datetime_str).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to parse datetime string: {}", e),
            )
        })?;

        ws.write_datetime_with_format(
            row,
            col,
            &datetime,
            DateTimeFormats::new(format).get(&datetime),
        )
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write datetime: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a value with formatting to a cell; `null` or `undefined` writes
    /// a blank cell with the format
    #[napi]
    pub fn write_with_format(
        &self,
        row: u32,
        col: u16,
        #[napi(ts_arg_type = "string | number | boolean | Date | null | undefined")] value: Unknown,
        format: &RuscFormat,
    ) -> Result<()> {
        self.write(row, col, value, Some(format))
    }

    /// Set the width of a column
    #[napi]
    pub fn set_column_width(&self, col: u16, width: f64) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_column_width(col, width).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set column width: {}", e),
            )
        })?;
        Ok(())
    }

    /// Set the height of a row
    #[napi]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_row_height(row, height).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set row height: {}", e),
            )
        })?;
        Ok(())
    }

    /// Set the default height for all worksheet rows efficiently
    #[napi]
    pub fn set_default_row_height(&self, height: f64) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_default_row_height(height);
        Ok(())
    }
    /// Set the width of a column in pixels
    #[napi]
    pub fn set_column_width_pixels(&self, col: u16, width: u32) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_width_pixels(col, width)
                .or_js_error("set column width")
        })
    }

    /// Set the width of a range of columns in character units
    #[napi]
    pub fn set_column_range_width(&self, first_col: u16, last_col: u16, width: f64) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_range_width(first_col, last_col, width)
                .or_js_error("set column range width")
        })
    }

    /// Set the width of a range of columns in pixels
    #[napi]
    pub fn set_column_range_width_pixels(
        &self,
        first_col: u16,
        last_col: u16,
        width: u32,
    ) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_range_width_pixels(first_col, last_col, width)
                .or_js_error("set column range width")
        })
    }

    /// Set the width of a column in pixels in auto-fit mode: as with a column
    /// Excel auto-fitted, Excel widens it when a number in it gets more digits
    #[napi]
    pub fn set_column_autofit_width(&self, col: u16, width: u32) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_autofit_width(col, width)
                .or_js_error("set column autofit width")
        })
    }

    /// Set the default format of a column, used by its cells that have no format
    #[napi]
    pub fn set_column_format(&self, col: u16, format: &RuscFormat) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_format(col, &format.to_format())
                .or_js_error("set column format")
        })
    }

    /// Set the default format of a range of columns
    #[napi]
    pub fn set_column_range_format(
        &self,
        first_col: u16,
        last_col: u16,
        format: &RuscFormat,
    ) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_range_format(first_col, last_col, &format.to_format())
                .or_js_error("set column range format")
        })
    }

    /// Hide a column
    #[napi]
    pub fn set_column_hidden(&self, col: u16) -> Result<()> {
        self.with_worksheet(|ws| ws.set_column_hidden(col).or_js_error("hide column"))
    }

    /// Hide a range of columns
    #[napi]
    pub fn set_column_range_hidden(&self, first_col: u16, last_col: u16) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_column_range_hidden(first_col, last_col)
                .or_js_error("hide column range")
        })
    }

    /// Set the height of a row in pixels
    #[napi]
    pub fn set_row_height_pixels(&self, row: u32, height: u32) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_row_height_pixels(row, height)
                .or_js_error("set row height")
        })
    }

    /// Set the default height of all worksheet rows in pixels
    #[napi]
    pub fn set_default_row_height_pixels(&self, height: u32) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_default_row_height_pixels(height);
            Ok(())
        })
    }

    /// Set the default format of a row, used by its cells that have no format
    #[napi]
    pub fn set_row_format(&self, row: u32, format: &RuscFormat) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_row_format(row, &format.to_format())
                .or_js_error("set row format")
        })
    }

    /// Hide a row
    #[napi]
    pub fn set_row_hidden(&self, row: u32) -> Result<()> {
        self.with_worksheet(|ws| ws.set_row_hidden(row).or_js_error("hide row"))
    }

    /// Unhide a row hidden with setRowHidden()
    #[napi]
    pub fn set_row_unhidden(&self, row: u32) -> Result<()> {
        self.with_worksheet(|ws| ws.set_row_unhidden(row).or_js_error("unhide row"))
    }

    /// Hide all rows that have no data, format or row settings, with a single
    /// worksheet setting instead of one per row
    #[napi]
    pub fn hide_unused_rows(&self, enable: bool) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.hide_unused_rows(enable);
            Ok(())
        })
    }

    /// Group a range of rows into an outline level; a range inside an existing
    /// group adds a nested level
    #[napi]
    pub fn group_rows(&self, first_row: u32, last_row: u32) -> Result<()> {
        self.with_worksheet(|ws| ws.group_rows(first_row, last_row).or_js_error("group rows"))
    }

    /// Group a range of rows into a collapsed outline level
    #[napi]
    pub fn group_rows_collapsed(&self, first_row: u32, last_row: u32) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.group_rows_collapsed(first_row, last_row)
                .or_js_error("group rows")
        })
    }

    /// Group a range of columns into an outline level; a range inside an existing
    /// group adds a nested level
    #[napi]
    pub fn group_columns(&self, first_col: u16, last_col: u16) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.group_columns(first_col, last_col)
                .or_js_error("group columns")
        })
    }

    /// Group a range of columns into a collapsed outline level
    #[napi]
    pub fn group_columns_collapsed(&self, first_col: u16, last_col: u16) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.group_columns_collapsed(first_col, last_col)
                .or_js_error("group columns")
        })
    }

    /// Show the row outline expand/collapse symbols above the groups instead of below
    #[napi]
    pub fn group_symbols_above(&self, enable: bool) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.group_symbols_above(enable);
            Ok(())
        })
    }

    /// Show the column outline expand/collapse symbols to the left of the groups
    /// instead of to the right
    #[napi]
    pub fn group_symbols_to_left(&self, enable: bool) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.group_symbols_to_left(enable);
            Ok(())
        })
    }

    /// Set the format of a cell, keeping its data, or of an empty cell
    #[napi]
    pub fn set_cell_format(&self, row: u32, col: u16, format: &RuscFormat) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_cell_format(row, col, &format.to_format())
                .or_js_error("set cell format")
        })
    }

    /// Set the format of every cell in a range, keeping their data
    #[napi]
    pub fn set_range_format(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        format: &RuscFormat,
    ) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_range_format(
                first_row,
                first_col,
                last_row,
                last_col,
                &format.to_format(),
            )
            .or_js_error("set range format")
        })
    }

    /// Set the format of every cell in a range, keeping their data, and draw the
    /// border of `borderFormat` around the range
    #[napi]
    pub fn set_range_format_with_border(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        cell_format: &RuscFormat,
        border_format: &RuscFormat,
    ) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.set_range_format_with_border(
                first_row,
                first_col,
                last_row,
                last_col,
                &cell_format.to_format(),
                &border_format.to_format(),
            )
            .or_js_error("set range format")
        })
    }

    /// Remove the data and format of a cell
    #[napi]
    pub fn clear_cell(&self, row: u32, col: u16) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.clear_cell(row, col);
            Ok(())
        })
    }

    /// Remove the format of a cell, keeping its data
    #[napi]
    pub fn clear_cell_format(&self, row: u32, col: u16) -> Result<()> {
        self.with_worksheet(|ws| {
            ws.clear_cell_format(row, col);
            Ok(())
        })
    }

    /// Turn off an Excel warning, such as "number stored as text", for a cell.
    /// Excel allows one rule per cell or range.
    #[napi]
    pub fn ignore_error(&self, row: u32, col: u16, error_type: String) -> Result<()> {
        let error_type = parse_ignore_error(&error_type)?;
        self.with_worksheet(|ws| {
            ws.ignore_error(row, col, error_type)
                .or_js_error("ignore error")
        })
    }

    /// Turn off an Excel warning for a range of cells. Excel allows one rule per
    /// cell or range.
    #[napi]
    pub fn ignore_error_range(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        error_type: String,
    ) -> Result<()> {
        let error_type = parse_ignore_error(&error_type)?;
        self.with_worksheet(|ws| {
            ws.ignore_error_range(first_row, first_col, last_row, last_col, error_type)
                .or_js_error("ignore error")
        })
    }

    /// Set the exclusive upper row bound inspected by autofit.
    #[napi]
    pub fn set_autofit_max_row(&self, max_row: u32) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_autofit_max_row(max_row);
        Ok(())
    }

    /// Limit the autofitted column width in pixels.
    #[napi]
    pub fn set_autofit_max_width(&self, max_width: u32) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_autofit_max_width(max_width);
        Ok(())
    }

    /// Autofit worksheet columns based on their cell contents.
    #[napi]
    pub fn autofit(&self) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.autofit();
        Ok(())
    }

    /// Set the name of the worksheet
    #[napi]
    pub fn set_name(&self, name: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_name(&name).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set worksheet name: {}", e),
            )
        })?;
        Ok(())
    }

    /// Merge a range of cells
    #[napi]
    pub fn merge_range(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        value: String,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.merge_range(first_row, first_col, last_row, last_col, &value, &f)
        } else {
            ws.merge_range(
                first_row,
                first_col,
                last_row,
                last_col,
                &value,
                &Format::new(),
            )
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to merge range: {}", e),
            )
        })?;

        Ok(())
    }

    /// Merge a range of cells with formatting
    #[napi]
    pub fn merge_range_with_format(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        value: String,
        format: &RuscFormat,
    ) -> Result<()> {
        self.merge_range(
            first_row,
            first_col,
            last_row,
            last_col,
            value,
            Some(format),
        )
    }

    /// Insert a note (post-it style comment) to a cell
    #[napi]
    pub fn insert_note(&self, row: u32, col: u16, note: &RuscNote) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_note(row, col, note.get_note()).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to insert note: {}", e),
            )
        })?;
        Ok(())
    }

    /// Set the default author name for all notes in the worksheet
    #[napi]
    pub fn set_default_note_author(&self, name: String) -> Result<()> {
        if name.chars().count() > 52 {
            return Err(Error::new(
                Status::GenericFailure,
                "Author name must be less than 52 characters".to_string(),
            ));
        }

        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_default_note_author(&name);
        Ok(())
    }

    /// Write a formula to a cell
    #[napi]
    pub fn write_formula(
        &self,
        row: u32,
        col: u16,
        formula: Either<&RuscFormula, RuscFormulaData>,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let formula = to_formula(formula);
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_formula_with_format(row, col, formula, &f)
        } else {
            ws.write_formula(row, col, formula)
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write formula: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a dynamic array formula to a cell (Excel 365)
    #[napi]
    pub fn write_dynamic_formula(
        &self,
        row: u32,
        col: u16,
        formula: Either<&RuscFormula, RuscFormulaData>,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let formula = to_formula(formula);
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_dynamic_array_formula_with_format(row, col, row, col, formula, &f)
        } else {
            ws.write_dynamic_array_formula(row, col, row, col, formula)
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write dynamic formula: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write an array formula to a range of cells
    #[napi]
    pub fn write_array_formula(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        formula: Either<&RuscFormula, RuscFormulaData>,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let formula = to_formula(formula);
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_array_formula_with_format(
                first_row, first_col, last_row, last_col, formula, &f,
            )
        } else {
            ws.write_array_formula(first_row, first_col, last_row, last_col, formula)
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write array formula: {}", e),
            )
        })?;

        Ok(())
    }

    /// Add data validation to a cell or range
    #[napi]
    pub fn add_data_validation(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        validation: &RuscDataValidation,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.add_data_validation(
            first_row,
            first_col,
            last_row,
            last_col,
            validation.get_validation(),
        )
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to add data validation: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a URL to a cell
    #[napi]
    pub fn write_url(
        &self,
        row: u32,
        col: u16,
        url: String,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let url_obj = Url::new(&url);

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_url_with_format(row, col, &url_obj, &f)
        } else {
            ws.write_url(row, col, &url_obj)
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write URL: {}", e),
            )
        })?;

        Ok(())
    }

    /// Write a URL to a cell with custom display text
    #[napi]
    pub fn write_url_with_text(
        &self,
        row: u32,
        col: u16,
        url: String,
        text: String,
        format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let url_obj = Url::new(&url).set_text(&text);

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_url_with_format(row, col, &url_obj, &f)
        } else {
            ws.write_url(row, col, &url_obj)
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write URL: {}", e),
            )
        })?;

        Ok(())
    }

    /// Set autofilter on a range
    #[napi]
    pub fn autofilter(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.autofilter(first_row, first_col, last_row, last_col)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to set autofilter: {}", e),
                )
            })?;

        Ok(())
    }

    /// Freeze panes at the specified row and column
    #[napi]
    pub fn set_freeze_panes(&self, row: u32, col: u16) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_freeze_panes(row, col).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set freeze panes: {}", e),
            )
        })?;

        Ok(())
    }

    /// Add 2-color scale conditional formatting
    #[napi]
    pub fn add_conditional_format_2color_scale(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        min_color: String,
        max_color: String,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let min = parse_rgb_color("min color", &min_color)?;
        let max = parse_rgb_color("max color", &max_color)?;

        let format = ConditionalFormat2ColorScale::new()
            .set_minimum_color(min)
            .set_maximum_color(max);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add 2-color scale: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add 3-color scale conditional formatting
    #[napi]
    // The arguments mirror the positional parameters of the public JS API.
    #[allow(clippy::too_many_arguments)]
    pub fn add_conditional_format_3color_scale(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        min_color: String,
        mid_color: String,
        max_color: String,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let min = parse_rgb_color("min color", &min_color)?;
        let mid = parse_rgb_color("mid color", &mid_color)?;
        let max = parse_rgb_color("max color", &max_color)?;

        let format = ConditionalFormat3ColorScale::new()
            .set_minimum_color(min)
            .set_midpoint_color(mid)
            .set_maximum_color(max);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add 3-color scale: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add data bar conditional formatting
    #[napi]
    pub fn add_conditional_format_data_bar(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        bar_color: String,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let color = parse_rgb_color("bar color", &bar_color)?;

        let format = ConditionalFormatDataBar::new().set_fill_color(color);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add data bar: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add cell-based conditional formatting (greater than, less than, etc.)
    #[napi]
    // The arguments mirror the positional parameters of the public JS API.
    #[allow(clippy::too_many_arguments)]
    pub fn add_conditional_format_cell(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        rule: String,
        value: f64,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let cell_rule = match rule.as_str() {
            "greaterThan" => ConditionalFormatCellRule::GreaterThan(value),
            "lessThan" => ConditionalFormatCellRule::LessThan(value),
            "greaterThanOrEqual" => ConditionalFormatCellRule::GreaterThanOrEqualTo(value),
            "lessThanOrEqual" => ConditionalFormatCellRule::LessThanOrEqualTo(value),
            "equal" => ConditionalFormatCellRule::EqualTo(value),
            "notEqual" => ConditionalFormatCellRule::NotEqualTo(value),
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid cell rule: {}. Valid values: 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'equal', 'notEqual'",
                        rule
                    ),
                ));
            }
        };

        let cond_format = ConditionalFormatCell::new()
            .set_rule(cell_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add cell conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add text-based conditional formatting
    #[napi]
    // The arguments mirror the positional parameters of the public JS API.
    #[allow(clippy::too_many_arguments)]
    pub fn add_conditional_format_text(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        rule: String,
        text: String,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let text_rule = match rule.as_str() {
            "contains" => ConditionalFormatTextRule::Contains(text),
            "notContains" => ConditionalFormatTextRule::DoesNotContain(text),
            "beginsWith" => ConditionalFormatTextRule::BeginsWith(text),
            "endsWith" => ConditionalFormatTextRule::EndsWith(text),
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid text rule: {}. Valid values: 'contains', 'notContains', 'beginsWith', 'endsWith'",
                        rule
                    ),
                ));
            }
        };

        let cond_format = ConditionalFormatText::new()
            .set_rule(text_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add text conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add average/standard deviation conditional formatting
    #[napi]
    pub fn add_conditional_format_average(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        rule: String,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let avg_rule = match rule.as_str() {
            "aboveAverage" => ConditionalFormatAverageRule::AboveAverage,
            "belowAverage" => ConditionalFormatAverageRule::BelowAverage,
            "equalOrAboveAverage" => ConditionalFormatAverageRule::EqualOrAboveAverage,
            "equalOrBelowAverage" => ConditionalFormatAverageRule::EqualOrBelowAverage,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid average rule: {}. Valid values: 'aboveAverage', 'belowAverage', 'equalOrAboveAverage', 'equalOrBelowAverage'",
                        rule
                    ),
                ));
            }
        };

        let cond_format = ConditionalFormatAverage::new()
            .set_rule(avg_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add average conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add top/bottom N conditional formatting
    #[napi]
    // The arguments mirror the positional parameters of the public JS API.
    #[allow(clippy::too_many_arguments)]
    pub fn add_conditional_format_top(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        rule: String,
        value: u16,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let top_rule = match rule.as_str() {
            "top" => ConditionalFormatTopRule::Top(value),
            "bottom" => ConditionalFormatTopRule::Bottom(value),
            "topPercent" => ConditionalFormatTopRule::TopPercent(value),
            "bottomPercent" => ConditionalFormatTopRule::BottomPercent(value),
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid top/bottom rule: {}. Valid values: 'top', 'bottom', 'topPercent', 'bottomPercent'",
                        rule
                    ),
                ));
            }
        };

        let cond_format = ConditionalFormatTop::new()
            .set_rule(top_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add top/bottom conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add duplicate/unique values conditional formatting
    #[napi]
    pub fn add_conditional_format_duplicate(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        invert: bool,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let mut cond_format = ConditionalFormatDuplicate::new().set_format(format.to_format());

        if invert {
            cond_format = cond_format.invert();
        }

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add duplicate conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add blank/non-blank conditional formatting
    #[napi]
    pub fn add_conditional_format_blank(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        invert: bool,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let mut cond_format = ConditionalFormatBlank::new().set_format(format.to_format());

        if invert {
            cond_format = cond_format.invert();
        }

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add blank conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add error/non-error conditional formatting
    #[napi]
    pub fn add_conditional_format_error(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        invert: bool,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let mut cond_format = ConditionalFormatError::new().set_format(format.to_format());

        if invert {
            cond_format = cond_format.invert();
        }

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add error conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add date-based conditional formatting
    #[napi]
    pub fn add_conditional_format_date(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        rule: String,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let date_rule = match rule.as_str() {
            "yesterday" => ConditionalFormatDateRule::Yesterday,
            "today" => ConditionalFormatDateRule::Today,
            "tomorrow" => ConditionalFormatDateRule::Tomorrow,
            "last7Days" => ConditionalFormatDateRule::Last7Days,
            "lastWeek" => ConditionalFormatDateRule::LastWeek,
            "thisWeek" => ConditionalFormatDateRule::ThisWeek,
            "nextWeek" => ConditionalFormatDateRule::NextWeek,
            "lastMonth" => ConditionalFormatDateRule::LastMonth,
            "thisMonth" => ConditionalFormatDateRule::ThisMonth,
            "nextMonth" => ConditionalFormatDateRule::NextMonth,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid date rule: {}. Valid values: 'yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek', 'lastMonth', 'thisMonth', 'nextMonth'",
                        rule
                    ),
                ));
            }
        };

        let cond_format = ConditionalFormatDate::new()
            .set_rule(date_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add date conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add formula-based conditional formatting
    #[napi]
    pub fn add_conditional_format_formula(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        formula: String,
        format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let cond_format = ConditionalFormatFormula::new()
            .set_rule(&*formula)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add formula conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add icon set conditional formatting
    #[napi]
    pub fn add_conditional_format_icon_set(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        icon_type: String,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let icon_set_type = match icon_type.as_str() {
            "threeArrows" => ConditionalFormatIconType::ThreeArrows,
            "threeArrowsGray" => ConditionalFormatIconType::ThreeArrowsGray,
            "threeFlags" => ConditionalFormatIconType::ThreeFlags,
            "threeTrafficLights" => ConditionalFormatIconType::ThreeTrafficLights,
            "threeTrafficLightsWithRim" => ConditionalFormatIconType::ThreeTrafficLightsWithRim,
            "threeSigns" => ConditionalFormatIconType::ThreeSigns,
            "threeSymbols" => ConditionalFormatIconType::ThreeSymbols,
            "threeSymbolsCircled" => ConditionalFormatIconType::ThreeSymbolsCircled,
            "threeStars" => ConditionalFormatIconType::ThreeStars,
            "threeTriangles" => ConditionalFormatIconType::ThreeTriangles,
            "fourArrows" => ConditionalFormatIconType::FourArrows,
            "fourArrowsGray" => ConditionalFormatIconType::FourArrowsGray,
            "fourRedToBlack" => ConditionalFormatIconType::FourRedToBlack,
            "fourHistograms" => ConditionalFormatIconType::FourHistograms,
            "fourTrafficLights" => ConditionalFormatIconType::FourTrafficLights,
            "fiveArrows" => ConditionalFormatIconType::FiveArrows,
            "fiveArrowsGray" => ConditionalFormatIconType::FiveArrowsGray,
            "fiveHistograms" => ConditionalFormatIconType::FiveHistograms,
            "fiveQuadrants" => ConditionalFormatIconType::FiveQuadrants,
            "fiveBoxes" => ConditionalFormatIconType::FiveBoxes,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid icon set type: {}. Valid values: 'threeArrows', 'threeArrowsGray', 'threeFlags', 'threeTrafficLights', 'threeTrafficLightsWithRim', 'threeSigns', 'threeSymbols', 'threeSymbolsCircled', 'threeStars', 'threeTriangles', 'fourArrows', 'fourArrowsGray', 'fourRedToBlack', 'fourHistograms', 'fourTrafficLights', 'fiveArrows', 'fiveArrowsGray', 'fiveHistograms', 'fiveQuadrants', 'fiveBoxes'",
                        icon_type
                    ),
                ));
            }
        };

        let cond_format = ConditionalFormatIconSet::new().set_icon_type(icon_set_type);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add icon set conditional format: {}", e),
                )
            })?;

        Ok(())
    }

    /// Set page margins (in inches)
    #[napi]
    pub fn set_margins(
        &self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_margins(left, right, top, bottom, header, footer);
        Ok(())
    }

    /// Set page orientation to portrait
    #[napi]
    pub fn set_portrait(&self) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_portrait();
        Ok(())
    }

    /// Set page orientation to landscape
    #[napi]
    pub fn set_landscape(&self) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_landscape();
        Ok(())
    }

    /// Set paper size (1=Letter, 9=A4, etc.)
    #[napi]
    pub fn set_paper_size(&self, paper_size: u8) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_paper_size(paper_size);
        Ok(())
    }

    /// Set header text
    #[napi]
    pub fn set_header(&self, header: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_header(&header);
        Ok(())
    }

    /// Set footer text
    #[napi]
    pub fn set_footer(&self, footer: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_footer(&footer);
        Ok(())
    }

    /// Set print area
    #[napi]
    pub fn set_print_area(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_area(first_row, first_col, last_row, last_col)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to set print area: {}", e),
                )
            })?;
        Ok(())
    }

    /// Set print gridlines
    #[napi]
    pub fn set_print_gridlines(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_gridlines(enable);
        Ok(())
    }

    /// Show or hide the gridlines of the worksheet on screen
    #[napi]
    pub fn set_screen_gridlines(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_screen_gridlines(enable);
        Ok(())
    }

    /// Center content horizontally on the page when printing
    #[napi]
    pub fn set_print_center_horizontally(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_center_horizontally(enable);
        Ok(())
    }

    /// Center content vertically on the page when printing
    #[napi]
    pub fn set_print_center_vertically(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_center_vertically(enable);
        Ok(())
    }

    /// Fit worksheet to N pages wide by M pages tall when printing
    #[napi]
    pub fn set_print_fit_to_pages(&self, width: u16, height: u16) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_fit_to_pages(width, height);
        Ok(())
    }

    /// Set print scale as a percentage (10-400)
    #[napi]
    pub fn set_print_scale(&self, scale: u16) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_scale(scale);
        Ok(())
    }

    /// Set the first page number for printing
    #[napi]
    pub fn set_print_first_page_number(&self, page_number: u16) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_first_page_number(page_number);
        Ok(())
    }

    /// Enable/disable black and white printing
    #[napi]
    pub fn set_print_black_and_white(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_black_and_white(enable);
        Ok(())
    }

    /// Enable/disable draft quality printing
    #[napi]
    pub fn set_print_draft(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_draft(enable);
        Ok(())
    }

    /// Enable/disable printing of row and column headings
    #[napi]
    pub fn set_print_headings(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_print_headings(enable);
        Ok(())
    }

    /// Repeat rows at the top of each printed page
    #[napi]
    pub fn set_repeat_rows(&self, first_row: u32, last_row: u32) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_repeat_rows(first_row, last_row).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set repeat rows: {}", e),
            )
        })?;
        Ok(())
    }

    /// Repeat columns at the left of each printed page
    #[napi]
    pub fn set_repeat_columns(&self, first_col: u16, last_col: u16) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_repeat_columns(first_col, last_col).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set repeat columns: {}", e),
            )
        })?;
        Ok(())
    }

    /// Scale header and footer with document
    #[napi]
    pub fn set_header_footer_scale_with_doc(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_header_footer_scale_with_doc(enable);
        Ok(())
    }

    /// Align header and footer with page margins
    #[napi]
    pub fn set_header_footer_align_with_page(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_header_footer_align_with_page(enable);
        Ok(())
    }

    /// Set the worksheet tab color
    #[napi]
    pub fn set_tab_color(&self, color: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_tab_color(parse_color(&color)?);
        Ok(())
    }

    /// Hide or show the worksheet
    #[napi]
    pub fn set_hidden(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_hidden(enable);
        Ok(())
    }

    /// Set the worksheet zoom level (10-400)
    #[napi]
    pub fn set_zoom(&self, zoom: u16) -> Result<()> {
        if !(10..=400).contains(&zoom) {
            return Err(Error::new(
                Status::GenericFailure,
                format!("Zoom must be between 10 and 400, got {}", zoom),
            ));
        }

        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_zoom(zoom);
        Ok(())
    }

    /// Hide zero values in worksheet cells, the equivalent of turning off the
    /// Excel option "Show a zero in cells that have zero value"
    #[napi]
    pub fn set_hidden_zeroes(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_hidden_zeroes(enable);
        Ok(())
    }

    /// Set the default cell selection when opening the worksheet
    #[napi]
    pub fn set_selection(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_selection(first_row, first_col, last_row, last_col)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to set selection: {}", e),
                )
            })?;
        Ok(())
    }

    /// Set which cell appears in the top-left corner when opening the worksheet
    #[napi]
    pub fn set_top_left_cell(&self, row: u32, col: u16) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.set_top_left_cell(row, col).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to set top left cell: {}", e),
            )
        })?;
        Ok(())
    }

    /// Protect the worksheet from editing
    #[napi]
    pub fn protect(&self) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.protect();
        Ok(())
    }

    /// Protect the worksheet with a password
    #[napi]
    pub fn protect_with_password(&self, password: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.protect_with_password(&password);
        Ok(())
    }

    /// Protect the worksheet with granular protection options
    #[napi]
    pub fn protect_with_options(&self, options: &RuscProtectionOptions) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.protect_with_options(options.get_options());
        Ok(())
    }

    /// Protect the worksheet with a password and granular protection options
    #[napi]
    pub fn protect_with_password_and_options(
        &self,
        password: String,
        options: &RuscProtectionOptions,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.protect_with_password(&password);
        ws.protect_with_options(options.get_options());
        Ok(())
    }

    /// Insert a chart at the specified cell position
    #[napi]
    pub fn insert_chart(&self, row: u32, col: u16, chart: &RuscChart) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        // Create a clone of the chart since rust_xlsxwriter takes by reference
        // Note: This is a workaround. In the future, we might need to handle chart lifecycle differently
        ws.insert_chart(row, col, &chart.chart).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to insert chart: {}", e),
            )
        })?;

        Ok(())
    }

    /// Insert a chart at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_chart_with_offset(
        &self,
        row: u32,
        col: u16,
        chart: &RuscChart,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_chart_with_offset(row, col, &chart.chart, x_offset, y_offset)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert chart with offset: {}", e),
                )
            })?;

        Ok(())
    }

    /// Insert an image at the specified cell position
    #[napi]
    pub fn insert_image(&self, row: u32, col: u16, image: &RuscImage) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_image(row, col, image.get_image()).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to insert image: {}", e),
            )
        })?;

        Ok(())
    }

    /// Insert an image at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_image_with_offset(
        &self,
        row: u32,
        col: u16,
        image: &RuscImage,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_image_with_offset(row, col, image.get_image(), x_offset, y_offset)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert image with offset: {}", e),
                )
            })?;

        Ok(())
    }

    /// Insert an image that fits to a cell
    #[napi]
    pub fn insert_image_fit_to_cell(
        &self,
        row: u32,
        col: u16,
        image: &RuscImage,
        keep_aspect_ratio: bool,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_image_fit_to_cell(row, col, image.get_image(), keep_aspect_ratio)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert fitted image: {}", e),
                )
            })?;

        Ok(())
    }
    /// Insert an image scaled proportionally and centered within a cell.
    #[napi]
    pub fn insert_image_fit_to_cell_centered(
        &self,
        row: u32,
        col: u16,
        image: &RuscImage,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_image_fit_to_cell_centered(row, col, image.get_image())
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert centered fitted image: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add a table to the worksheet
    #[napi]
    pub fn add_table(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        table: &RuscTable,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.add_table(first_row, first_col, last_row, last_col, table.get_table())
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add table: {}", e),
                )
            })?;

        Ok(())
    }

    /// Add a sparkline to a single cell
    #[napi]
    pub fn add_sparkline(&self, row: u32, col: u16, sparkline: &RuscSparkline) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let sparkline = sparkline.for_worksheet(&ws.name());
        ws.add_sparkline(row, col, &sparkline).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to add sparkline: {}", e),
            )
        })?;

        Ok(())
    }

    /// Add a sparkline group to a range of cells
    #[napi]
    pub fn add_sparkline_group(
        &self,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        sparkline: &RuscSparkline,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let sparkline = sparkline.for_worksheet(&ws.name());
        ws.add_sparkline_group(first_row, first_col, last_row, last_col, &sparkline)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to add sparkline group: {}", e),
                )
            })?;

        Ok(())
    }

    /// Insert a button at the specified cell position
    #[napi]
    pub fn insert_button(&self, row: u32, col: u16, button: &RuscButton) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_button(row, col, button.get_button())
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert button: {}", e),
                )
            })?;

        Ok(())
    }

    /// Insert a button at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_button_with_offset(
        &self,
        row: u32,
        col: u16,
        button: &RuscButton,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_button_with_offset(row, col, button.get_button(), x_offset, y_offset)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert button with offset: {}", e),
                )
            })?;

        Ok(())
    }

    /// Insert a shape (textbox) at the specified cell position
    #[napi]
    pub fn insert_shape(&self, row: u32, col: u16, shape: &RuscShape) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_shape(row, col, shape.get_shape()).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to insert shape: {}", e),
            )
        })?;

        Ok(())
    }

    /// Insert a shape (textbox) at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_shape_with_offset(
        &self,
        row: u32,
        col: u16,
        shape: &RuscShape,
        x_offset: u32,
        y_offset: u32,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        ws.insert_shape_with_offset(row, col, shape.get_shape(), x_offset, y_offset)
            .map_err(|e| {
                Error::new(
                    Status::GenericFailure,
                    format!("Failed to insert shape with offset: {}", e),
                )
            })?;

        Ok(())
    }
}

impl RuscWorksheet {
    /// Writes a rich string, with the cell format when one is given.
    fn write_rich_text(
        &self,
        row: u32,
        col: u16,
        rich_text: &Either<&RuscRichText, Vec<(String, Option<&RuscFormat>)>>,
        cell_format: Option<&RuscFormat>,
    ) -> Result<()> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;

        let segments = rich_string_segments(rich_text);
        let cell_format = cell_format.map(RuscFormat::to_format);

        if segments.is_empty() {
            if let Some(format) = &cell_format {
                ws.write_blank(row, col, format).map_err(|e| {
                    Error::new(
                        Status::GenericFailure,
                        format!("Failed to write blank cell: {}", e),
                    )
                })?;
            }
            return Ok(());
        }

        let rich_string: Vec<(&Format, &str)> = segments
            .iter()
            .map(|(format, text)| (format, *text))
            .collect();
        match &cell_format {
            Some(format) => ws.write_rich_string_with_format(row, col, &rich_string, format),
            None => ws.write_rich_string(row, col, &rich_string),
        }
        .map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to write rich string: {}", e),
            )
        })?;

        Ok(())
    }

    /// Runs `f` on this worksheet while holding the workbook lock.
    fn with_worksheet<T>(
        &self,
        f: impl FnOnce(&mut rust_xlsxwriter::Worksheet) -> Result<T>,
    ) -> Result<T> {
        let mut wb = self.workbook.lock();
        let ws = wb.worksheet_from_index(self.index).map_err(|e| {
            Error::new(
                Status::GenericFailure,
                format!("Failed to get worksheet: {}", e),
            )
        })?;
        f(ws)
    }
}
