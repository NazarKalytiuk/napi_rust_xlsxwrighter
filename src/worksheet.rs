use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{
    ConditionalFormat2ColorScale, ConditionalFormat3ColorScale,
    ConditionalFormatAverage, ConditionalFormatAverageRule, ConditionalFormatBlank,
    ConditionalFormatCell, ConditionalFormatCellRule, ConditionalFormatDataBar,
    ConditionalFormatDate, ConditionalFormatDateRule, ConditionalFormatDuplicate,
    ConditionalFormatError, ConditionalFormatFormula, ConditionalFormatIconSet,
    ConditionalFormatIconType, ConditionalFormatText, ConditionalFormatTextRule,
    ConditionalFormatTop, ConditionalFormatTopRule, ExcelDateTime, Format, Url, Workbook,
};
use std::sync::{Arc, Mutex};

use crate::button::RuscButton;
use crate::chart::RuscChart;
use crate::format::RuscFormat;
use crate::formula::RuscFormula;
use crate::image::RuscImage;
use crate::note::RuscNote;
use crate::protection::RuscProtectionOptions;
use crate::richtext::RuscRichText;
use crate::shape::RuscShape;
use crate::sparkline::RuscSparkline;
use crate::table::RuscTable;
use crate::utils::parse_color;
use crate::validation::RuscDataValidation;


/// Worksheet wrapper for Excel sheet manipulation
#[napi]
pub struct RuscWorksheet {
    pub(crate) workbook: Arc<Mutex<Workbook>>,
    pub(crate) index: usize,
}

#[napi]
impl RuscWorksheet {
    /// Write a value to a cell (auto-detects type)
    #[napi]
    pub fn write(&self, row: u32, col: u16, value: Either3<String, f64, bool>, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            match value {
                Either3::A(s) => ws.write_string_with_format(row, col, &s, &f),
                Either3::B(n) => ws.write_number_with_format(row, col, n, &f),
                Either3::C(b) => ws.write_boolean_with_format(row, col, b, &f),
            }
        } else {
            match value {
                Either3::A(s) => ws.write_string(row, col, &s),
                Either3::B(n) => ws.write_number(row, col, n),
                Either3::C(b) => ws.write_boolean(row, col, b),
            }
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write: {}", e)))?;

        Ok(())
    }

    /// Write a string to a cell
    #[napi]
    pub fn write_string(&self, row: u32, col: u16, value: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.write_string(row, col, &value)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write string: {}", e)))?;
        Ok(())
    }

    /// Write a number to a cell
    #[napi]
    pub fn write_number(&self, row: u32, col: u16, value: f64) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.write_number(row, col, value)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write number: {}", e)))?;
        Ok(())
    }

    /// Write a boolean to a cell
    #[napi]
    pub fn write_boolean(&self, row: u32, col: u16, value: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.write_boolean(row, col, value)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write boolean: {}", e)))?;
        Ok(())
    }

    /// Write a blank cell with formatting
    #[napi]
    pub fn write_blank(&self, row: u32, col: u16, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.write_blank(row, col, &format.to_format())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write blank cell: {}", e)))?;
        Ok(())
    }

    /// Write a rich text string (multi-format text) to a cell
    #[napi]
    pub fn write_rich_string(&self, row: u32, col: u16, rich_text: &RuscRichText) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        // Build the segments array for rust_xlsxwriter, filtering out empty strings
        let segments = rich_text.get_segments();
        let mut xlsxwriter_segments: Vec<(Format, String)> = Vec::new();
        for (text, format) in segments {
            if !text.is_empty() {
                let fmt = format.to_format();
                xlsxwriter_segments.push((fmt, text.clone()));
            }
        }

        // If all segments were blank, silently skip (no-op)
        if xlsxwriter_segments.is_empty() {
            return Ok(());
        }

        // Convert to slice of references as required by rust_xlsxwriter
        let segment_refs: Vec<(&Format, &str)> = xlsxwriter_segments
            .iter()
            .map(|(fmt, text)| (fmt, text.as_str()))
            .collect();

        ws.write_rich_string(row, col, &segment_refs)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write rich string: {}", e)))?;

        Ok(())
    }

    /// Write a rich text string with an additional cell format
    #[napi]
    pub fn write_rich_string_with_format(
        &self,
        row: u32,
        col: u16,
        rich_text: &RuscRichText,
        cell_format: &RuscFormat,
    ) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        // Build the segments array, filtering out empty strings
        let segments = rich_text.get_segments();
        let mut xlsxwriter_segments: Vec<(Format, String)> = Vec::new();
        for (text, format) in segments {
            if !text.is_empty() {
                let fmt = format.to_format();
                xlsxwriter_segments.push((fmt, text.clone()));
            }
        }

        // If all segments were blank, write a blank cell with the cell format instead
        if xlsxwriter_segments.is_empty() {
            let format = cell_format.to_format();
            ws.write_blank(row, col, &format)
                .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write blank cell: {}", e)))?;
            return Ok(());
        }

        // Convert to slice of references
        let segment_refs: Vec<(&Format, &str)> = xlsxwriter_segments
            .iter()
            .map(|(fmt, text)| (fmt, text.as_str()))
            .collect();

        let format = cell_format.to_format();
        ws.write_rich_string_with_format(row, col, &segment_refs, &format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write rich string with format: {}", e)))?;

        Ok(())
    }

    /// Write a datetime from a Unix timestamp (seconds since 1970-01-01 UTC)
    #[napi]
    pub fn write_datetime_from_timestamp(&self, row: u32, col: u16, timestamp: f64, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let datetime = ExcelDateTime::from_timestamp(timestamp as i64)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to create datetime from timestamp: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_datetime_with_format(row, col, &datetime, &f)
        } else {
            ws.write_datetime(row, col, &datetime)
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write datetime: {}", e)))?;

        Ok(())
    }

    /// Write a datetime from year, month, day components
    #[napi]
    pub fn write_datetime_from_ymd(&self, row: u32, col: u16, year: u16, month: u8, day: u8, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let datetime = ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to create datetime from ymd: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_datetime_with_format(row, col, &datetime, &f)
        } else {
            ws.write_datetime(row, col, &datetime)
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write datetime: {}", e)))?;

        Ok(())
    }

    /// Write a datetime from year, month, day, hour, minute, second components
    #[napi]
    pub fn write_datetime_from_ymd_hms(&self, row: u32, col: u16, year: u16, month: u8, day: u8, hour: u16, min: u8, sec: f64, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let datetime = ExcelDateTime::from_ymd(year, month, day)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to create date: {}", e)))?
            .and_hms(hour, min, sec)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add time: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_datetime_with_format(row, col, &datetime, &f)
        } else {
            ws.write_datetime(row, col, &datetime)
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write datetime: {}", e)))?;

        Ok(())
    }

    /// Write a datetime from an ISO 8601 string (e.g., "2023-01-25T12:30:00")
    #[napi]
    pub fn write_datetime_from_string(&self, row: u32, col: u16, datetime_str: String, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let datetime = ExcelDateTime::parse_from_str(&datetime_str)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to parse datetime string: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_datetime_with_format(row, col, &datetime, &f)
        } else {
            ws.write_datetime(row, col, &datetime)
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write datetime: {}", e)))?;

        Ok(())
    }

    /// Write a value with formatting to a cell
    #[napi]
    pub fn write_with_format(&self, row: u32, col: u16, value: Either3<String, f64, bool>, format: &RuscFormat) -> Result<()> {
        self.write(row, col, value, Some(format))
    }

    /// Set the width of a column
    #[napi]
    pub fn set_column_width(&self, col: u16, width: f64) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_column_width(col, width)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set column width: {}", e)))?;
        Ok(())
    }

    /// Set the height of a row
    #[napi]
    pub fn set_row_height(&self, row: u32, height: f64) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_row_height(row, height)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set row height: {}", e)))?;
        Ok(())
    }

    /// Set the name of the worksheet
    #[napi]
    pub fn set_name(&self, name: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_name(&name)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set worksheet name: {}", e)))?;
        Ok(())
    }

    /// Merge a range of cells
    #[napi]
    pub fn merge_range(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, value: String, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.merge_range(first_row, first_col, last_row, last_col, &value, &f)
        } else {
            ws.merge_range(first_row, first_col, last_row, last_col, &value, &Format::new())
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to merge range: {}", e)))?;

        Ok(())
    }

    /// Merge a range of cells with formatting
    #[napi]
    pub fn merge_range_with_format(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, value: String, format: &RuscFormat) -> Result<()> {
        self.merge_range(first_row, first_col, last_row, last_col, value, Some(format))
    }

    /// Insert a note (post-it style comment) to a cell
    #[napi]
    pub fn insert_note(&self, row: u32, col: u16, note: &RuscNote) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_note(row, col, note.get_note())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert note: {}", e)))?;
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

        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_default_note_author(&name);
        Ok(())
    }

    /// Write a formula to a cell
    #[napi]
    pub fn write_formula(&self, row: u32, col: u16, formula: &RuscFormula, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_formula_with_format(row, col, formula.get_formula(), &f)
        } else {
            ws.write_formula(row, col, formula.get_formula())
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write formula: {}", e)))?;

        Ok(())
    }

    /// Write a dynamic array formula to a cell (Excel 365)
    #[napi]
    pub fn write_dynamic_formula(&self, row: u32, col: u16, formula: &RuscFormula, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_dynamic_array_formula_with_format(row, col, row, col, formula.get_formula(), &f)
        } else {
            ws.write_dynamic_array_formula(row, col, row, col, formula.get_formula())
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write dynamic formula: {}", e)))?;

        Ok(())
    }

    /// Write an array formula to a range of cells
    #[napi]
    pub fn write_array_formula(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, formula: &RuscFormula, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_array_formula_with_format(first_row, first_col, last_row, last_col, formula.get_formula(), &f)
        } else {
            ws.write_array_formula(first_row, first_col, last_row, last_col, formula.get_formula())
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write array formula: {}", e)))?;

        Ok(())
    }

    /// Add data validation to a cell or range
    #[napi]
    pub fn add_data_validation(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, validation: &RuscDataValidation) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.add_data_validation(first_row, first_col, last_row, last_col, validation.get_validation())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add data validation: {}", e)))?;

        Ok(())
    }

    /// Write a URL to a cell
    #[napi]
    pub fn write_url(&self, row: u32, col: u16, url: String, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let url_obj = Url::new(&url);

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_url_with_format(row, col, &url_obj, &f)
        } else {
            ws.write_url(row, col, &url_obj)
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write URL: {}", e)))?;

        Ok(())
    }

    /// Write a URL to a cell with custom display text
    #[napi]
    pub fn write_url_with_text(&self, row: u32, col: u16, url: String, text: String, format: Option<&RuscFormat>) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let url_obj = Url::new(&url).set_text(&text);

        if let Some(fmt) = format {
            let f = fmt.to_format();
            ws.write_url_with_format(row, col, &url_obj, &f)
        } else {
            ws.write_url(row, col, &url_obj)
        }
        .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to write URL: {}", e)))?;

        Ok(())
    }

    /// Set autofilter on a range
    #[napi]
    pub fn autofilter(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.autofilter(first_row, first_col, last_row, last_col)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set autofilter: {}", e)))?;

        Ok(())
    }

    /// Freeze panes at the specified row and column
    #[napi]
    pub fn set_freeze_panes(&self, row: u32, col: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_freeze_panes(row, col)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set freeze panes: {}", e)))?;

        Ok(())
    }

    /// Add 2-color scale conditional formatting
    #[napi]
    pub fn add_conditional_format_2color_scale(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, min_color: String, max_color: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let min = parse_color(&min_color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid min color: {}", min_color)))?;
        let max = parse_color(&max_color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid max color: {}", max_color)))?;

        let format = ConditionalFormat2ColorScale::new()
            .set_minimum_color(min)
            .set_maximum_color(max);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add 2-color scale: {}", e)))?;

        Ok(())
    }

    /// Add 3-color scale conditional formatting
    #[napi]
    pub fn add_conditional_format_3color_scale(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, min_color: String, mid_color: String, max_color: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let min = parse_color(&min_color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid min color: {}", min_color)))?;
        let mid = parse_color(&mid_color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid mid color: {}", mid_color)))?;
        let max = parse_color(&max_color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid max color: {}", max_color)))?;

        let format = ConditionalFormat3ColorScale::new()
            .set_minimum_color(min)
            .set_midpoint_color(mid)
            .set_maximum_color(max);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add 3-color scale: {}", e)))?;

        Ok(())
    }

    /// Add data bar conditional formatting
    #[napi]
    pub fn add_conditional_format_data_bar(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, bar_color: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let color = parse_color(&bar_color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid bar color: {}", bar_color)))?;

        let format = ConditionalFormatDataBar::new()
            .set_fill_color(color);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add data bar: {}", e)))?;

        Ok(())
    }

    /// Add cell-based conditional formatting (greater than, less than, etc.)
    #[napi]
    pub fn add_conditional_format_cell(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, rule: String, value: f64, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

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
                    format!("Invalid cell rule: {}. Valid values: 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'equal', 'notEqual'", rule),
                ))
            }
        };

        let cond_format = ConditionalFormatCell::new()
            .set_rule(cell_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add cell conditional format: {}", e)))?;

        Ok(())
    }

    /// Add text-based conditional formatting
    #[napi]
    pub fn add_conditional_format_text(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, rule: String, text: String, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let text_rule = match rule.as_str() {
            "contains" => ConditionalFormatTextRule::Contains(text),
            "notContains" => ConditionalFormatTextRule::DoesNotContain(text),
            "beginsWith" => ConditionalFormatTextRule::BeginsWith(text),
            "endsWith" => ConditionalFormatTextRule::EndsWith(text),
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!("Invalid text rule: {}. Valid values: 'contains', 'notContains', 'beginsWith', 'endsWith'", rule),
                ))
            }
        };

        let cond_format = ConditionalFormatText::new()
            .set_rule(text_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add text conditional format: {}", e)))?;

        Ok(())
    }

    /// Add average/standard deviation conditional formatting
    #[napi]
    pub fn add_conditional_format_average(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, rule: String, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let avg_rule = match rule.as_str() {
            "aboveAverage" => ConditionalFormatAverageRule::AboveAverage,
            "belowAverage" => ConditionalFormatAverageRule::BelowAverage,
            "equalOrAboveAverage" => ConditionalFormatAverageRule::EqualOrAboveAverage,
            "equalOrBelowAverage" => ConditionalFormatAverageRule::EqualOrBelowAverage,
            _ => return Err(Error::new(
                Status::GenericFailure,
                format!("Invalid average rule: {}. Valid values: 'aboveAverage', 'belowAverage', 'equalOrAboveAverage', 'equalOrBelowAverage'", rule),
            ))
        };

        let cond_format = ConditionalFormatAverage::new()
            .set_rule(avg_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add average conditional format: {}", e)))?;

        Ok(())
    }

    /// Add top/bottom N conditional formatting
    #[napi]
    pub fn add_conditional_format_top(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, rule: String, value: u16, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let top_rule = match rule.as_str() {
            "top" => ConditionalFormatTopRule::Top(value),
            "bottom" => ConditionalFormatTopRule::Bottom(value),
            "topPercent" => ConditionalFormatTopRule::TopPercent(value),
            "bottomPercent" => ConditionalFormatTopRule::BottomPercent(value),
            _ => return Err(Error::new(
                Status::GenericFailure,
                format!("Invalid top/bottom rule: {}. Valid values: 'top', 'bottom', 'topPercent', 'bottomPercent'", rule),
            ))
        };

        let cond_format = ConditionalFormatTop::new()
            .set_rule(top_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add top/bottom conditional format: {}", e)))?;

        Ok(())
    }

    /// Add duplicate/unique values conditional formatting
    #[napi]
    pub fn add_conditional_format_duplicate(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, invert: bool, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let mut cond_format = ConditionalFormatDuplicate::new().set_format(format.to_format());

        if invert {
            cond_format = cond_format.invert();
        }

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add duplicate conditional format: {}", e)))?;

        Ok(())
    }

    /// Add blank/non-blank conditional formatting
    #[napi]
    pub fn add_conditional_format_blank(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, invert: bool, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let mut cond_format = ConditionalFormatBlank::new().set_format(format.to_format());

        if invert {
            cond_format = cond_format.invert();
        }

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add blank conditional format: {}", e)))?;

        Ok(())
    }

    /// Add error/non-error conditional formatting
    #[napi]
    pub fn add_conditional_format_error(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, invert: bool, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let mut cond_format = ConditionalFormatError::new().set_format(format.to_format());

        if invert {
            cond_format = cond_format.invert();
        }

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add error conditional format: {}", e)))?;

        Ok(())
    }

    /// Add date-based conditional formatting
    #[napi]
    pub fn add_conditional_format_date(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, rule: String, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

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
            _ => return Err(Error::new(
                Status::GenericFailure,
                format!("Invalid date rule: {}. Valid values: 'yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek', 'lastMonth', 'thisMonth', 'nextMonth'", rule),
            ))
        };

        let cond_format = ConditionalFormatDate::new()
            .set_rule(date_rule)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add date conditional format: {}", e)))?;

        Ok(())
    }

    /// Add formula-based conditional formatting
    #[napi]
    pub fn add_conditional_format_formula(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, formula: String, format: &RuscFormat) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let cond_format = ConditionalFormatFormula::new()
            .set_rule(&*formula)
            .set_format(format.to_format());

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add formula conditional format: {}", e)))?;

        Ok(())
    }

    /// Add icon set conditional formatting
    #[napi]
    pub fn add_conditional_format_icon_set(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, icon_type: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

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
            _ => return Err(Error::new(
                Status::GenericFailure,
                format!("Invalid icon set type: {}. Valid values: 'threeArrows', 'threeArrowsGray', 'threeFlags', 'threeTrafficLights', 'threeTrafficLightsWithRim', 'threeSigns', 'threeSymbols', 'threeSymbolsCircled', 'threeStars', 'threeTriangles', 'fourArrows', 'fourArrowsGray', 'fourRedToBlack', 'fourHistograms', 'fourTrafficLights', 'fiveArrows', 'fiveArrowsGray', 'fiveHistograms', 'fiveQuadrants', 'fiveBoxes'", icon_type),
            ))
        };

        let cond_format = ConditionalFormatIconSet::new().set_icon_type(icon_set_type);

        ws.add_conditional_format(first_row, first_col, last_row, last_col, &cond_format)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add icon set conditional format: {}", e)))?;

        Ok(())
    }

    /// Set page margins (in inches)
    #[napi]
    pub fn set_margins(&self, left: f64, right: f64, top: f64, bottom: f64, header: f64, footer: f64) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_margins(left, right, top, bottom, header, footer);
        Ok(())
    }

    /// Set page orientation to portrait
    #[napi]
    pub fn set_portrait(&self) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_portrait();
        Ok(())
    }

    /// Set page orientation to landscape
    #[napi]
    pub fn set_landscape(&self) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_landscape();
        Ok(())
    }

    /// Set paper size (1=Letter, 9=A4, etc.)
    #[napi]
    pub fn set_paper_size(&self, paper_size: u8) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_paper_size(paper_size);
        Ok(())
    }

    /// Set header text
    #[napi]
    pub fn set_header(&self, header: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_header(&header);
        Ok(())
    }

    /// Set footer text
    #[napi]
    pub fn set_footer(&self, footer: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_footer(&footer);
        Ok(())
    }

    /// Set print area
    #[napi]
    pub fn set_print_area(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_area(first_row, first_col, last_row, last_col)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set print area: {}", e)))?;
        Ok(())
    }

    /// Set print gridlines
    #[napi]
    pub fn set_print_gridlines(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_gridlines(enable);
        Ok(())
    }

    /// Center content horizontally on the page when printing
    #[napi]
    pub fn set_print_center_horizontally(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_center_horizontally(enable);
        Ok(())
    }

    /// Center content vertically on the page when printing
    #[napi]
    pub fn set_print_center_vertically(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_center_vertically(enable);
        Ok(())
    }

    /// Fit worksheet to N pages wide by M pages tall when printing
    #[napi]
    pub fn set_print_fit_to_pages(&self, width: u16, height: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_fit_to_pages(width, height);
        Ok(())
    }

    /// Set print scale as a percentage (10-400)
    #[napi]
    pub fn set_print_scale(&self, scale: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_scale(scale);
        Ok(())
    }

    /// Set the first page number for printing
    #[napi]
    pub fn set_print_first_page_number(&self, page_number: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_first_page_number(page_number);
        Ok(())
    }

    /// Enable/disable black and white printing
    #[napi]
    pub fn set_print_black_and_white(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_black_and_white(enable);
        Ok(())
    }

    /// Enable/disable draft quality printing
    #[napi]
    pub fn set_print_draft(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_draft(enable);
        Ok(())
    }

    /// Enable/disable printing of row and column headings
    #[napi]
    pub fn set_print_headings(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_print_headings(enable);
        Ok(())
    }

    /// Repeat rows at the top of each printed page
    #[napi]
    pub fn set_repeat_rows(&self, first_row: u32, last_row: u32) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_repeat_rows(first_row, last_row)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set repeat rows: {}", e)))?;
        Ok(())
    }

    /// Repeat columns at the left of each printed page
    #[napi]
    pub fn set_repeat_columns(&self, first_col: u16, last_col: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_repeat_columns(first_col, last_col)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set repeat columns: {}", e)))?;
        Ok(())
    }

    /// Scale header and footer with document
    #[napi]
    pub fn set_header_footer_scale_with_doc(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_header_footer_scale_with_doc(enable);
        Ok(())
    }

    /// Align header and footer with page margins
    #[napi]
    pub fn set_header_footer_align_with_page(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_header_footer_align_with_page(enable);
        Ok(())
    }

    /// Set the worksheet tab color
    #[napi]
    pub fn set_tab_color(&self, color: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        let color_obj = parse_color(&color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid color: {}", color)))?;

        ws.set_tab_color(color_obj);
        Ok(())
    }

    /// Hide or show the worksheet
    #[napi]
    pub fn set_hidden(&self, enable: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

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

        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_zoom(zoom);
        Ok(())
    }

    /// Set the default cell selection when opening the worksheet
    #[napi]
    pub fn set_selection(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_selection(first_row, first_col, last_row, last_col)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set selection: {}", e)))?;
        Ok(())
    }

    /// Set which cell appears in the top-left corner when opening the worksheet
    #[napi]
    pub fn set_top_left_cell(&self, row: u32, col: u16) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.set_top_left_cell(row, col)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set top left cell: {}", e)))?;
        Ok(())
    }

    /// Protect the worksheet from editing
    #[napi]
    pub fn protect(&self) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.protect();
        Ok(())
    }

    /// Protect the worksheet with a password
    #[napi]
    pub fn protect_with_password(&self, password: String) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.protect_with_password(&password);
        Ok(())
    }

    /// Protect the worksheet with granular protection options
    #[napi]
    pub fn protect_with_options(&self, options: &RuscProtectionOptions) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.protect_with_options(options.get_options());
        Ok(())
    }

    /// Protect the worksheet with a password and granular protection options
    #[napi]
    pub fn protect_with_password_and_options(&self, password: String, options: &RuscProtectionOptions) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.protect_with_password(&password);
        ws.protect_with_options(options.get_options());
        Ok(())
    }

    /// Insert a chart at the specified cell position
    #[napi]
    pub fn insert_chart(&self, row: u32, col: u16, chart: &RuscChart) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        // Create a clone of the chart since rust_xlsxwriter takes by reference
        // Note: This is a workaround. In the future, we might need to handle chart lifecycle differently
        ws.insert_chart(row, col, &chart.chart)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert chart: {}", e)))?;

        Ok(())
    }

    /// Insert a chart at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_chart_with_offset(&self, row: u32, col: u16, chart: &RuscChart, x_offset: u32, y_offset: u32) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_chart_with_offset(row, col, &chart.chart, x_offset, y_offset)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert chart with offset: {}", e)))?;

        Ok(())
    }

    /// Insert an image at the specified cell position
    #[napi]
    pub fn insert_image(&self, row: u32, col: u16, image: &RuscImage) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_image(row, col, image.get_image())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert image: {}", e)))?;

        Ok(())
    }

    /// Insert an image at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_image_with_offset(&self, row: u32, col: u16, image: &RuscImage, x_offset: u32, y_offset: u32) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_image_with_offset(row, col, image.get_image(), x_offset, y_offset)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert image with offset: {}", e)))?;

        Ok(())
    }

    /// Insert an image that fits to a cell
    #[napi]
    pub fn insert_image_fit_to_cell(&self, row: u32, col: u16, image: &RuscImage, keep_aspect_ratio: bool) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_image_fit_to_cell(row, col, image.get_image(), keep_aspect_ratio)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert fitted image: {}", e)))?;

        Ok(())
    }

    /// Add a table to the worksheet
    #[napi]
    pub fn add_table(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, table: &RuscTable) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.add_table(first_row, first_col, last_row, last_col, table.get_table())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add table: {}", e)))?;

        Ok(())
    }

    /// Add a sparkline to a single cell
    #[napi]
    pub fn add_sparkline(&self, row: u32, col: u16, sparkline: &RuscSparkline) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.add_sparkline(row, col, sparkline.get_sparkline())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add sparkline: {}", e)))?;

        Ok(())
    }

    /// Add a sparkline group to a range of cells
    #[napi]
    pub fn add_sparkline_group(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, sparkline: &RuscSparkline) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.add_sparkline_group(first_row, first_col, last_row, last_col, sparkline.get_sparkline())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add sparkline group: {}", e)))?;

        Ok(())
    }

    /// Insert a button at the specified cell position
    #[napi]
    pub fn insert_button(&self, row: u32, col: u16, button: &RuscButton) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_button(row, col, button.get_button())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert button: {}", e)))?;

        Ok(())
    }

    /// Insert a button at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_button_with_offset(&self, row: u32, col: u16, button: &RuscButton, x_offset: u32, y_offset: u32) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_button_with_offset(row, col, button.get_button(), x_offset, y_offset)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert button with offset: {}", e)))?;

        Ok(())
    }

    /// Insert a shape (textbox) at the specified cell position
    #[napi]
    pub fn insert_shape(&self, row: u32, col: u16, shape: &RuscShape) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_shape(row, col, shape.get_shape())
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert shape: {}", e)))?;

        Ok(())
    }

    /// Insert a shape (textbox) at the specified cell position with pixel offsets
    #[napi]
    pub fn insert_shape_with_offset(&self, row: u32, col: u16, shape: &RuscShape, x_offset: u32, y_offset: u32) -> Result<()> {
        let mut wb = self.workbook.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to get worksheet: {}", e)))?;

        ws.insert_shape_with_offset(row, col, shape.get_shape(), x_offset, y_offset)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to insert shape with offset: {}", e)))?;

        Ok(())
    }
}

