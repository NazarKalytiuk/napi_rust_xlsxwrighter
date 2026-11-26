// Module declarations
mod button;
mod chart;
mod datetime;
mod format;
mod formula;
mod image;
mod memory;
mod note;
mod properties;
mod protection;
mod richtext;
mod shape;
mod sparkline;
mod table;
mod utils;
mod validation;
mod workbook;
mod worksheet;

// Re-export public API
pub use button::RuscButton;
pub use chart::RuscChart;
pub use format::RuscFormat;
pub use formula::{RuscFormula, RuscFormulaData};
pub use image::RuscImage;
pub use note::RuscNote;
pub use properties::RuscDocProperties;
pub use protection::RuscProtectionOptions;
pub use richtext::RuscRichText;
pub use shape::{
    RuscShape, RuscShapeFont, RuscShapeFormat, RuscShapeLine, RuscShapePatternFill,
    RuscShapeSolidFill, RuscShapeText,
};
pub use sparkline::RuscSparkline;
pub use table::{RuscTable, RuscTableColumn};
pub use validation::RuscDataValidation;
pub use workbook::RuscWorkbook;
pub use worksheet::RuscWorksheet;

use napi_derive::napi;

/// Returns [currentBytes, peakBytes] of Rust allocator usage.
/// Useful for profiling memory consumption of the native layer.
///
/// @example
/// const [current, peak] = rustMemoryStats();
/// console.log(`Rust heap: ${(current / 1024 / 1024).toFixed(2)} MB (peak: ${(peak / 1024 / 1024).toFixed(2)} MB)`);
#[napi]
pub fn rust_memory_stats() -> (f64, f64) {
    let (current, peak) = memory::get_stats();
    (current as f64, peak as f64)
}

/// Resets peak memory tracking to the current allocation level.
/// Call before a section you want to measure.
///
/// @example
/// rustMemoryResetPeak();
/// // ... do work ...
/// const [current, peak] = rustMemoryStats();
#[napi]
pub fn rust_memory_reset_peak() {
    memory::reset_peak();
}
