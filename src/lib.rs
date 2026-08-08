// Module declarations
mod button;
mod chart;
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
pub use formula::RuscFormula;
pub use image::RuscImage;
pub use note::RuscNote;
pub use properties::RuscDocProperties;
pub use protection::RuscProtectionOptions;
pub use richtext::RuscRichText;
pub use shape::RuscShape;
pub use sparkline::RuscSparkline;
pub use table::{RuscTable, RuscTableColumn};
pub use validation::RuscDataValidation;
pub use workbook::RuscWorkbook;
pub use worksheet::RuscWorksheet;

use napi_derive::napi;

/// Returns [currentBytes, peakBytes] of Rust allocator usage
#[napi]
pub fn rust_memory_stats() -> Vec<f64> {
    let (current, peak) = memory::get_stats();
    vec![current as f64, peak as f64]
}

/// Resets peak memory tracking to current usage
#[napi]
pub fn rust_memory_reset_peak() {
    memory::reset_peak();
}
