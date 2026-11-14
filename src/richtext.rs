use napi::bindgen_prelude::Either;
use napi_derive::napi;
use rust_xlsxwriter::Format;

use crate::format::RuscFormat;

/// Wrapper for rich text strings (multi-format text in a single cell)
///
/// @deprecated Pass `[text, format]` segments to `RuscWorksheet.writeRichString`
/// instead. An instance keeps its Rust memory until its finalizer runs after the
/// event loop turns, so instances created in a synchronous loop accumulate.
#[napi]
pub struct RuscRichText {
    pub(crate) segments: Vec<(String, RuscFormat)>,
}

impl Default for RuscRichText {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscRichText {
    /// Create a new empty rich text string
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscRichText {
            segments: Vec::new(),
        }
    }

    /// Add a text segment, optionally formatted; without a format the segment uses
    /// the default cell font. The format is cloned internally for storage.
    #[napi]
    pub fn add_segment(&mut self, text: String, format: Option<&RuscFormat>) -> &Self {
        self.segments
            .push((text, format.cloned().unwrap_or_default()));
        self
    }
}

/// The `rust_xlsxwriter` segments of a native rich string argument: a deprecated
/// `RuscRichText` instance or `[text, format]` segments. Empty segments, which
/// Excel rejects, are left out; a segment without a format uses the default font.
pub(crate) fn rich_string_segments<'a>(
    rich_text: &'a Either<&RuscRichText, Vec<(String, Option<&RuscFormat>)>>,
) -> Vec<(Format, &'a str)> {
    let segment = |text: &'a str, format: Option<&RuscFormat>| {
        (format.map(RuscFormat::to_format).unwrap_or_default(), text)
    };
    match rich_text {
        Either::A(rich_text) => rich_text
            .segments
            .iter()
            .filter(|(text, _)| !text.is_empty())
            .map(|(text, format)| segment(text, Some(format)))
            .collect(),
        Either::B(segments) => segments
            .iter()
            .filter(|(text, _)| !text.is_empty())
            .map(|(text, format)| segment(text, *format))
            .collect(),
    }
}
