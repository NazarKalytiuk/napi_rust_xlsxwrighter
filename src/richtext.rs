use napi_derive::napi;

use crate::format::RuscFormat;

/// Wrapper for rich text strings (multi-format text in a single cell)
#[napi]
pub struct RuscRichText {
    pub(crate) segments: Vec<(String, RuscFormat)>,
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

    /// Add a text segment with formatting
    /// The format is cloned internally for storage
    #[napi]
    pub fn add_segment(&mut self, text: String, format: &RuscFormat) -> &Self {
        // Clone the format for storage
        self.segments.push((text, format.clone()));
        self
    }

    pub(crate) fn get_segments(&self) -> &Vec<(String, RuscFormat)> {
        &self.segments
    }
}
