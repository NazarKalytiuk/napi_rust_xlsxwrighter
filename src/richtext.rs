use napi_derive::napi;

use crate::format::RuscFormat;

/// Wrapper for rich text strings (multi-format text in a single cell)
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

    pub(crate) fn get_segments(&self) -> &Vec<(String, RuscFormat)> {
        &self.segments
    }
}
