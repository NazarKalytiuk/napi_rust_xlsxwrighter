use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Note;

use crate::utils::parse_color;

#[napi]
pub struct RuscNote {
    pub(crate) note: Note,
}

#[napi]
impl RuscNote {
    /// Create a new note with text
    #[napi(constructor)]
    pub fn new(text: String) -> Self {
        RuscNote {
            note: Note::new(text),
        }
    }

    /// Set the author name for the note
    #[napi]
    pub fn set_author(&mut self, name: String) -> Result<()> {
        if name.chars().count() > 52 {
            return Err(Error::new(
                Status::GenericFailure,
                "Author name must be less than 52 characters".to_string(),
            ));
        }
        self.note = self.note.clone().set_author(&name);
        Ok(())
    }

    /// Set the width of the note in pixels
    #[napi]
    pub fn set_width(&mut self, width: u32) -> Result<()> {
        self.note = self.note.clone().set_width(width);
        Ok(())
    }

    /// Set the height of the note in pixels
    #[napi]
    pub fn set_height(&mut self, height: u32) -> Result<()> {
        self.note = self.note.clone().set_height(height);
        Ok(())
    }

    /// Set the background color of the note
    #[napi]
    pub fn set_background_color(&mut self, color: String) -> Result<()> {
        let c = parse_color(&color)
            .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid color: {}", color)))?;
        self.note = self.note.clone().set_background_color(c);
        Ok(())
    }

    /// Make the note visible by default (normally it's hidden until you hover over the cell)
    #[napi]
    pub fn set_visible(&mut self, visible: bool) -> Result<()> {
        self.note = self.note.clone().set_visible(visible);
        Ok(())
    }

    /// Set alt text for accessibility
    #[napi]
    pub fn set_alt_text(&mut self, alt_text: String) -> Result<()> {
        self.note = self.note.clone().set_alt_text(&alt_text);
        Ok(())
    }

    /// Get the internal Note object
    pub(crate) fn get_note(&self) -> &Note {
        &self.note
    }
}

