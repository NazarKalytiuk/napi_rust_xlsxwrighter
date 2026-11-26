use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Note;

use crate::utils::parse_rgb_color;

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
        self.note = self
            .note
            .clone()
            .set_background_color(parse_rgb_color("color", &color)?);
        Ok(())
    }

    /// Set the font name of the note text (Tahoma by default)
    #[napi]
    pub fn set_font_name(&mut self, font_name: String) -> Result<()> {
        self.note = self.note.clone().set_font_name(font_name);
        Ok(())
    }

    /// Set the font size of the note text in points (8 by default)
    #[napi]
    pub fn set_font_size(&mut self, font_size: f64) -> Result<()> {
        self.note = self.note.clone().set_font_size(font_size);
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
