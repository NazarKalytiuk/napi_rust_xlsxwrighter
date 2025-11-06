// Button module for Excel form control buttons
use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Button;

#[napi]
pub struct RuscButton {
    pub(crate) button: Button,
}

#[napi]
impl RuscButton {
    /// Create a new Button
    #[napi(factory)]
    pub fn new() -> Self {
        RuscButton {
            button: Button::new(),
        }
    }

    /// Set the button caption/text
    #[napi]
    pub fn set_caption(&mut self, caption: String) -> &Self {
        self.button = self.button.clone().set_caption(&caption);
        self
    }

    /// Set the VBA macro to run when button is clicked
    #[napi]
    pub fn set_macro(&mut self, macro_name: String) -> &Self {
        self.button = self.button.clone().set_macro(&macro_name);
        self
    }

    /// Set the button width in pixels
    #[napi]
    pub fn set_width(&mut self, width: u32) -> &Self {
        self.button = self.button.clone().set_width(width);
        self
    }

    /// Set the button height in pixels
    #[napi]
    pub fn set_height(&mut self, height: u32) -> &Self {
        self.button = self.button.clone().set_height(height);
        self
    }

    /// Set the button alt text for accessibility
    #[napi]
    pub fn set_alt_text(&mut self, alt_text: String) -> &Self {
        self.button = self.button.clone().set_alt_text(&alt_text);
        self
    }

    /// Set how the button moves and sizes with cells
    #[napi]
    pub fn set_object_movement(&mut self, option: String) -> Result<&Self> {
        use rust_xlsxwriter::ObjectMovement;

        let movement = match option.as_str() {
            "moveAndSize" => ObjectMovement::MoveAndSizeWithCells,
            "moveOnly" => ObjectMovement::MoveButDontSizeWithCells,
            "none" => ObjectMovement::DontMoveOrSizeWithCells,
            "moveAndSizeAfter" => ObjectMovement::MoveAndSizeWithCellsAfter,
            _ => return Err(Error::new(Status::InvalidArg, format!("Invalid object movement option: {}. Use 'moveAndSize', 'moveOnly', 'none', or 'moveAndSizeAfter'", option))),
        };

        self.button = self.button.clone().set_object_movement(movement);
        Ok(self)
    }

    /// Get the internal button reference
    pub(crate) fn get_button(&self) -> &Button {
        &self.button
    }
}
