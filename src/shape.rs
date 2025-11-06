// Shape module for Excel shapes (textboxes)
use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Shape;

#[napi]
pub struct RuscShape {
    pub(crate) shape: Shape,
}

#[napi]
impl RuscShape {
    /// Create a new textbox shape
    #[napi(factory)]
    pub fn textbox() -> Self {
        RuscShape {
            shape: Shape::textbox(),
        }
    }

    /// Set the text content of the textbox
    #[napi]
    pub fn set_text(&mut self, text: String) -> &Self {
        self.shape = self.shape.clone().set_text(&text);
        self
    }

    /// Set the width in pixels
    #[napi]
    pub fn set_width(&mut self, width: u32) -> &Self {
        self.shape = self.shape.clone().set_width(width);
        self
    }

    /// Set the height in pixels
    #[napi]
    pub fn set_height(&mut self, height: u32) -> &Self {
        self.shape = self.shape.clone().set_height(height);
        self
    }

    /// Set the alt text for accessibility
    #[napi]
    pub fn set_alt_text(&mut self, alt_text: String) -> &Self {
        self.shape = self.shape.clone().set_alt_text(&alt_text);
        self
    }

    /// Set how the shape moves and sizes with cells
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

        self.shape = self.shape.clone().set_object_movement(movement);
        Ok(self)
    }

    /// Get the internal shape reference
    pub(crate) fn get_shape(&self) -> &Shape {
        &self.shape
    }
}
