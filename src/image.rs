use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{Image, ObjectMovement, Url};

/// Wrapper for Image to insert images into worksheets
#[napi]
pub struct RuscImage {
    pub(crate) image: Image,
}

#[napi]
impl RuscImage {
    /// Create a new Image from a file path
    /// Supports PNG, JPG, GIF, and BMP formats
    #[napi(factory)]
    pub fn new(path: String) -> Result<RuscImage> {
        let image = Image::new(&path)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to load image: {}", e)))?;

        Ok(RuscImage { image })
    }

    /// Create an Image from a buffer (u8 array)
    #[napi(factory)]
    pub fn new_from_buffer(buffer: Buffer) -> Result<RuscImage> {
        let data = buffer.to_vec();
        let image = Image::new_from_buffer(&data)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to load image from buffer: {}", e)))?;

        Ok(RuscImage { image })
    }

    /// Set the width of the image in pixels
    #[napi]
    pub fn set_width(&mut self, width: u32) -> &Self {
        self.image = self.image.clone().set_width(width);
        self
    }

    /// Set the height of the image in pixels
    #[napi]
    pub fn set_height(&mut self, height: u32) -> &Self {
        self.image = self.image.clone().set_height(height);
        self
    }

    /// Scale the width of the image as a percentage
    #[napi]
    pub fn set_scale_width(&mut self, scale: f64) -> &Self {
        self.image = self.image.clone().set_scale_width(scale);
        self
    }

    /// Scale the height of the image as a percentage
    #[napi]
    pub fn set_scale_height(&mut self, scale: f64) -> &Self {
        self.image = self.image.clone().set_scale_height(scale);
        self
    }

    /// Scale the image to a specific size maintaining or ignoring aspect ratio
    #[napi]
    pub fn set_scale_to_size(&mut self, width: f64, height: f64, keep_aspect_ratio: bool) -> &Self {
        self.image = self.image.clone().set_scale_to_size(width, height, keep_aspect_ratio);
        self
    }

    /// Set alt text for the image (for accessibility)
    #[napi]
    pub fn set_alt_text(&mut self, alt_text: String) -> &Self {
        self.image = self.image.clone().set_alt_text(alt_text);
        self
    }

    /// Set the image as decorative (no alt text required)
    #[napi]
    pub fn set_decorative(&mut self, enable: bool) -> &Self {
        self.image = self.image.clone().set_decorative(enable);
        self
    }

    /// Set how the image moves and sizes with cells
    /// Options: 'moveAndSize', 'moveOnly', 'none', 'moveAndSizeAfter'
    #[napi]
    pub fn set_object_movement(&mut self, option: String) -> Result<&Self> {
        let movement = match option.as_str() {
            "moveAndSize" => ObjectMovement::MoveAndSizeWithCells,
            "moveOnly" => ObjectMovement::MoveButDontSizeWithCells,
            "none" => ObjectMovement::DontMoveOrSizeWithCells,
            "moveAndSizeAfter" => ObjectMovement::MoveAndSizeWithCellsAfter,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!("Invalid object movement option: {}. Valid values: 'moveAndSize', 'moveOnly', 'none', 'moveAndSizeAfter'", option),
                ))
            }
        };

        self.image = self.image.clone().set_object_movement(movement);
        Ok(self)
    }

    /// Add a URL link to the image
    #[napi]
    pub fn set_url(&mut self, url: String) -> Result<&Self> {
        let url_obj = Url::new(&url);
        self.image = self.image.clone().set_url(url_obj)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set URL: {}", e)))?;
        Ok(self)
    }

    pub(crate) fn get_image(&self) -> &Image {
        &self.image
    }
}
