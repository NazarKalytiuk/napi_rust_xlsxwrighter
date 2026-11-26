use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{
    Color, FontScheme, Format, FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern,
    FormatScript, FormatUnderline,
};

use crate::utils::{parse_border, parse_color};

#[napi]
#[derive(Clone)]
pub struct RuscFormat {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    font_size: Option<f64>,
    font_name: Option<String>,
    font_scheme: Option<FontScheme>,
    font_color: Option<Color>,
    underline: Option<String>,
    strikethrough: bool,
    font_script: Option<FormatScript>,
    background_color: Option<Color>,
    foreground_color: Option<Color>,
    pattern: Option<FormatPattern>,
    align: Option<String>,
    vertical_align: Option<String>,
    text_wrap: bool,
    indent: Option<u8>,
    rotation: Option<i16>,
    shrink: bool,
    reading_direction: Option<u8>,
    border: Option<FormatBorder>,
    border_top: Option<FormatBorder>,
    border_bottom: Option<FormatBorder>,
    border_left: Option<FormatBorder>,
    border_right: Option<FormatBorder>,
    border_diagonal: Option<FormatBorder>,
    border_color: Option<Color>,
    border_top_color: Option<Color>,
    border_bottom_color: Option<Color>,
    border_left_color: Option<Color>,
    border_right_color: Option<Color>,
    border_diagonal_color: Option<Color>,
    border_diagonal_type: Option<FormatDiagonalBorder>,
    num_format: Option<String>,
    unlocked: bool,
    hidden: bool,
    quote_prefix: bool,
}

impl Default for RuscFormat {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscFormat {
    /// Create a new format
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscFormat {
            bold: false,
            italic: false,
            font_size: None,
            font_name: None,
            font_scheme: None,
            font_color: None,
            underline: None,
            strikethrough: false,
            font_script: None,
            background_color: None,
            foreground_color: None,
            pattern: None,
            align: None,
            vertical_align: None,
            text_wrap: false,
            indent: None,
            rotation: None,
            shrink: false,
            reading_direction: None,
            border: None,
            border_top: None,
            border_bottom: None,
            border_left: None,
            border_right: None,
            border_diagonal: None,
            border_color: None,
            border_top_color: None,
            border_bottom_color: None,
            border_left_color: None,
            border_right_color: None,
            border_diagonal_color: None,
            border_diagonal_type: None,
            num_format: None,
            unlocked: false,
            hidden: false,
            quote_prefix: false,
        }
    }

    /// Set bold text
    #[napi]
    pub fn set_bold(&mut self) -> Result<()> {
        self.bold = true;
        Ok(())
    }

    /// Set italic text
    #[napi]
    pub fn set_italic(&mut self) -> Result<()> {
        self.italic = true;
        Ok(())
    }

    /// Set font size
    #[napi]
    pub fn set_font_size(&mut self, size: f64) -> Result<()> {
        self.font_size = Some(size);
        Ok(())
    }

    /// Set font color
    #[napi]
    pub fn set_font_color(&mut self, color: String) -> Result<()> {
        self.font_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set background color
    #[napi]
    pub fn set_background_color(&mut self, color: String) -> Result<()> {
        self.background_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set horizontal alignment
    #[napi]
    pub fn set_align(&mut self, align: String) -> Result<()> {
        match align.as_str() {
            "left" | "center" | "right" | "fill" | "justify" | "centerAcross" | "distributed" => {
                self.align = Some(align);
                Ok(())
            }
            _ => Err(Error::new(
                Status::GenericFailure,
                format!("Invalid alignment: {}", align),
            )),
        }
    }

    /// Set vertical alignment
    #[napi]
    pub fn set_vertical_align(&mut self, align: String) -> Result<()> {
        match align.as_str() {
            "top" | "center" | "bottom" | "justify" | "distributed" => {
                self.vertical_align = Some(align);
                Ok(())
            }
            _ => Err(Error::new(
                Status::GenericFailure,
                format!("Invalid vertical alignment: {}", align),
            )),
        }
    }

    /// Set border style
    #[napi]
    pub fn set_border(&mut self, border: String) -> Result<()> {
        self.border = Some(parse_border(&border)?);
        Ok(())
    }

    /// Set number format
    #[napi]
    pub fn set_num_format(&mut self, num_format: String) -> Result<()> {
        self.num_format = Some(num_format);
        Ok(())
    }

    /// Set font name (e.g., 'Arial', 'Calibri', 'Times New Roman')
    #[napi]
    pub fn set_font_name(&mut self, font_name: String) -> Result<()> {
        self.font_name = Some(font_name);
        Ok(())
    }
    /// Set whether the font uses a body, headings, or custom font scheme.
    #[napi]
    pub fn set_font_scheme(&mut self, scheme: String) -> Result<()> {
        self.font_scheme = Some(match scheme.as_str() {
            "body" => FontScheme::Body,
            "headings" => FontScheme::Headings,
            "none" => FontScheme::None,
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid font scheme: {}. Valid values: 'body', 'headings', 'none'",
                        scheme
                    ),
                ));
            }
        });
        Ok(())
    }

    /// Set text underline style
    #[napi]
    pub fn set_underline(&mut self, underline: String) -> Result<()> {
        match underline.as_str() {
            "single" | "double" | "singleAccounting" | "doubleAccounting" => {
                self.underline = Some(underline);
                Ok(())
            }
            _ => Err(Error::new(
                Status::GenericFailure,
                format!(
                    "Invalid underline style: {}. Valid values: 'single', 'double', 'singleAccounting', 'doubleAccounting'",
                    underline
                ),
            )),
        }
    }

    /// Set strikethrough text
    #[napi]
    pub fn set_strikethrough(&mut self) -> Result<()> {
        self.strikethrough = true;
        Ok(())
    }

    /// Set foreground color (for patterns)
    #[napi]
    pub fn set_foreground_color(&mut self, color: String) -> Result<()> {
        self.foreground_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set fill pattern
    #[napi]
    pub fn set_pattern(&mut self, pattern: String) -> Result<()> {
        let pattern_enum = match pattern.as_str() {
            "solid" => FormatPattern::Solid,
            "mediumGray" => FormatPattern::MediumGray,
            "darkGray" => FormatPattern::DarkGray,
            "lightGray" => FormatPattern::LightGray,
            "darkHorizontal" => FormatPattern::DarkHorizontal,
            "darkVertical" => FormatPattern::DarkVertical,
            "darkDown" => FormatPattern::DarkDown,
            "darkUp" => FormatPattern::DarkUp,
            "darkGrid" => FormatPattern::DarkGrid,
            "darkTrellis" => FormatPattern::DarkTrellis,
            "lightHorizontal" => FormatPattern::LightHorizontal,
            "lightVertical" => FormatPattern::LightVertical,
            "lightDown" => FormatPattern::LightDown,
            "lightUp" => FormatPattern::LightUp,
            "lightGrid" => FormatPattern::LightGrid,
            "lightTrellis" => FormatPattern::LightTrellis,
            "gray125" => FormatPattern::Gray125,
            "gray0625" => FormatPattern::Gray0625,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!("Invalid pattern: {}", pattern),
                ));
            }
        };
        self.pattern = Some(pattern_enum);
        Ok(())
    }

    /// Enable text wrapping in the cell
    #[napi]
    pub fn set_text_wrap(&mut self) -> Result<()> {
        self.text_wrap = true;
        Ok(())
    }

    /// Set text indentation level (0-15)
    #[napi]
    pub fn set_indent(&mut self, indent: u32) -> Result<()> {
        if indent > 15 {
            return Err(Error::new(
                Status::GenericFailure,
                "Indent level must be between 0 and 15".to_string(),
            ));
        }
        self.indent = Some(indent as u8);
        Ok(())
    }

    /// Set text rotation in degrees (-90 to 90, or 270 for vertical text)
    #[napi]
    pub fn set_rotation(&mut self, rotation: i32) -> Result<()> {
        if !(-90..=90).contains(&rotation) && rotation != 270 {
            return Err(Error::new(
                Status::GenericFailure,
                "Rotation must be between -90 and 90, or 270 for vertical text".to_string(),
            ));
        }
        self.rotation = Some(rotation as i16);
        Ok(())
    }

    /// Shrink text to fit cell
    #[napi]
    pub fn set_shrink(&mut self) -> Result<()> {
        self.shrink = true;
        Ok(())
    }

    /// Set top border style
    #[napi]
    pub fn set_border_top(&mut self, border: String) -> Result<()> {
        self.border_top = Some(parse_border(&border)?);
        Ok(())
    }

    /// Set bottom border style
    #[napi]
    pub fn set_border_bottom(&mut self, border: String) -> Result<()> {
        self.border_bottom = Some(parse_border(&border)?);
        Ok(())
    }

    /// Set left border style
    #[napi]
    pub fn set_border_left(&mut self, border: String) -> Result<()> {
        self.border_left = Some(parse_border(&border)?);
        Ok(())
    }

    /// Set right border style
    #[napi]
    pub fn set_border_right(&mut self, border: String) -> Result<()> {
        self.border_right = Some(parse_border(&border)?);
        Ok(())
    }

    /// Set color for all borders
    #[napi]
    pub fn set_border_color(&mut self, color: String) -> Result<()> {
        self.border_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set top border color
    #[napi]
    pub fn set_border_top_color(&mut self, color: String) -> Result<()> {
        self.border_top_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set bottom border color
    #[napi]
    pub fn set_border_bottom_color(&mut self, color: String) -> Result<()> {
        self.border_bottom_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set left border color
    #[napi]
    pub fn set_border_left_color(&mut self, color: String) -> Result<()> {
        self.border_left_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set right border color
    #[napi]
    pub fn set_border_right_color(&mut self, color: String) -> Result<()> {
        self.border_right_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Unlock cell (for use with worksheet protection)
    #[napi]
    pub fn set_unlocked(&mut self) -> Result<()> {
        self.unlocked = true;
        Ok(())
    }

    /// Lock cell (cell will be protected when worksheet is protected)
    #[napi]
    pub fn set_locked(&mut self) -> Result<()> {
        self.unlocked = false;
        Ok(())
    }

    /// Hide formulas in cell (for use with worksheet protection)
    #[napi]
    pub fn set_hidden(&mut self) -> Result<()> {
        self.hidden = true;
        Ok(())
    }

    /// Set font script (superscript or subscript)
    #[napi]
    pub fn set_font_script(&mut self, script: String) -> Result<()> {
        self.font_script = Some(match script.to_lowercase().as_str() {
            "superscript" => FormatScript::Superscript,
            "subscript" => FormatScript::Subscript,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid font script: {}. Use 'superscript' or 'subscript'",
                        script
                    ),
                ));
            }
        });
        Ok(())
    }

    /// Set diagonal border style
    #[napi]
    pub fn set_border_diagonal(&mut self, border: String) -> Result<()> {
        self.border_diagonal = Some(parse_border(&border)?);
        Ok(())
    }

    /// Set diagonal border color
    #[napi]
    pub fn set_border_diagonal_color(&mut self, color: String) -> Result<()> {
        self.border_diagonal_color = Some(parse_color(&color)?);
        Ok(())
    }

    /// Set diagonal border type (up, down, or both)
    #[napi]
    pub fn set_border_diagonal_type(&mut self, border_type: String) -> Result<()> {
        self.border_diagonal_type = Some(match border_type.to_lowercase().as_str() {
            "up" => FormatDiagonalBorder::BorderUp,
            "down" => FormatDiagonalBorder::BorderDown,
            "both" | "updown" => FormatDiagonalBorder::BorderUpDown,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!(
                        "Invalid diagonal border type: {}. Use 'up', 'down', or 'both'",
                        border_type
                    ),
                ));
            }
        });
        Ok(())
    }

    /// Set reading direction (0=context, 1=left-to-right, 2=right-to-left)
    #[napi]
    pub fn set_reading_direction(&mut self, direction: u8) -> Result<()> {
        if direction > 2 {
            return Err(Error::new(
                Status::GenericFailure,
                "Reading direction must be 0 (context), 1 (LTR), or 2 (RTL)".to_string(),
            ));
        }
        self.reading_direction = Some(direction);
        Ok(())
    }

    /// Set quote prefix (force string interpretation with leading apostrophe)
    #[napi]
    pub fn set_quote_prefix(&mut self) -> Result<()> {
        self.quote_prefix = true;
        Ok(())
    }

    /// Whether the format sets a number format of its own.
    pub(crate) fn has_num_format(&self) -> bool {
        self.num_format.is_some()
    }

    /// Internal method to convert to rust_xlsxwriter Format
    pub(crate) fn to_format(&self) -> Format {
        let mut fmt = Format::new();

        // Font properties
        if self.bold {
            fmt = fmt.set_bold();
        }

        if self.italic {
            fmt = fmt.set_italic();
        }

        if self.strikethrough {
            fmt = fmt.set_font_strikethrough();
        }

        if let Some(size) = self.font_size {
            fmt = fmt.set_font_size(size);
        }

        if let Some(ref name) = self.font_name {
            fmt = fmt.set_font_name(name);
        }

        if let Some(scheme) = self.font_scheme {
            fmt = fmt.set_font_scheme(scheme);
        }
        if let Some(color) = self.font_color {
            fmt = fmt.set_font_color(color);
        }

        if let Some(ref underline) = self.underline {
            let underline_enum = match underline.as_str() {
                "single" => FormatUnderline::Single,
                "double" => FormatUnderline::Double,
                "singleAccounting" => FormatUnderline::SingleAccounting,
                "doubleAccounting" => FormatUnderline::DoubleAccounting,
                _ => FormatUnderline::Single,
            };
            fmt = fmt.set_underline(underline_enum);
        }

        // Fill/background properties
        if let Some(color) = self.background_color {
            fmt = fmt.set_background_color(color);
        }

        if let Some(color) = self.foreground_color {
            fmt = fmt.set_foreground_color(color);
        }

        if let Some(pattern) = self.pattern {
            fmt = fmt.set_pattern(pattern);
        }

        // Alignment properties
        if let Some(ref align) = self.align {
            let align_enum = match align.as_str() {
                "left" => FormatAlign::Left,
                "center" => FormatAlign::Center,
                "right" => FormatAlign::Right,
                "fill" => FormatAlign::Fill,
                "justify" => FormatAlign::Justify,
                "centerAcross" => FormatAlign::CenterAcross,
                "distributed" => FormatAlign::Distributed,
                _ => FormatAlign::Left,
            };
            fmt = fmt.set_align(align_enum);
        }

        if let Some(ref align) = self.vertical_align {
            let align_enum = match align.as_str() {
                "top" => FormatAlign::Top,
                "center" => FormatAlign::VerticalCenter,
                "bottom" => FormatAlign::Bottom,
                "justify" => FormatAlign::VerticalJustify,
                "distributed" => FormatAlign::VerticalDistributed,
                _ => FormatAlign::Top,
            };
            fmt = fmt.set_align(align_enum);
        }

        if self.text_wrap {
            fmt = fmt.set_text_wrap();
        }

        if let Some(indent) = self.indent {
            fmt = fmt.set_indent(indent);
        }

        if let Some(rotation) = self.rotation {
            fmt = fmt.set_rotation(rotation);
        }

        if self.shrink {
            fmt = fmt.set_shrink();
        }

        // Border properties
        if let Some(border) = self.border {
            fmt = fmt.set_border(border);
        }

        if let Some(border) = self.border_top {
            fmt = fmt.set_border_top(border);
        }

        if let Some(border) = self.border_bottom {
            fmt = fmt.set_border_bottom(border);
        }

        if let Some(border) = self.border_left {
            fmt = fmt.set_border_left(border);
        }

        if let Some(border) = self.border_right {
            fmt = fmt.set_border_right(border);
        }

        if let Some(color) = self.border_color {
            fmt = fmt.set_border_color(color);
        }

        if let Some(color) = self.border_top_color {
            fmt = fmt.set_border_top_color(color);
        }

        if let Some(color) = self.border_bottom_color {
            fmt = fmt.set_border_bottom_color(color);
        }

        if let Some(color) = self.border_left_color {
            fmt = fmt.set_border_left_color(color);
        }

        if let Some(color) = self.border_right_color {
            fmt = fmt.set_border_right_color(color);
        }

        // Number format
        if let Some(ref num_format) = self.num_format {
            fmt = fmt.set_num_format(num_format);
        }

        // Protection
        if self.unlocked {
            fmt = fmt.set_unlocked();
        }

        if self.hidden {
            fmt = fmt.set_hidden();
        }

        // Font script (superscript/subscript)
        if let Some(script) = self.font_script {
            fmt = fmt.set_font_script(script);
        }

        // Diagonal borders
        if let Some(border) = self.border_diagonal {
            fmt = fmt.set_border_diagonal(border);
        }

        if let Some(color) = self.border_diagonal_color {
            fmt = fmt.set_border_diagonal_color(color);
        }

        if let Some(border_type) = self.border_diagonal_type {
            fmt = fmt.set_border_diagonal_type(border_type);
        }

        // Reading direction
        if let Some(direction) = self.reading_direction {
            fmt = fmt.set_reading_direction(direction);
        }

        // Quote prefix
        if self.quote_prefix {
            fmt = fmt.set_quote_prefix();
        }

        fmt
    }
}
