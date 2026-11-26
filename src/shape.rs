// Shape module for Excel shapes (textboxes) and their formatting
use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{
    Shape, ShapeFont, ShapeFormat, ShapeLine, ShapeLineDashType, ShapePatternFill,
    ShapePatternFillType, ShapeSolidFill, ShapeText, ShapeTextDirection,
    ShapeTextHorizontalAlignment, ShapeTextVerticalAlignment,
};

use crate::utils::parse_color;

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
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid object movement option: {}. Use 'moveAndSize', 'moveOnly', 'none', or 'moveAndSizeAfter'",
                        option
                    ),
                ));
            }
        };

        self.shape = self.shape.clone().set_object_movement(movement);
        Ok(self)
    }

    /// Set the fill and line of the shape. The shape keeps a copy, so later
    /// changes to `format` don't affect it.
    #[napi]
    pub fn set_format(&mut self, format: &RuscShapeFormat) -> &Self {
        self.shape = self.shape.clone().set_format(&format.format);
        self
    }

    /// Set the font of the shape text. The shape keeps a copy.
    #[napi]
    pub fn set_font(&mut self, font: &RuscShapeFont) -> &Self {
        self.shape = self.shape.clone().set_font(&font.font);
        self
    }

    /// Set the alignment and direction of the shape text. The shape keeps a
    /// copy.
    #[napi]
    pub fn set_text_options(&mut self, options: &RuscShapeText) -> &Self {
        self.shape = self.shape.clone().set_text_options(&options.text);
        self
    }

    /// Get the internal shape reference
    pub(crate) fn get_shape(&self) -> &Shape {
        &self.shape
    }
}

/// The fill and line of a shape.
#[napi]
pub struct RuscShapeFormat {
    format: ShapeFormat,
}

impl Default for RuscShapeFormat {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscShapeFormat {
    /// Create a shape format with Excel's default fill and line
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscShapeFormat {
            format: ShapeFormat::new(),
        }
    }

    /// Fill the shape with a solid color. The format keeps a copy.
    #[napi]
    pub fn set_solid_fill(&mut self, fill: &RuscShapeSolidFill) -> &Self {
        self.format = self.format.clone().set_solid_fill(&fill.fill);
        self
    }

    /// Fill the shape with a pattern. The format keeps a copy.
    #[napi]
    pub fn set_pattern_fill(&mut self, fill: &RuscShapePatternFill) -> &Self {
        self.format = self.format.clone().set_pattern_fill(&fill.fill);
        self
    }

    /// Give the shape no fill, so the cells behind it show through
    #[napi]
    pub fn set_no_fill(&mut self) -> &Self {
        self.format = self.format.clone().set_no_fill();
        self
    }

    /// Set the border line of the shape. The format keeps a copy.
    #[napi]
    pub fn set_line(&mut self, line: &RuscShapeLine) -> &Self {
        self.format = self.format.clone().set_line(&line.line);
        self
    }

    /// Give the shape no border line
    #[napi]
    pub fn set_no_line(&mut self) -> &Self {
        self.format = self.format.clone().set_no_line();
        self
    }
}

/// The border line of a shape.
#[napi]
pub struct RuscShapeLine {
    line: ShapeLine,
}

impl Default for RuscShapeLine {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscShapeLine {
    /// Create a shape line: a solid 0.75 pt line in the theme color
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscShapeLine {
            line: ShapeLine::new(),
        }
    }

    /// Set the line color: a color name, `#RRGGBB` or `theme:<color>:<shade>`
    #[napi]
    pub fn set_color(&mut self, color: String) -> Result<&Self> {
        self.line = self.line.clone().set_color(parse_color(&color)?);
        Ok(self)
    }

    /// Set the line width in points, from 0 to 1584. Excel rounds it to the
    /// nearest quarter point.
    #[napi]
    pub fn set_width(&mut self, width: f64) -> Result<&Self> {
        if !(0.0..=1584.0).contains(&width) {
            return Err(Error::new(
                Status::InvalidArg,
                format!("Shape line width must be between 0 and 1584 points, got {width}"),
            ));
        }
        self.line = self.line.clone().set_width(width);
        Ok(self)
    }

    /// Set the dash style of the line
    #[napi]
    pub fn set_dash_type(&mut self, dash_type: String) -> Result<&Self> {
        self.line = self
            .line
            .clone()
            .set_dash_type(parse_dash_type(&dash_type)?);
        Ok(self)
    }

    /// Set the line transparency in percent, from 0 (opaque) to 100
    #[napi]
    pub fn set_transparency(&mut self, transparency: f64) -> Result<&Self> {
        self.line = self
            .line
            .clone()
            .set_transparency(parse_transparency(transparency)?);
        Ok(self)
    }

    /// Hide the line, or show it again
    #[napi]
    pub fn set_hidden(&mut self, hidden: bool) -> &Self {
        self.line = self.line.clone().set_hidden(hidden);
        self
    }
}

/// A solid fill of a shape.
#[napi]
pub struct RuscShapeSolidFill {
    fill: ShapeSolidFill,
}

impl Default for RuscShapeSolidFill {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscShapeSolidFill {
    /// Create a solid fill in the theme color
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscShapeSolidFill {
            fill: ShapeSolidFill::new(),
        }
    }

    /// Set the fill color: a color name, `#RRGGBB` or `theme:<color>:<shade>`
    #[napi]
    pub fn set_color(&mut self, color: String) -> Result<&Self> {
        self.fill = self.fill.clone().set_color(parse_color(&color)?);
        Ok(self)
    }

    /// Set the fill transparency in percent, from 0 (opaque) to 100
    #[napi]
    pub fn set_transparency(&mut self, transparency: f64) -> Result<&Self> {
        self.fill = self
            .fill
            .clone()
            .set_transparency(parse_transparency(transparency)?);
        Ok(self)
    }
}

/// A pattern fill of a shape.
#[napi]
pub struct RuscShapePatternFill {
    fill: ShapePatternFill,
}

impl Default for RuscShapePatternFill {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscShapePatternFill {
    /// Create a pattern fill: `dotted5Percent` in the theme colors
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscShapePatternFill {
            fill: ShapePatternFill::new(),
        }
    }

    /// Set the pattern
    #[napi]
    pub fn set_pattern(&mut self, pattern: String) -> Result<&Self> {
        self.fill = self.fill.clone().set_pattern(parse_pattern(&pattern)?);
        Ok(self)
    }

    /// Set the color of the pattern: a color name, `#RRGGBB` or
    /// `theme:<color>:<shade>`
    #[napi]
    pub fn set_foreground_color(&mut self, color: String) -> Result<&Self> {
        self.fill = self.fill.clone().set_foreground_color(parse_color(&color)?);
        Ok(self)
    }

    /// Set the color behind the pattern: a color name, `#RRGGBB` or
    /// `theme:<color>:<shade>`
    #[napi]
    pub fn set_background_color(&mut self, color: String) -> Result<&Self> {
        self.fill = self.fill.clone().set_background_color(parse_color(&color)?);
        Ok(self)
    }
}

/// The font of shape text.
#[napi]
pub struct RuscShapeFont {
    font: ShapeFont,
}

impl Default for RuscShapeFont {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscShapeFont {
    /// Create a shape font: the theme font at 11 pt
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscShapeFont {
            font: ShapeFont::new(),
        }
    }

    /// Set the font name, e.g. "Arial"
    #[napi]
    pub fn set_name(&mut self, name: String) -> &Self {
        self.font = self.font.clone().set_name(name);
        self
    }

    /// Set the font size in points, from 1 to 4000 in steps of 0.5
    #[napi]
    pub fn set_size(&mut self, size: f64) -> Result<&Self> {
        // DrawingML stores the size as a whole number of hundredths of a
        // point (100 to 400000) and rust_xlsxwriter writes `size * 100`. Half
        // points, the steps of Excel's font size box, give whole numbers
        // there, while other fractions like 8.15 may not (815.0000000000001).
        if !(1.0..=4000.0).contains(&size) || (size * 2.0).fract() != 0.0 {
            return Err(Error::new(
                Status::InvalidArg,
                format!(
                    "Shape font size must be between 1 and 4000 points in steps of 0.5, got {size}"
                ),
            ));
        }
        self.font = self.font.clone().set_size(size);
        Ok(self)
    }

    /// Make the text bold
    #[napi]
    pub fn set_bold(&mut self) -> &Self {
        self.font = self.font.clone().set_bold();
        self
    }

    /// Make the text italic
    #[napi]
    pub fn set_italic(&mut self) -> &Self {
        self.font = self.font.clone().set_italic();
        self
    }

    /// Underline the text with a single line
    #[napi]
    pub fn set_underline(&mut self) -> &Self {
        self.font = self.font.clone().set_underline();
        self
    }

    /// Set the text color: a color name, `#RRGGBB` or `theme:<color>:<shade>`
    #[napi]
    pub fn set_color(&mut self, color: String) -> Result<&Self> {
        self.font = self.font.clone().set_color(parse_color(&color)?);
        Ok(self)
    }
}

/// The alignment and direction of shape text.
#[napi]
pub struct RuscShapeText {
    text: ShapeText,
}

impl Default for RuscShapeText {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
impl RuscShapeText {
    /// Create text options: horizontal text at the top, default alignment
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscShapeText {
            text: ShapeText::new(),
        }
    }

    /// Set the horizontal alignment of the text
    #[napi]
    pub fn set_horizontal_alignment(&mut self, alignment: String) -> Result<&Self> {
        let alignment = match alignment.as_str() {
            "default" => ShapeTextHorizontalAlignment::Default,
            "left" => ShapeTextHorizontalAlignment::Left,
            "center" => ShapeTextHorizontalAlignment::Center,
            "right" => ShapeTextHorizontalAlignment::Right,
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid shape text horizontal alignment: {alignment}. Use 'default', 'left', 'center' or 'right'"
                    ),
                ));
            }
        };
        self.text = self.text.clone().set_horizontal_alignment(alignment);
        Ok(self)
    }

    /// Set the vertical alignment of the text
    #[napi]
    pub fn set_vertical_alignment(&mut self, alignment: String) -> Result<&Self> {
        let alignment = match alignment.as_str() {
            "top" => ShapeTextVerticalAlignment::Top,
            "middle" => ShapeTextVerticalAlignment::Middle,
            "bottom" => ShapeTextVerticalAlignment::Bottom,
            "topCentered" => ShapeTextVerticalAlignment::TopCentered,
            "middleCentered" => ShapeTextVerticalAlignment::MiddleCentered,
            "bottomCentered" => ShapeTextVerticalAlignment::BottomCentered,
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid shape text vertical alignment: {alignment}. Use 'top', 'middle', 'bottom', 'topCentered', 'middleCentered' or 'bottomCentered'"
                    ),
                ));
            }
        };
        self.text = self.text.clone().set_vertical_alignment(alignment);
        Ok(self)
    }

    /// Set the direction of the text
    #[napi]
    pub fn set_direction(&mut self, direction: String) -> Result<&Self> {
        let direction = match direction.as_str() {
            "horizontal" => ShapeTextDirection::Horizontal,
            "rotate90" => ShapeTextDirection::Rotate90,
            "rotate270" => ShapeTextDirection::Rotate270,
            "stacked" => ShapeTextDirection::Stacked,
            "rotate90EastAsian" => ShapeTextDirection::Rotate90EastAsian,
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid shape text direction: {direction}. Use 'horizontal', 'rotate90', 'rotate270', 'stacked' or 'rotate90EastAsian'"
                    ),
                ));
            }
        };
        self.text = self.text.clone().set_direction(direction);
        Ok(self)
    }
}

/// Checks a transparency percentage, which rust_xlsxwriter would otherwise
/// ignore silently when it is over 100.
fn parse_transparency(transparency: f64) -> Result<u8> {
    if transparency.fract() != 0.0 || !(0.0..=100.0).contains(&transparency) {
        return Err(Error::new(
            Status::InvalidArg,
            format!("Transparency must be an integer from 0 to 100, got {transparency}"),
        ));
    }
    Ok(transparency as u8)
}

fn parse_dash_type(dash_type: &str) -> Result<ShapeLineDashType> {
    Ok(match dash_type {
        "solid" => ShapeLineDashType::Solid,
        "roundDot" => ShapeLineDashType::RoundDot,
        "squareDot" => ShapeLineDashType::SquareDot,
        "dash" => ShapeLineDashType::Dash,
        "dashDot" => ShapeLineDashType::DashDot,
        "longDash" => ShapeLineDashType::LongDash,
        "longDashDot" => ShapeLineDashType::LongDashDot,
        "longDashDotDot" => ShapeLineDashType::LongDashDotDot,
        _ => {
            return Err(Error::new(
                Status::InvalidArg,
                format!(
                    "Invalid shape line dash type: {dash_type}. Use 'solid', 'roundDot', 'squareDot', 'dash', 'dashDot', 'longDash', 'longDashDot' or 'longDashDotDot'"
                ),
            ));
        }
    })
}

fn parse_pattern(pattern: &str) -> Result<ShapePatternFillType> {
    use ShapePatternFillType as P;
    Ok(match pattern {
        "dotted5Percent" => P::Dotted5Percent,
        "dotted10Percent" => P::Dotted10Percent,
        "dotted20Percent" => P::Dotted20Percent,
        "dotted25Percent" => P::Dotted25Percent,
        "dotted30Percent" => P::Dotted30Percent,
        "dotted40Percent" => P::Dotted40Percent,
        "dotted50Percent" => P::Dotted50Percent,
        "dotted60Percent" => P::Dotted60Percent,
        "dotted70Percent" => P::Dotted70Percent,
        "dotted75Percent" => P::Dotted75Percent,
        "dotted80Percent" => P::Dotted80Percent,
        "dotted90Percent" => P::Dotted90Percent,
        "diagonalStripesLightDownwards" => P::DiagonalStripesLightDownwards,
        "diagonalStripesLightUpwards" => P::DiagonalStripesLightUpwards,
        "diagonalStripesDarkDownwards" => P::DiagonalStripesDarkDownwards,
        "diagonalStripesDarkUpwards" => P::DiagonalStripesDarkUpwards,
        "diagonalStripesWideDownwards" => P::DiagonalStripesWideDownwards,
        "diagonalStripesWideUpwards" => P::DiagonalStripesWideUpwards,
        "verticalStripesLight" => P::VerticalStripesLight,
        "horizontalStripesLight" => P::HorizontalStripesLight,
        "verticalStripesNarrow" => P::VerticalStripesNarrow,
        "horizontalStripesNarrow" => P::HorizontalStripesNarrow,
        "verticalStripesDark" => P::VerticalStripesDark,
        "horizontalStripesDark" => P::HorizontalStripesDark,
        "stripesBackslashes" => P::StripesBackslashes,
        "stripesForwardSlashes" => P::StripesForwardSlashes,
        "horizontalStripesAlternating" => P::HorizontalStripesAlternating,
        "verticalStripesAlternating" => P::VerticalStripesAlternating,
        "smallConfetti" => P::SmallConfetti,
        "largeConfetti" => P::LargeConfetti,
        "zigzag" => P::Zigzag,
        "wave" => P::Wave,
        "diagonalBrick" => P::DiagonalBrick,
        "horizontalBrick" => P::HorizontalBrick,
        "weave" => P::Weave,
        "plaid" => P::Plaid,
        "divot" => P::Divot,
        "dottedGrid" => P::DottedGrid,
        "dottedDiamond" => P::DottedDiamond,
        "shingle" => P::Shingle,
        "trellis" => P::Trellis,
        "sphere" => P::Sphere,
        "smallGrid" => P::SmallGrid,
        "largeGrid" => P::LargeGrid,
        "smallCheckerboard" => P::SmallCheckerboard,
        "largeCheckerboard" => P::LargeCheckerboard,
        "outlinedDiamondGrid" => P::OutlinedDiamondGrid,
        "solidDiamondGrid" => P::SolidDiamondGrid,
        _ => {
            return Err(Error::new(
                Status::InvalidArg,
                format!("Invalid shape pattern fill type: {pattern}"),
            ));
        }
    })
}
