// Sparkline module for Excel sparklines
use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{ChartEmptyCells, Color, Sparkline, SparklineType};

#[napi]
pub struct RuscSparkline {
    pub(crate) sparkline: Sparkline,
}

#[napi]
impl RuscSparkline {
    /// Create a new Sparkline
    #[napi(factory)]
    pub fn new() -> Self {
        RuscSparkline {
            sparkline: Sparkline::new(),
        }
    }

    /// Set the data range for the sparkline (e.g., "Sheet1!A1:E1")
    #[napi]
    pub fn set_range(&mut self, range: String) -> Result<&Self> {
        // Parse the range string (format: "Sheet!A1:B10" or just "A1:B10")
        let parts: Vec<&str> = if range.contains('!') {
            range.split('!').collect()
        } else {
            vec!["Sheet1", &range]
        };

        if parts.len() != 2 {
            return Err(Error::new(
                Status::InvalidArg,
                "Invalid range format. Use 'Sheet!A1:B10' or 'A1:B10'",
            ));
        }

        let sheet_name = parts[0];
        let cell_range = parts[1];

        // Parse cell range (e.g., "A1:E1")
        let cells: Vec<&str> = cell_range.split(':').collect();
        if cells.len() != 2 {
            return Err(Error::new(
                Status::InvalidArg,
                "Invalid cell range. Use 'A1:E1' format",
            ));
        }

        let (row1, col1) = Self::parse_cell_reference(cells[0])?;
        let (row2, col2) = Self::parse_cell_reference(cells[1])?;

        self.sparkline = self
            .sparkline
            .clone()
            .set_range((sheet_name, row1, col1, row2, col2));
        Ok(self)
    }

    /// Helper function to parse cell reference (e.g., "A1" -> (0, 0))
    fn parse_cell_reference(cell: &str) -> Result<(u32, u16)> {
        let mut col = 0u16;
        let mut row_str = String::new();

        for ch in cell.chars() {
            if ch.is_ascii_uppercase() {
                col = col * 26 + (ch as u16 - 'A' as u16 + 1);
            } else if ch.is_ascii_digit() {
                row_str.push(ch);
            } else {
                return Err(Error::new(Status::InvalidArg, "Invalid cell reference"));
            }
        }

        if row_str.is_empty() {
            return Err(Error::new(Status::InvalidArg, "Missing row number"));
        }

        let row: u32 = row_str
            .parse()
            .map_err(|_| Error::new(Status::InvalidArg, "Invalid row number"))?;

        Ok((row - 1, col - 1)) // Convert to 0-based indexing
    }

    /// Set the type of sparkline: 'line', 'column', or 'winLose'
    #[napi]
    pub fn set_type(&mut self, sparkline_type: String) -> Result<&Self> {
        let spark_type = match sparkline_type.to_lowercase().as_str() {
            "line" => SparklineType::Line,
            "column" => SparklineType::Column,
            "winlose" | "win_lose" => SparklineType::WinLose,
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid sparkline type: {}. Use 'line', 'column', or 'winLose'",
                        sparkline_type
                    ),
                ))
            }
        };
        self.sparkline = self.sparkline.clone().set_type(spark_type);
        Ok(self)
    }

    /// Show markers for the highest point
    #[napi]
    pub fn show_high_point(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_high_point(enable);
        self
    }

    /// Show markers for the lowest point
    #[napi]
    pub fn show_low_point(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_low_point(enable);
        self
    }

    /// Show markers for the first point
    #[napi]
    pub fn show_first_point(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_first_point(enable);
        self
    }

    /// Show markers for the last point
    #[napi]
    pub fn show_last_point(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_last_point(enable);
        self
    }

    /// Show markers for negative points
    #[napi]
    pub fn show_negative_points(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_negative_points(enable);
        self
    }

    /// Show markers for all points
    #[napi]
    pub fn show_markers(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_markers(enable);
        self
    }

    /// Show the horizontal axis
    #[napi]
    pub fn show_axis(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_axis(enable);
        self
    }

    /// Show data from hidden rows/columns
    #[napi]
    pub fn show_hidden_data(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().show_hidden_data(enable);
        self
    }

    /// Display sparkline right to left
    #[napi]
    pub fn set_right_to_left(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().set_right_to_left(enable);
        self
    }

    /// Set the color of the sparkline (hex color like "#FF0000")
    #[napi]
    pub fn set_sparkline_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self.sparkline.clone().set_sparkline_color(parsed_color);
        Ok(self)
    }

    /// Set the color of the high point marker
    #[napi]
    pub fn set_high_point_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self.sparkline.clone().set_high_point_color(parsed_color);
        Ok(self)
    }

    /// Set the color of the low point marker
    #[napi]
    pub fn set_low_point_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self.sparkline.clone().set_low_point_color(parsed_color);
        Ok(self)
    }

    /// Set the color of the first point marker
    #[napi]
    pub fn set_first_point_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self.sparkline.clone().set_first_point_color(parsed_color);
        Ok(self)
    }

    /// Set the color of the last point marker
    #[napi]
    pub fn set_last_point_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self.sparkline.clone().set_last_point_color(parsed_color);
        Ok(self)
    }

    /// Set the color of negative point markers
    #[napi]
    pub fn set_negative_points_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self
            .sparkline
            .clone()
            .set_negative_points_color(parsed_color);
        Ok(self)
    }

    /// Set the color of all point markers
    #[napi]
    pub fn set_markers_color(&mut self, color: String) -> Result<&Self> {
        let parsed_color = Self::parse_color(&color)?;
        self.sparkline = self.sparkline.clone().set_markers_color(parsed_color);
        Ok(self)
    }

    /// Set the line weight/width
    #[napi]
    pub fn set_line_weight(&mut self, weight: f64) -> &Self {
        self.sparkline = self.sparkline.clone().set_line_weight(weight);
        self
    }

    /// Set the custom maximum value for vertical axis
    #[napi]
    pub fn set_custom_max(&mut self, max: f64) -> &Self {
        self.sparkline = self.sparkline.clone().set_custom_max(max);
        self
    }

    /// Set the custom minimum value for vertical axis
    #[napi]
    pub fn set_custom_min(&mut self, min: f64) -> &Self {
        self.sparkline = self.sparkline.clone().set_custom_min(min);
        self
    }

    /// Set maximum value based on group maximum
    #[napi]
    pub fn set_group_max(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().set_group_max(enable);
        self
    }

    /// Set minimum value based on group minimum
    #[napi]
    pub fn set_group_min(&mut self, enable: bool) -> &Self {
        self.sparkline = self.sparkline.clone().set_group_min(enable);
        self
    }

    /// Set a predefined style (1-36)
    #[napi]
    pub fn set_style(&mut self, style: u8) -> &Self {
        self.sparkline = self.sparkline.clone().set_style(style);
        self
    }

    /// Set how empty cells are displayed: 'gaps', 'zero', or 'connected'
    #[napi]
    pub fn show_empty_cells_as(&mut self, option: String) -> Result<&Self> {
        let empty_cells = match option.to_lowercase().as_str() {
            "gaps" => ChartEmptyCells::Gaps,
            "zero" => ChartEmptyCells::Zero,
            "connected" => ChartEmptyCells::Connected,
            _ => {
                return Err(Error::new(
                    Status::InvalidArg,
                    format!(
                        "Invalid empty cells option: {}. Use 'gaps', 'zero', or 'connected'",
                        option
                    ),
                ))
            }
        };
        self.sparkline = self.sparkline.clone().show_empty_cells_as(empty_cells);
        Ok(self)
    }

    /// Helper function to parse color from hex string
    fn parse_color(color_str: &str) -> Result<Color> {
        if color_str.starts_with('#') && color_str.len() == 7 {
            let r = u8::from_str_radix(&color_str[1..3], 16)
                .map_err(|_| Error::new(Status::InvalidArg, "Invalid hex color"))?;
            let g = u8::from_str_radix(&color_str[3..5], 16)
                .map_err(|_| Error::new(Status::InvalidArg, "Invalid hex color"))?;
            let b = u8::from_str_radix(&color_str[5..7], 16)
                .map_err(|_| Error::new(Status::InvalidArg, "Invalid hex color"))?;
            let color_value = ((r as u32) << 16) | ((g as u32) << 8) | (b as u32);
            Ok(Color::RGB(color_value))
        } else {
            Err(Error::new(
                Status::InvalidArg,
                "Color must be in hex format (#RRGGBB)",
            ))
        }
    }

    /// Get the internal sparkline reference
    pub(crate) fn get_sparkline(&self) -> &Sparkline {
        &self.sparkline
    }
}
