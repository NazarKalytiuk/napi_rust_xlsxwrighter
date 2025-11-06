use napi::bindgen_prelude::*;
use rust_xlsxwriter::{
    ChartType, Color, DataValidationRule, FormatBorder,
};

/// Helper function to parse validation rules for integers
pub fn parse_validation_rule_int(rule: String, value: i32) -> Result<DataValidationRule<i32>> {
    match rule.as_str() {
        "equal" => Ok(DataValidationRule::EqualTo(value)),
        "notEqual" => Ok(DataValidationRule::NotEqualTo(value)),
        "greaterThan" => Ok(DataValidationRule::GreaterThan(value)),
        "lessThan" => Ok(DataValidationRule::LessThan(value)),
        "greaterThanOrEqual" => Ok(DataValidationRule::GreaterThanOrEqualTo(value)),
        "lessThanOrEqual" => Ok(DataValidationRule::LessThanOrEqualTo(value)),
        _ => Err(Error::new(
            Status::GenericFailure,
            format!(
                "Invalid validation rule: {}. Valid values: 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'",
                rule
            ),
        )),
    }
}

/// Helper function to parse validation rules for floats
pub fn parse_validation_rule_float(rule: String, value: f64) -> Result<DataValidationRule<f64>> {
    match rule.as_str() {
        "equal" => Ok(DataValidationRule::EqualTo(value)),
        "notEqual" => Ok(DataValidationRule::NotEqualTo(value)),
        "greaterThan" => Ok(DataValidationRule::GreaterThan(value)),
        "lessThan" => Ok(DataValidationRule::LessThan(value)),
        "greaterThanOrEqual" => Ok(DataValidationRule::GreaterThanOrEqualTo(value)),
        "lessThanOrEqual" => Ok(DataValidationRule::LessThanOrEqualTo(value)),
        _ => Err(Error::new(
            Status::GenericFailure,
            format!(
                "Invalid validation rule: {}. Valid values: 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'",
                rule
            ),
        )),
    }
}

/// Helper function to validate border styles
pub fn is_valid_border(border: &str) -> bool {
    matches!(border, "thin" | "medium" | "thick" | "double" | "dotted" | "dashed" |
             "dashDot" | "dashDotDot" | "slantDashDot" | "mediumDashed" |
             "mediumDashDot" | "mediumDashDotDot")
}

/// Helper function to parse border strings
pub fn parse_border(border: &str) -> Option<FormatBorder> {
    match border {
        "thin" => Some(FormatBorder::Thin),
        "medium" => Some(FormatBorder::Medium),
        "thick" => Some(FormatBorder::Thick),
        "double" => Some(FormatBorder::Double),
        "dotted" => Some(FormatBorder::Dotted),
        "dashed" => Some(FormatBorder::Dashed),
        "dashDot" => Some(FormatBorder::DashDot),
        "dashDotDot" => Some(FormatBorder::DashDotDot),
        "slantDashDot" => Some(FormatBorder::SlantDashDot),
        "mediumDashed" => Some(FormatBorder::MediumDashed),
        "mediumDashDot" => Some(FormatBorder::MediumDashDot),
        "mediumDashDotDot" => Some(FormatBorder::MediumDashDotDot),
        _ => None,
    }
}

/// Helper function to parse color strings
pub fn parse_color(color: &str) -> Option<Color> {
    match color.to_lowercase().as_str() {
        "black" => Some(Color::Black),
        "blue" => Some(Color::Blue),
        "brown" => Some(Color::Brown),
        "cyan" => Some(Color::Cyan),
        "gray" => Some(Color::Gray),
        "green" => Some(Color::Green),
        "lime" => Some(Color::Lime),
        "magenta" => Some(Color::Magenta),
        "navy" => Some(Color::Navy),
        "orange" => Some(Color::Orange),
        "pink" => Some(Color::Pink),
        "purple" => Some(Color::Purple),
        "red" => Some(Color::Red),
        "silver" => Some(Color::Silver),
        "white" => Some(Color::White),
        "yellow" => Some(Color::Yellow),
        _ => {
            // Try to parse as hex color #RRGGBB
            if color.starts_with('#') && color.len() == 7 {
                if let (Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&color[1..3], 16),
                    u8::from_str_radix(&color[3..5], 16),
                    u8::from_str_radix(&color[5..7], 16),
                ) {
                    let rgb_value: u32 = ((r as u32) << 16) | ((g as u32) << 8) | (b as u32);
                    return Some(Color::RGB(rgb_value));
                }
            }
            None
        }
    }
}

/// Helper function to parse chart type strings
pub fn parse_chart_type(chart_type: &str) -> Result<ChartType> {
    match chart_type.to_lowercase().as_str() {
        "area" => Ok(ChartType::Area),
        "areastacked" | "area_stacked" => Ok(ChartType::AreaStacked),
        "areapercentstacked" | "area_percent_stacked" => Ok(ChartType::AreaPercentStacked),
        "bar" => Ok(ChartType::Bar),
        "barstacked" | "bar_stacked" => Ok(ChartType::BarStacked),
        "barpercentstacked" | "bar_percent_stacked" => Ok(ChartType::BarPercentStacked),
        "column" => Ok(ChartType::Column),
        "columnstacked" | "column_stacked" => Ok(ChartType::ColumnStacked),
        "columnpercentstacked" | "column_percent_stacked" => Ok(ChartType::ColumnPercentStacked),
        "doughnut" => Ok(ChartType::Doughnut),
        "line" => Ok(ChartType::Line),
        "linestacked" | "line_stacked" => Ok(ChartType::LineStacked),
        "linepercentstacked" | "line_percent_stacked" => Ok(ChartType::LinePercentStacked),
        "pie" => Ok(ChartType::Pie),
        "radar" => Ok(ChartType::Radar),
        "radarwithmarkers" | "radar_with_markers" => Ok(ChartType::RadarWithMarkers),
        "radarfilled" | "radar_filled" => Ok(ChartType::RadarFilled),
        "scatter" => Ok(ChartType::Scatter),
        "scatterstraight" | "scatter_straight" => Ok(ChartType::ScatterStraight),
        "scatterstraightwithmarkers" | "scatter_straight_with_markers" => Ok(ChartType::ScatterStraightWithMarkers),
        "scattersmooth" | "scatter_smooth" => Ok(ChartType::ScatterSmooth),
        "scattersmoothwithmarkers" | "scatter_smooth_with_markers" => Ok(ChartType::ScatterSmoothWithMarkers),
        "stock" => Ok(ChartType::Stock),
        _ => Err(Error::new(
            Status::GenericFailure,
            format!("Invalid chart type: {}. Valid types: area, bar, column, line, pie, scatter, doughnut, radar, stock", chart_type),
        )),
    }
}
