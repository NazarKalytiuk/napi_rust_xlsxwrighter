use napi::bindgen_prelude::*;
use rust_xlsxwriter::{
    ChartType, Color, DataValidationRule, FormatBorder, IgnoreError, IntoDataValidationValue,
};

/// Parses a single-value `ValidationRule` of the JavaScript API, comparing
/// cell input with `value`.
pub fn parse_validation_rule<T: IntoDataValidationValue>(
    rule: &str,
    value: T,
) -> Result<DataValidationRule<T>> {
    match rule {
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

/// Parses a `BorderStyle` of the JavaScript API.
pub fn parse_border(border: &str) -> Result<FormatBorder> {
    Ok(match border {
        "thin" => FormatBorder::Thin,
        "medium" => FormatBorder::Medium,
        "thick" => FormatBorder::Thick,
        "double" => FormatBorder::Double,
        "dotted" => FormatBorder::Dotted,
        "dashed" => FormatBorder::Dashed,
        "dashDot" => FormatBorder::DashDot,
        "dashDotDot" => FormatBorder::DashDotDot,
        "slantDashDot" => FormatBorder::SlantDashDot,
        "mediumDashed" => FormatBorder::MediumDashed,
        "mediumDashDot" => FormatBorder::MediumDashDot,
        "mediumDashDotDot" => FormatBorder::MediumDashDotDot,
        "hair" => FormatBorder::Hair,
        _ => {
            return Err(Error::new(
                Status::GenericFailure,
                format!("Invalid border style: {border}"),
            ));
        }
    })
}

/// Parses a `Color` of the JavaScript API: a named color, `#RRGGBB`, or a
/// theme palette color `theme:<color>:<shade>`.
///
/// Excel writes theme colors only where a color is part of a cell style or
/// the worksheet tab, so other colors are parsed with [`parse_rgb_color`].
pub fn parse_color(color: &str) -> Result<Color> {
    parse_labeled_color("color", color)
}

/// Parses a named or `#RRGGBB` color for a property that rust_xlsxwriter
/// stores as an RGB value, such as a note background or a conditional format
/// color, where it would silently replace a theme color with a default one.
/// `label` names the color in errors, e.g. `"min color"`.
pub fn parse_rgb_color(label: &str, color: &str) -> Result<Color> {
    match parse_labeled_color(label, color)? {
        Color::Theme(_, _) => Err(invalid_color(
            label,
            color,
            "Theme colors aren't supported here; use a named or #RRGGBB color",
        )),
        rgb => Ok(rgb),
    }
}

fn parse_labeled_color(label: &str, color: &str) -> Result<Color> {
    if let Some(theme) = color.strip_prefix("theme:") {
        return parse_theme_color(theme).ok_or_else(|| {
            invalid_color(
                label,
                color,
                "Theme colors are 'theme:<color>:<shade>' with a color of 0-9 and a shade of 0-5",
            )
        });
    }
    let named = match color.to_lowercase().as_str() {
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
        _ => None,
    };
    named
        .or_else(|| parse_hex_color(color))
        .ok_or_else(|| Error::new(Status::GenericFailure, format!("Invalid {label}: {color}")))
}

fn invalid_color(label: &str, color: &str, detail: &str) -> Error {
    Error::new(
        Status::GenericFailure,
        format!("Invalid {label}: {color}. {detail}"),
    )
}

/// Parses the `#RRGGBB` form of a color.
fn parse_hex_color(color: &str) -> Option<Color> {
    let hex = color.strip_prefix('#')?;
    if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    u32::from_str_radix(hex, 16).ok().map(Color::RGB)
}

/// Parses the `<color>:<shade>` part of a `theme:<color>:<shade>` color.
fn parse_theme_color(theme: &str) -> Option<Color> {
    let index = |part: Option<&str>, max: u8| {
        part.filter(|digits| !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()))
            .and_then(|digits| digits.parse::<u8>().ok())
            .filter(|value| *value <= max)
    };
    let mut parts = theme.split(':');
    match (index(parts.next(), 9), index(parts.next(), 5), parts.next()) {
        (Some(color), Some(shade), None) => Some(Color::Theme(color, shade)),
        _ => None,
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
        "scatterstraightwithmarkers" | "scatter_straight_with_markers" => {
            Ok(ChartType::ScatterStraightWithMarkers)
        }
        "scattersmooth" | "scatter_smooth" => Ok(ChartType::ScatterSmooth),
        "scattersmoothwithmarkers" | "scatter_smooth_with_markers" => {
            Ok(ChartType::ScatterSmoothWithMarkers)
        }
        "stock" => Ok(ChartType::Stock),
        _ => Err(Error::new(
            Status::GenericFailure,
            format!(
                "Invalid chart type: {}. Valid types: area, bar, column, line, pie, scatter, doughnut, radar, stock",
                chart_type
            ),
        )),
    }
}

/// Parses an `IgnoreErrorType` of the JavaScript API.
pub fn parse_ignore_error(error_type: &str) -> Result<IgnoreError> {
    Ok(match error_type {
        "numberStoredAsText" => IgnoreError::NumberStoredAsText,
        "formulaError" => IgnoreError::FormulaError,
        "formulaDiffers" => IgnoreError::FormulaDiffers,
        "formulaRefersToEmptyCells" => IgnoreError::FormulaRefersToEmptyCells,
        "formulaOmitsCells" => IgnoreError::FormulaOmitsCells,
        "dataValidationError" => IgnoreError::DataValidationError,
        "unlockedCellsWithFormula" => IgnoreError::UnlockedCellsWithFormula,
        "inconsistentColumnFormula" => IgnoreError::InconsistentColumnFormula,
        "twoDigitTextYear" => IgnoreError::TwoDigitTextYear,
        _ => {
            return Err(Error::new(
                Status::InvalidArg,
                format!("Invalid ignore error type: {error_type}"),
            ));
        }
    })
}
