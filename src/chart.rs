use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Chart;

use crate::utils::parse_chart_type;

/// Chart wrapper for Excel charts
#[napi]
pub struct RuscChart {
    pub(crate) chart: Chart,
}

#[napi]
impl RuscChart {
    /// Create a new chart with the specified type
    #[napi(constructor)]
    pub fn new(chart_type: String) -> Result<Self> {
        let chart_type_enum = parse_chart_type(&chart_type)?;
        Ok(RuscChart {
            chart: Chart::new(chart_type_enum),
        })
    }

    /// Add a data series to the chart
    #[napi]
    pub fn add_series(
        &mut self,
        categories: Option<String>,
        values: String,
        name: Option<String>,
    ) -> Result<()> {
        let series = self.chart.add_series();

        if let Some(cat) = categories {
            series.set_categories(&*cat);
        }

        series.set_values(&*values);

        if let Some(n) = name {
            series.set_name(&*n);
        }

        Ok(())
    }

    /// Set the chart title
    #[napi]
    pub fn set_title(&mut self, title: String) -> Result<()> {
        self.chart.title().set_name(&title);
        Ok(())
    }

    /// Set the X-axis label
    #[napi]
    pub fn set_x_axis_name(&mut self, name: String) -> Result<()> {
        self.chart.x_axis().set_name(&name);
        Ok(())
    }

    /// Set the Y-axis label
    #[napi]
    pub fn set_y_axis_name(&mut self, name: String) -> Result<()> {
        self.chart.y_axis().set_name(&name);
        Ok(())
    }

    /// Set the chart style (1-48)
    #[napi]
    pub fn set_style(&mut self, style: u8) -> Result<()> {
        if style < 1 || style > 48 {
            return Err(Error::new(
                Status::GenericFailure,
                "Chart style must be between 1 and 48".to_string(),
            ));
        }
        self.chart.set_style(style);
        Ok(())
    }

    /// Hide the chart legend
    #[napi]
    pub fn set_legend_hidden(&mut self) -> Result<()> {
        self.chart.legend().set_hidden();
        Ok(())
    }
}
