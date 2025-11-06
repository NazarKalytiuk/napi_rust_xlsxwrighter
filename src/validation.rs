use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{DataValidation, DataValidationErrorStyle, DataValidationRule, Formula};

use crate::utils::{parse_validation_rule_float, parse_validation_rule_int};

#[napi]
pub struct RuscDataValidation {
    pub(crate) validation: DataValidation,
}

#[napi]
impl RuscDataValidation {
    /// Create a new data validation with "any value" type
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscDataValidation {
            validation: DataValidation::new(),
        }
    }

    /// Allow whole number validation
    #[napi]
    pub fn allow_whole_number(&mut self, rule: String, value: i32) -> Result<()> {
        let rule_enum = parse_validation_rule_int(rule, value)?;
        self.validation = DataValidation::new().allow_whole_number(rule_enum);
        Ok(())
    }

    /// Allow whole number validation between two values
    #[napi]
    pub fn allow_whole_number_between(&mut self, min: i32, max: i32) -> Result<()> {
        self.validation = DataValidation::new().allow_whole_number(DataValidationRule::Between(min, max));
        Ok(())
    }

    /// Allow decimal number validation
    #[napi]
    pub fn allow_decimal_number(&mut self, rule: String, value: f64) -> Result<()> {
        let rule_enum = parse_validation_rule_float(rule, value)?;
        self.validation = DataValidation::new().allow_decimal_number(rule_enum);
        Ok(())
    }

    /// Allow decimal number validation between two values
    #[napi]
    pub fn allow_decimal_number_between(&mut self, min: f64, max: f64) -> Result<()> {
        self.validation = DataValidation::new().allow_decimal_number(DataValidationRule::Between(min, max));
        Ok(())
    }

    /// Allow list validation with strings
    #[napi]
    pub fn allow_list_strings(&mut self, list: Vec<String>) -> Result<()> {
        let string_refs: Vec<&str> = list.iter().map(|s| s.as_str()).collect();
        self.validation = DataValidation::new()
            .allow_list_strings(&string_refs)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set list: {}", e)))?;
        Ok(())
    }

    /// Allow list validation from a formula/cell range
    #[napi]
    pub fn allow_list_formula(&mut self, formula: String) -> Result<()> {
        self.validation = DataValidation::new().allow_list_formula(Formula::new(formula));
        Ok(())
    }

    /// Allow custom formula validation
    #[napi]
    pub fn allow_custom(&mut self, formula: String) -> Result<()> {
        self.validation = DataValidation::new().allow_custom(Formula::new(formula));
        Ok(())
    }

    /// Allow any value (mainly used for input messages)
    #[napi]
    pub fn allow_any_value(&mut self) -> Result<()> {
        self.validation = DataValidation::new().allow_any_value();
        Ok(())
    }

    /// Set whether to ignore blank cells
    #[napi]
    pub fn ignore_blank(&mut self, enable: bool) -> Result<()> {
        self.validation = self.validation.clone().ignore_blank(enable);
        Ok(())
    }

    /// Set whether to show the dropdown arrow
    #[napi]
    pub fn show_dropdown(&mut self, enable: bool) -> Result<()> {
        self.validation = self.validation.clone().show_dropdown(enable);
        Ok(())
    }

    /// Set whether to show the input message
    #[napi]
    pub fn show_input_message(&mut self, enable: bool) -> Result<()> {
        self.validation = self.validation.clone().show_input_message(enable);
        Ok(())
    }

    /// Set whether to show the error message
    #[napi]
    pub fn show_error_message(&mut self, enable: bool) -> Result<()> {
        self.validation = self.validation.clone().show_error_message(enable);
        Ok(())
    }

    /// Set the input message title
    #[napi]
    pub fn set_input_title(&mut self, title: String) -> Result<()> {
        if title.chars().count() > 32 {
            return Err(Error::new(
                Status::GenericFailure,
                "Input title must be 32 characters or less".to_string(),
            ));
        }
        self.validation = self.validation.clone()
            .set_input_title(title)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set input title: {}", e)))?;
        Ok(())
    }

    /// Set the input message text
    #[napi]
    pub fn set_input_message(&mut self, message: String) -> Result<()> {
        if message.chars().count() > 255 {
            return Err(Error::new(
                Status::GenericFailure,
                "Input message must be 255 characters or less".to_string(),
            ));
        }
        self.validation = self.validation.clone()
            .set_input_message(message)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set input message: {}", e)))?;
        Ok(())
    }

    /// Set the error message title
    #[napi]
    pub fn set_error_title(&mut self, title: String) -> Result<()> {
        if title.chars().count() > 32 {
            return Err(Error::new(
                Status::GenericFailure,
                "Error title must be 32 characters or less".to_string(),
            ));
        }
        self.validation = self.validation.clone()
            .set_error_title(title)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set error title: {}", e)))?;
        Ok(())
    }

    /// Set the error message text
    #[napi]
    pub fn set_error_message(&mut self, message: String) -> Result<()> {
        if message.chars().count() > 255 {
            return Err(Error::new(
                Status::GenericFailure,
                "Error message must be 255 characters or less".to_string(),
            ));
        }
        self.validation = self.validation.clone()
            .set_error_message(message)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set error message: {}", e)))?;
        Ok(())
    }

    /// Set the error style
    #[napi]
    pub fn set_error_style(&mut self, style: String) -> Result<()> {
        let style_enum = match style.as_str() {
            "stop" => DataValidationErrorStyle::Stop,
            "warning" => DataValidationErrorStyle::Warning,
            "information" => DataValidationErrorStyle::Information,
            _ => {
                return Err(Error::new(
                    Status::GenericFailure,
                    format!("Invalid error style: {}. Valid values: 'stop', 'warning', 'information'", style),
                ))
            }
        };
        self.validation = self.validation.clone().set_error_style(style_enum);
        Ok(())
    }

    /// Get the internal DataValidation object
    pub(crate) fn get_validation(&self) -> &DataValidation {
        &self.validation
    }
}

// Helper function to parse validation rules for integers
