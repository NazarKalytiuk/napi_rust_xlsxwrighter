use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Formula;

#[napi]
pub struct RuscFormula {
    pub(crate) formula: Formula,
}

#[napi]
impl RuscFormula {
    /// Create a new formula
    #[napi(constructor)]
    pub fn new(formula: String) -> Self {
        RuscFormula {
            formula: Formula::new(formula),
        }
    }

    /// Set the result value for the formula (for apps that don't auto-calculate)
    #[napi]
    pub fn set_result(&mut self, result: String) -> Result<()> {
        self.formula = self.formula.clone().set_result(result);
        Ok(())
    }

    /// Get the internal Formula object
    pub(crate) fn get_formula(&self) -> &Formula {
        &self.formula
    }
}

