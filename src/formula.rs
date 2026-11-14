use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Formula;

/// The data of a JavaScript `Formula`: the formula text and, for applications
/// that don't recalculate, its cached result.
///
/// It is plain data rather than a class so that a formula written per row owns no
/// Rust memory between writes.
#[napi(object)]
pub struct RuscFormulaData {
    pub formula: String,
    pub result: Option<String>,
}

impl From<RuscFormulaData> for Formula {
    fn from(data: RuscFormulaData) -> Self {
        let formula = Formula::new(data.formula);
        match data.result {
            Some(result) => formula.set_result(result),
            None => formula,
        }
    }
}

/// A formula for the native formula writers.
///
/// @deprecated Pass a `RuscFormulaData` object instead. An instance keeps its Rust
/// memory until its finalizer runs after the event loop turns, so instances
/// created in a synchronous loop accumulate.
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
}

/// The `rust_xlsxwriter` formula for a native formula writer argument: a
/// deprecated `RuscFormula` instance or `RuscFormulaData`.
pub(crate) fn to_formula(formula: Either<&RuscFormula, RuscFormulaData>) -> Formula {
    match formula {
        Either::A(formula) => formula.formula.clone(),
        Either::B(data) => data.into(),
    }
}
