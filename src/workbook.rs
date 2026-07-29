use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::Workbook;
use parking_lot::Mutex;
use std::sync::Arc;

use crate::properties::RuscDocProperties;
use crate::worksheet::RuscWorksheet;

/// Workbook wrapper for Excel file creation
#[napi]
pub struct RuscWorkbook {
    pub(crate) workbook: Arc<Mutex<Workbook>>,
}

#[napi]
impl RuscWorkbook {
    /// Create a new workbook
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscWorkbook {
            workbook: Arc::new(Mutex::new(Workbook::new())),
        }
    }

    /// Add a new worksheet to the workbook and return a reference to it
    #[napi]
    pub fn add_worksheet(&self, name: Option<String>) -> Result<RuscWorksheet> {
        let workbook = Arc::clone(&self.workbook);
        let index = {
            let mut wb = workbook.lock();
            let worksheet = wb.add_worksheet();

            if let Some(n) = name {
                worksheet.set_name(&n)
                    .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set worksheet name: {}", e)))?;
            }

            wb.worksheets().len() - 1
        };

        Ok(RuscWorksheet {
            workbook,
            index,
        })
    }

    /// Add a new worksheet with constant memory mode to the workbook
    #[napi]
    pub fn add_worksheet_with_constant_memory(&self, name: Option<String>) -> Result<RuscWorksheet> {
        let workbook = Arc::clone(&self.workbook);
        let index = {
            let mut wb = workbook.lock();
            let worksheet = wb.add_worksheet_with_constant_memory();

            if let Some(n) = name {
                worksheet.set_name(&n)
                    .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set worksheet name: {}", e)))?;
            }

            wb.worksheets().len() - 1
        };

        Ok(RuscWorksheet {
            workbook,
            index,
        })
    }

    /// Add a new worksheet with low memory mode to the workbook
    #[napi]
    pub fn add_worksheet_with_low_memory(&self, name: Option<String>) -> Result<RuscWorksheet> {
        let workbook = Arc::clone(&self.workbook);
        let index = {
            let mut wb = workbook.lock();
            let worksheet = wb.add_worksheet_with_low_memory();

            if let Some(n) = name {
                worksheet.set_name(&n)
                    .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set worksheet name: {}", e)))?;
            }

            wb.worksheets().len() - 1
        };

        Ok(RuscWorksheet {
            workbook,
            index,
        })
    }

    /// Set a custom temporary directory for constant memory mode
    #[napi]
    pub fn set_tempdir(&self, dir: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        wb.set_tempdir(&dir)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to set tempdir: {}", e)))?;
        Ok(())
    }

    /// Save the workbook to a file
    #[napi]
    pub fn save(&self, filename: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        wb.save(&filename)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to save workbook: {}", e)))?;
        Ok(())
    }

    /// Save the workbook to a buffer
    #[napi]
    pub fn save_to_buffer(&self) -> Result<Buffer> {
        let mut wb = self.workbook.lock();
        let buffer = wb.save_to_buffer()
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to save to buffer: {}", e)))?;

        Ok(Buffer::from(buffer))
    }

    /// Define a named range or formula
    ///
    /// # Arguments
    ///
    /// * `name` - The name for the range (e.g., "Sales" or "Sheet1!Sales" for local names)
    /// * `formula` - The formula or range (e.g., "=Sheet1!$A$1:$A$10" or "=SUM(A1:A10)")
    #[napi]
    pub fn define_name(&self, name: String, formula: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        wb.define_name(&name, &formula)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to define name: {}", e)))?;
        Ok(())
    }

    /// Set the workbook to be opened in read-only recommended mode
    #[napi]
    pub fn read_only_recommended(&self) -> Result<()> {
        let mut wb = self.workbook.lock();
        wb.read_only_recommended();
        Ok(())
    }

    /// Add a VBA macro file to the workbook
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the vbaProject.bin file
    #[napi]
    pub fn add_vba_project(&self, path: String) -> Result<()> {
        let mut wb = self.workbook.lock();
        wb.add_vba_project(&path)
            .map_err(|e| Error::new(Status::GenericFailure, format!("Failed to add VBA project: {}", e)))?;
        Ok(())
    }

    /// Set document properties (metadata) for the workbook
    ///
    /// # Arguments
    ///
    /// * `properties` - DocProperties object with metadata
    #[napi]
    pub fn set_properties(&self, properties: &RuscDocProperties) -> Result<()> {
        let mut wb = self.workbook.lock();
        wb.set_properties(properties.get_properties());
        Ok(())
    }
}
