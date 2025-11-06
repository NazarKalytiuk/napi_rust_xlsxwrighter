use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::ProtectionOptions;

/// Wrapper for ProtectionOptions to configure worksheet protection
#[napi]
pub struct RuscProtectionOptions {
    pub(crate) options: ProtectionOptions,
}

#[napi]
impl RuscProtectionOptions {
    /// Create a new ProtectionOptions object with default settings
    #[napi(constructor)]
    pub fn new() -> Self {
        RuscProtectionOptions {
            options: ProtectionOptions::new(),
        }
    }

    /// Allow or prevent users from selecting locked cells (default: true)
    #[napi]
    pub fn set_select_locked_cells(&mut self, allow: bool) -> &Self {
        self.options.select_locked_cells = allow;
        self
    }

    /// Allow or prevent users from selecting unlocked cells (default: true)
    #[napi]
    pub fn set_select_unlocked_cells(&mut self, allow: bool) -> &Self {
        self.options.select_unlocked_cells = allow;
        self
    }

    /// Allow or prevent users from formatting cells (default: false)
    #[napi]
    pub fn set_format_cells(&mut self, allow: bool) -> &Self {
        self.options.format_cells = allow;
        self
    }

    /// Allow or prevent users from formatting columns (default: false)
    #[napi]
    pub fn set_format_columns(&mut self, allow: bool) -> &Self {
        self.options.format_columns = allow;
        self
    }

    /// Allow or prevent users from formatting rows (default: false)
    #[napi]
    pub fn set_format_rows(&mut self, allow: bool) -> &Self {
        self.options.format_rows = allow;
        self
    }

    /// Allow or prevent users from inserting columns (default: false)
    #[napi]
    pub fn set_insert_columns(&mut self, allow: bool) -> &Self {
        self.options.insert_columns = allow;
        self
    }

    /// Allow or prevent users from inserting rows (default: false)
    #[napi]
    pub fn set_insert_rows(&mut self, allow: bool) -> &Self {
        self.options.insert_rows = allow;
        self
    }

    /// Allow or prevent users from inserting hyperlinks (default: false)
    #[napi]
    pub fn set_insert_links(&mut self, allow: bool) -> &Self {
        self.options.insert_links = allow;
        self
    }

    /// Allow or prevent users from deleting columns (default: false)
    #[napi]
    pub fn set_delete_columns(&mut self, allow: bool) -> &Self {
        self.options.delete_columns = allow;
        self
    }

    /// Allow or prevent users from deleting rows (default: false)
    #[napi]
    pub fn set_delete_rows(&mut self, allow: bool) -> &Self {
        self.options.delete_rows = allow;
        self
    }

    /// Allow or prevent users from sorting data (default: false)
    #[napi]
    pub fn set_sort(&mut self, allow: bool) -> &Self {
        self.options.sort = allow;
        self
    }

    /// Allow or prevent users from using autofilters (default: false)
    #[napi]
    pub fn set_use_autofilter(&mut self, allow: bool) -> &Self {
        self.options.use_autofilter = allow;
        self
    }

    /// Allow or prevent users from using pivot tables (default: false)
    #[napi]
    pub fn set_use_pivot_tables(&mut self, allow: bool) -> &Self {
        self.options.use_pivot_tables = allow;
        self
    }

    /// Allow or prevent users from editing scenarios (default: false)
    #[napi]
    pub fn set_edit_scenarios(&mut self, allow: bool) -> &Self {
        self.options.edit_scenarios = allow;
        self
    }

    /// Allow or prevent users from editing objects like images (default: false)
    #[napi]
    pub fn set_edit_objects(&mut self, allow: bool) -> &Self {
        self.options.edit_objects = allow;
        self
    }

    /// Allow or prevent editing contents in chartsheets (default: true)
    #[napi]
    pub fn set_contents(&mut self, allow: bool) -> &Self {
        self.options.contents = allow;
        self
    }

    pub(crate) fn get_options(&self) -> &ProtectionOptions {
        &self.options
    }
}
