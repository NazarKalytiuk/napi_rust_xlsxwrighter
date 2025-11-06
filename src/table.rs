// Table module for Excel tables
use napi::bindgen_prelude::*;
use napi_derive::napi;
use rust_xlsxwriter::{Table, TableColumn, TableFunction, TableStyle};

#[napi]
pub struct RuscTable {
    pub(crate) table: Table,
}

#[napi]
impl RuscTable {
    /// Create a new Table
    #[napi(factory)]
    pub fn new() -> Self {
        RuscTable {
            table: Table::new(),
        }
    }

    /// Set whether the table has a header row
    #[napi]
    pub fn set_header_row(&mut self, show: bool) -> &Self {
        self.table = self.table.clone().set_header_row(show);
        self
    }

    /// Set whether the table has a total row
    #[napi]
    pub fn set_total_row(&mut self, show: bool) -> &Self {
        self.table = self.table.clone().set_total_row(show);
        self
    }

    /// Set whether the table has banded rows
    #[napi]
    pub fn set_banded_rows(&mut self, show: bool) -> &Self {
        self.table = self.table.clone().set_banded_rows(show);
        self
    }

    /// Set whether the table has banded columns
    #[napi]
    pub fn set_banded_columns(&mut self, show: bool) -> &Self {
        self.table = self.table.clone().set_banded_columns(show);
        self
    }

    /// Set whether the first column is formatted
    #[napi]
    pub fn set_first_column(&mut self, show: bool) -> &Self {
        self.table = self.table.clone().set_first_column(show);
        self
    }

    /// Set whether the last column is formatted
    #[napi]
    pub fn set_last_column(&mut self, show: bool) -> &Self {
        self.table = self.table.clone().set_last_column(show);
        self
    }

    /// Set the name of the table
    #[napi]
    pub fn set_name(&mut self, name: String) -> &Self {
        self.table = self.table.clone().set_name(&name);
        self
    }

    /// Set the style of the table (1-21 for light styles, 22-49 for medium styles, 50-61 for dark styles)
    #[napi]
    pub fn set_style(&mut self, style_id: u8) -> &Self {
        let style = match style_id {
            1 => TableStyle::Light1,
            2 => TableStyle::Light2,
            3 => TableStyle::Light3,
            4 => TableStyle::Light4,
            5 => TableStyle::Light5,
            6 => TableStyle::Light6,
            7 => TableStyle::Light7,
            8 => TableStyle::Light8,
            9 => TableStyle::Light9,
            10 => TableStyle::Light10,
            11 => TableStyle::Light11,
            12 => TableStyle::Light12,
            13 => TableStyle::Light13,
            14 => TableStyle::Light14,
            15 => TableStyle::Light15,
            16 => TableStyle::Light16,
            17 => TableStyle::Light17,
            18 => TableStyle::Light18,
            19 => TableStyle::Light19,
            20 => TableStyle::Light20,
            21 => TableStyle::Light21,
            22 => TableStyle::Medium1,
            23 => TableStyle::Medium2,
            24 => TableStyle::Medium3,
            25 => TableStyle::Medium4,
            26 => TableStyle::Medium5,
            27 => TableStyle::Medium6,
            28 => TableStyle::Medium7,
            29 => TableStyle::Medium8,
            30 => TableStyle::Medium9,
            31 => TableStyle::Medium10,
            32 => TableStyle::Medium11,
            33 => TableStyle::Medium12,
            34 => TableStyle::Medium13,
            35 => TableStyle::Medium14,
            36 => TableStyle::Medium15,
            37 => TableStyle::Medium16,
            38 => TableStyle::Medium17,
            39 => TableStyle::Medium18,
            40 => TableStyle::Medium19,
            41 => TableStyle::Medium20,
            42 => TableStyle::Medium21,
            43 => TableStyle::Medium22,
            44 => TableStyle::Medium23,
            45 => TableStyle::Medium24,
            46 => TableStyle::Medium25,
            47 => TableStyle::Medium26,
            48 => TableStyle::Medium27,
            49 => TableStyle::Medium28,
            50 => TableStyle::Dark1,
            51 => TableStyle::Dark2,
            52 => TableStyle::Dark3,
            53 => TableStyle::Dark4,
            54 => TableStyle::Dark5,
            55 => TableStyle::Dark6,
            56 => TableStyle::Dark7,
            57 => TableStyle::Dark8,
            58 => TableStyle::Dark9,
            59 => TableStyle::Dark10,
            60 => TableStyle::Dark11,
            _ => TableStyle::Medium9, // Default
        };
        self.table = self.table.clone().set_style(style);
        self
    }

    /// Add columns to the table
    #[napi]
    pub fn set_columns(&mut self, columns: Vec<&RuscTableColumn>) -> &Self {
        let table_columns: Vec<TableColumn> = columns.iter().map(|c| c.column.clone()).collect();
        self.table = self.table.clone().set_columns(&table_columns);
        self
    }

    /// Get the internal table reference
    pub(crate) fn get_table(&self) -> &Table {
        &self.table
    }
}

#[napi]
pub struct RuscTableColumn {
    pub(crate) column: TableColumn,
}

#[napi]
impl RuscTableColumn {
    /// Create a new table column
    #[napi(factory)]
    pub fn new() -> Self {
        RuscTableColumn {
            column: TableColumn::new(),
        }
    }

    /// Set the header name for the column
    #[napi]
    pub fn set_header(&mut self, header: String) -> &Self {
        self.column = self.column.clone().set_header(&header);
        self
    }

    /// Set the total function for the column
    #[napi]
    pub fn set_total_function(&mut self, function: String) -> Result<&Self> {
        let func = match function.as_str() {
            "average" => TableFunction::Average,
            "count" => TableFunction::Count,
            "countNumbers" | "countNums" => TableFunction::CountNumbers,
            "max" | "maximum" => TableFunction::Max,
            "min" | "minimum" => TableFunction::Min,
            "stdDev" | "standardDeviation" => TableFunction::StdDev,
            "sum" => TableFunction::Sum,
            "var" | "variance" => TableFunction::Var,
            _ => return Err(Error::new(Status::InvalidArg, format!("Invalid table function: {}. Use 'average', 'count', 'countNumbers', 'max', 'min', 'stdDev', 'sum', or 'var'", function))),
        };
        self.column = self.column.clone().set_total_function(func);
        Ok(self)
    }

    /// Set a custom total label for the column
    #[napi]
    pub fn set_total_label(&mut self, label: String) -> &Self {
        self.column = self.column.clone().set_total_label(&label);
        self
    }

    /// Set a custom formula for the column
    #[napi]
    pub fn set_formula(&mut self, formula: String) -> &Self {
        use rust_xlsxwriter::Formula;
        self.column = self.column.clone().set_formula(Formula::new(&formula));
        self
    }
}
