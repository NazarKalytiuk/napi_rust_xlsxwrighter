//! A business-report export: a header row, then one order per row in 30
//! columns of integers, decimals, unique and repeated strings, datetimes,
//! dates, booleans and empty cells.
//!
//! `benchmarks/export-workload.js` is the JavaScript copy; both compute every
//! value from the row index with the same integer arithmetic, so the numbers
//! are the same `f64` in both languages.

use rust_xlsxwriter::{ExcelDateTime, Format, Worksheet, XlsxError};

/// Columns of the export.
pub const COLUMNS: u16 = 30;

const HEADERS: [&str; COLUMNS as usize] = [
    "Order ID",
    "Order number",
    "Created at",
    "Customer ID",
    "Customer",
    "Email",
    "Country",
    "City",
    "Status",
    "Category",
    "SKU",
    "Product",
    "Quantity",
    "Unit price",
    "Discount",
    "Amount",
    "Tax",
    "Currency",
    "Paid",
    "Paid at",
    "Ship date",
    "Warehouse",
    "Weight (kg)",
    "Gift",
    "Rating",
    "Sales rep",
    "Channel",
    "Coupon",
    "Notes",
    "Updated at",
];

const COUNTRIES: [&str; 30] = [
    "UA", "PL", "DE", "FR", "ES", "IT", "NL", "BE", "AT", "CZ", "SK", "HU", "RO", "BG", "GR", "PT",
    "IE", "GB", "SE", "NO", "DK", "FI", "EE", "LV", "LT", "US", "CA", "MX", "JP", "AU",
];
const STATUSES: [&str; 6] = ["new", "paid", "packed", "shipped", "delivered", "cancelled"];
const CURRENCIES: [&str; 4] = ["EUR", "USD", "GBP", "UAH"];
const CHANNELS: [&str; 5] = ["web", "mobile", "store", "phone", "partner"];

/// 2024-01-01T00:00:00Z.
const CREATED_BASE: i64 = 1_704_067_200;
/// 2024-01-02T00:00:00Z.
const SHIP_DATE_BASE: i64 = 1_704_153_600;
const DAY: i64 = 86_400;

/// A cell of the export.
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    /// Seconds since the Unix epoch, UTC.
    Timestamp(i64),
    Empty,
}

/// The value of the export cell at `row`, `col`.
pub fn value(row: u32, col: u16) -> Value {
    if row == 0 {
        return Value::Text(HEADERS[usize::from(col)].to_owned());
    }

    let r = u64::from(row - 1);
    let customer = (r * 7) % 50_000;
    let product = (r * 31) % 20_000;
    let quantity = 1 + (r * 17) % 50;
    let price_cents = (r * 7_919) % 100_000 + 99;
    let discount_percent = (r * 5) % 30;
    let net_cents_x100 = quantity * price_cents * (100 - discount_percent);
    let created = CREATED_BASE + r as i64 * 37;

    match col {
        0 => Value::Number((r + 1) as f64),
        1 => Value::Text(format!("ORD-{:08}", r + 1)),
        2 => Value::Timestamp(created),
        3 => Value::Number((100_000 + customer) as f64),
        4 => Value::Text(format!("Customer {customer}")),
        5 => Value::Text(format!("customer{customer}@example.com")),
        6 => Value::Text(COUNTRIES[(r % 30) as usize].to_owned()),
        7 => Value::Text(format!("City {}", (r * 13) % 500)),
        8 => Value::Text(STATUSES[((r * 3) % 6) as usize].to_owned()),
        9 => Value::Text(format!("Category {}", (r * 11) % 25)),
        10 => Value::Text(format!("SKU-{product:05}")),
        11 => Value::Text(format!("Product {product}")),
        12 => Value::Number(quantity as f64),
        13 => Value::Number(price_cents as f64 / 100.0),
        14 => Value::Number(discount_percent as f64 / 100.0),
        15 => Value::Number(net_cents_x100 as f64 / 10_000.0),
        16 => Value::Number((net_cents_x100 * 20) as f64 / 1_000_000.0),
        17 => Value::Text(CURRENCIES[(r % 4) as usize].to_owned()),
        18 => Value::Bool(r % 7 != 0),
        19 if r % 7 != 0 => Value::Timestamp(created + 3_600 * (1 + (r % 48) as i64)),
        20 if r % 5 != 4 => Value::Timestamp(SHIP_DATE_BASE + (r / 200) as i64 * DAY),
        21 => Value::Text(format!("WH-{}", r % 12)),
        22 => Value::Number(((r * 37) % 50_000) as f64 / 1_000.0),
        23 => Value::Bool(r % 11 == 0),
        24 if r % 3 != 0 => Value::Number((1 + r % 5) as f64),
        25 => Value::Text(format!("Rep {}", (r * 3) % 120)),
        26 => Value::Text(CHANNELS[(r % 5) as usize].to_owned()),
        27 if r % 4 == 0 => Value::Text(format!("SAVE{}", 5 + (r % 9) * 5)),
        28 if r % 10 == 0 => Value::Text(format!(
            "Delivery note for order {}: leave the parcel at the reception desk if nobody answers the door.",
            r + 1
        )),
        29 => Value::Timestamp(created + DAY * (r % 30) as i64 + 600),
        _ => Value::Empty,
    }
}

/// Write `rows` rows of the export, one typed call per cell. Datetimes use
/// Excel's built-in `m/d/yyyy h:mm` format and whole days `m/d/yyyy`, as the
/// package writes a JavaScript `Date` without a format.
pub fn write_export(worksheet: &mut Worksheet, rows: u32, cols: u16) -> Result<(), XlsxError> {
    let datetime_format = Format::new().set_num_format_index(22);
    let date_format = Format::new().set_num_format_index(14);

    for row in 0..rows {
        for col in 0..cols {
            match value(row, col) {
                Value::Number(number) => {
                    worksheet.write_number(row, col, number)?;
                }
                Value::Text(text) => {
                    worksheet.write_string(row, col, text)?;
                }
                Value::Bool(boolean) => {
                    worksheet.write_boolean(row, col, boolean)?;
                }
                Value::Timestamp(seconds) => {
                    let datetime = ExcelDateTime::from_timestamp(seconds)?;
                    let format = if seconds % DAY == 0 {
                        &date_format
                    } else {
                        &datetime_format
                    };
                    worksheet.write_datetime_with_format(row, col, &datetime, format)?;
                }
                Value::Empty => {}
            }
        }
    }

    Ok(())
}
