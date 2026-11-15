'use strict';

// A business-report export: a header row, then one order per row in 30
// columns of integers, decimals, unique and repeated strings, datetimes, dates,
// booleans and empty cells (null).
//
// benchmarks/pure-rust/src/export.rs is the Rust copy; both compute every
// value from the row index with the same integer arithmetic, so the numbers
// are the same double in both languages. Dates are UTC.

const COLUMNS = 30;

const HEADERS = [
  'Order ID',
  'Order number',
  'Created at',
  'Customer ID',
  'Customer',
  'Email',
  'Country',
  'City',
  'Status',
  'Category',
  'SKU',
  'Product',
  'Quantity',
  'Unit price',
  'Discount',
  'Amount',
  'Tax',
  'Currency',
  'Paid',
  'Paid at',
  'Ship date',
  'Warehouse',
  'Weight (kg)',
  'Gift',
  'Rating',
  'Sales rep',
  'Channel',
  'Coupon',
  'Notes',
  'Updated at',
];

// prettier-ignore
const COUNTRIES = [
  'UA', 'PL', 'DE', 'FR', 'ES', 'IT', 'NL', 'BE', 'AT', 'CZ', 'SK', 'HU', 'RO', 'BG', 'GR', 'PT',
  'IE', 'GB', 'SE', 'NO', 'DK', 'FI', 'EE', 'LV', 'LT', 'US', 'CA', 'MX', 'JP', 'AU',
];
const STATUSES = ['new', 'paid', 'packed', 'shipped', 'delivered', 'cancelled'];
const CURRENCIES = ['EUR', 'USD', 'GBP', 'UAH'];
const CHANNELS = ['web', 'mobile', 'store', 'phone', 'partner'];

// 2024-01-01T00:00:00Z and 2024-01-02T00:00:00Z, in seconds.
const CREATED_BASE = 1_704_067_200;
const SHIP_DATE_BASE = 1_704_153_600;
const DAY = 86_400;

const dateAt = (seconds) => new Date(seconds * 1_000);

/** The value of the export cell at (row, col). */
function exportValue(row, col) {
  if (row === 0) {
    return HEADERS[col];
  }

  const r = row - 1;
  const customer = (r * 7) % 50_000;
  const product = (r * 31) % 20_000;
  const quantity = 1 + ((r * 17) % 50);
  const priceCents = ((r * 7_919) % 100_000) + 99;
  const discountPercent = (r * 5) % 30;
  const netCentsX100 = quantity * priceCents * (100 - discountPercent);
  const created = CREATED_BASE + r * 37;

  switch (col) {
    case 0:
      return r + 1;
    case 1:
      return `ORD-${String(r + 1).padStart(8, '0')}`;
    case 2:
      return dateAt(created);
    case 3:
      return 100_000 + customer;
    case 4:
      return `Customer ${customer}`;
    case 5:
      return `customer${customer}@example.com`;
    case 6:
      return COUNTRIES[r % 30];
    case 7:
      return `City ${(r * 13) % 500}`;
    case 8:
      return STATUSES[(r * 3) % 6];
    case 9:
      return `Category ${(r * 11) % 25}`;
    case 10:
      return `SKU-${String(product).padStart(5, '0')}`;
    case 11:
      return `Product ${product}`;
    case 12:
      return quantity;
    case 13:
      return priceCents / 100;
    case 14:
      return discountPercent / 100;
    case 15:
      return netCentsX100 / 10_000;
    case 16:
      return (netCentsX100 * 20) / 1_000_000;
    case 17:
      return CURRENCIES[r % 4];
    case 18:
      return r % 7 !== 0;
    case 19:
      return r % 7 !== 0 ? dateAt(created + 3_600 * (1 + (r % 48))) : null;
    case 20:
      return r % 5 !== 4 ? dateAt(SHIP_DATE_BASE + Math.floor(r / 200) * DAY) : null;
    case 21:
      return `WH-${r % 12}`;
    case 22:
      return ((r * 37) % 50_000) / 1_000;
    case 23:
      return r % 11 === 0;
    case 24:
      return r % 3 !== 0 ? 1 + (r % 5) : null;
    case 25:
      return `Rep ${(r * 3) % 120}`;
    case 26:
      return CHANNELS[r % 5];
    case 27:
      return r % 4 === 0 ? `SAVE${5 + (r % 9) * 5}` : null;
    case 28:
      return r % 10 === 0
        ? `Delivery note for order ${r + 1}: leave the parcel at the reception desk if nobody answers the door.`
        : null;
    case 29:
      return dateAt(created + DAY * (r % 30) + 600);
    default:
      return null;
  }
}

module.exports = { COLUMNS, exportValue };
