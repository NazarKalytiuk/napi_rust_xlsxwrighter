/**
 * Example demonstrating low memory mode for large datasets with repeated strings
 *
 * Low memory mode balances memory efficiency with file size optimization.
 * It maintains a shared string table, making it ideal for datasets with
 * many repeated string values.
 *
 * IMPORTANT LIMITATIONS:
 * - Data MUST be written row-by-row in sequential order
 * - Cannot write to previous rows once you've moved to a new row
 * - merge_range() only works for the current row
 * - add_table() does not work
 * - Better file size than constant memory (uses shared strings)
 *
 * COMPARISON:
 * - Memory usage: Low (scales with unique strings)
 * - File size: Smaller than constant memory
 * - Best for: Data with repeated strings (categories, status values, etc.)
 */

const { Workbook, Format } = require('..');

console.log('Creating large Excel file with low memory mode...');
console.log('This dataset has many repeated strings, demonstrating the benefit of shared string table.\n');

const startTime = Date.now();

// Create a new workbook
const workbook = new Workbook();

// Add worksheet with low memory mode
// This maintains a shared string table for better file size
const worksheet = workbook.addWorksheetWithLowMemory('Sales Data');

// Create formats
const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('green')
  .setFontColor('white')
  .setAlign('center');

const currencyFormat = new Format()
  .setNumFormat('$#,##0.00');

const percentFormat = new Format()
  .setNumFormat('0.0%');

// Write headers
worksheet.write(0, 0, 'Order ID', headerFormat);
worksheet.write(0, 1, 'Region', headerFormat);
worksheet.write(0, 2, 'Product Category', headerFormat);
worksheet.write(0, 3, 'Product', headerFormat);
worksheet.write(0, 4, 'Sales Rep', headerFormat);
worksheet.write(0, 5, 'Status', headerFormat);
worksheet.write(0, 6, 'Amount', headerFormat);
worksheet.write(0, 7, 'Commission', headerFormat);

// Set column widths
worksheet.setColumnWidth(0, 12);
worksheet.setColumnWidth(1, 15);
worksheet.setColumnWidth(2, 18);
worksheet.setColumnWidth(3, 20);
worksheet.setColumnWidth(4, 18);
worksheet.setColumnWidth(5, 12);
worksheet.setColumnWidth(6, 12);
worksheet.setColumnWidth(7, 12);

// Define repeated string values (these will be stored efficiently in shared string table)
const regions = ['North', 'South', 'East', 'West', 'Central'];
const categories = ['Electronics', 'Furniture', 'Office Supplies', 'Technology', 'Accessories'];
const products = [
  'Laptop Pro 15"',
  'Desk Chair Executive',
  'Printer Laser Color',
  'Monitor 27" 4K',
  'Keyboard Wireless',
  'Mouse Ergonomic',
  'Desk Standing',
  'Headphones Noise-Canceling'
];
const salesReps = [
  'Alice Johnson',
  'Bob Smith',
  'Carol Williams',
  'David Brown',
  'Emma Davis',
  'Frank Miller',
  'Grace Wilson',
  'Henry Moore'
];
const statuses = ['Pending', 'Approved', 'Shipped', 'Delivered', 'Cancelled'];

// Generate dataset with lots of repeated strings
const numRows = 50000;

for (let row = 1; row <= numRows; row++) {
  const amount = Math.random() * 5000 + 100;
  const commission = amount * 0.08;

  // Write data - notice how region, category, product, salesRep, and status repeat often
  worksheet.write(row, 0, 10000 + row);
  worksheet.write(row, 1, regions[row % regions.length]);
  worksheet.write(row, 2, categories[row % categories.length]);
  worksheet.write(row, 3, products[row % products.length]);
  worksheet.write(row, 4, salesReps[row % salesReps.length]);
  worksheet.write(row, 5, statuses[row % statuses.length]);
  worksheet.write(row, 6, amount, currencyFormat);
  worksheet.write(row, 7, commission, currencyFormat);

  // Progress indicator
  if (row % 10000 === 0) {
    const elapsedSeconds = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`Written ${row.toLocaleString()} rows in ${elapsedSeconds}s`);
  }
}

console.log('\nSaving workbook...');

// Save the workbook
workbook.save('low-memory.xlsx');

const totalSeconds = ((Date.now() - startTime) / 1000).toFixed(2);
console.log(`\nCompleted in ${totalSeconds}s`);
console.log(`Created low-memory.xlsx with ${numRows.toLocaleString()} rows`);
console.log('\nBenefits of low memory mode:');
console.log('- Memory usage scaled with unique strings (not total rows)');
console.log('- File size is optimized due to shared string table');
console.log('- Ideal for datasets with categorical/repeated string data');

console.log('\nString statistics:');
console.log(`- Total rows: ${numRows.toLocaleString()}`);
console.log(`- Unique strings: ${regions.length + categories.length + products.length + salesReps.length + statuses.length}`);
console.log(`- String reuse ratio: ~${Math.floor(numRows / 31)}x`);
