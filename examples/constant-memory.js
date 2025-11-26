/**
 * Example demonstrating constant memory mode for large datasets
 *
 * Constant memory mode is designed for generating very large Excel files
 * with minimal memory usage. Memory usage remains flat (~0.02 MB) regardless
 * of dataset size.
 *
 * IMPORTANT LIMITATIONS:
 * - Data MUST be written row-by-row in sequential order
 * - Cannot write to previous rows once you've moved to a new row
 * - merge_range() only works for the current row
 * - add_table() does not work
 * - File sizes are larger (strings stored inline, not shared)
 */

const { Workbook, Format } = require('..');

console.log('Creating large Excel file with constant memory mode...');
console.log('This will write 100,000 rows with minimal memory usage.\n');

const startTime = Date.now();

// Create a new workbook
const workbook = new Workbook();

// Optional: Set custom temp directory (must be called before adding worksheets)
// workbook.setTempdir('./temp');

// Add worksheet with constant memory mode
const worksheet = workbook.addWorksheetWithConstantMemory('Large Dataset');

// Create formats
const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('blue')
  .setFontColor('white')
  .setAlign('center');

const currencyFormat = new Format()
  .setNumFormat('$#,##0.00');

const dateFormat = new Format()
  .setNumFormat('yyyy-mm-dd');

// Write headers (row 0)
worksheet.write(0, 0, 'ID', headerFormat);
worksheet.write(0, 1, 'Product', headerFormat);
worksheet.write(0, 2, 'Quantity', headerFormat);
worksheet.write(0, 3, 'Price', headerFormat);
worksheet.write(0, 4, 'Total', headerFormat);
worksheet.write(0, 5, 'Date', headerFormat);

// Set column widths
worksheet.setColumnWidth(0, 10);
worksheet.setColumnWidth(1, 20);
worksheet.setColumnWidth(2, 12);
worksheet.setColumnWidth(3, 12);
worksheet.setColumnWidth(4, 12);
worksheet.setColumnWidth(5, 15);

// Generate large dataset (100,000 rows)
// IMPORTANT: Must write rows sequentially!
const products = ['Widget A', 'Widget B', 'Widget C', 'Gadget X', 'Gadget Y'];
const numRows = 100000;

for (let row = 1; row <= numRows; row++) {
  const productIdx = row % products.length;
  const quantity = Math.floor(Math.random() * 100) + 1;
  const price = Math.random() * 100 + 10;
  const total = quantity * price;

  // Write data for current row
  worksheet.write(row, 0, row);
  worksheet.write(row, 1, products[productIdx]);
  worksheet.write(row, 2, quantity);
  worksheet.write(row, 3, price, currencyFormat);
  worksheet.write(row, 4, total, currencyFormat);
  worksheet.write(row, 5, `2024-01-${String((row % 28) + 1).padStart(2, '0')}`);

  // Progress indicator
  if (row % 10000 === 0) {
    const elapsedSeconds = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`Written ${row.toLocaleString()} rows in ${elapsedSeconds}s`);
  }
}

console.log('\nSaving workbook...');

// Save the workbook
workbook.save('constant-memory.xlsx');

const totalSeconds = ((Date.now() - startTime) / 1000).toFixed(2);
console.log(`\nCompleted in ${totalSeconds}s`);
console.log(`Created constant-memory.xlsx with ${numRows.toLocaleString()} rows`);
console.log('\nMemory usage remained constant throughout the process!');
console.log('Note: File size will be larger due to inline string storage.');
