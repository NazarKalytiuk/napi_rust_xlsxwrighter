/**
 * Basic example demonstrating rusc-xlsx usage
 * Run this after building: npm run build && node example.js
 */

const { Workbook, Format } = require('./wrapper.js');

console.log('Creating Excel workbook with rusc-xlsx...\n');

// Create a new workbook
const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Demo Sheet');

// Create formats
const headerFormat = new Format()
  .setBold()
  .setFontColor('white')
  .setBackgroundColor('#4472C4')
  .setAlign('center')
  .setVerticalAlign('center');

const currencyFormat = new Format()
  .setNumFormat('$#,##0.00');

const percentFormat = new Format()
  .setNumFormat('0.00%');

const dateFormat = new Format()
  .setNumFormat('yyyy-mm-dd');

// Write headers
console.log('Writing headers...');
worksheet.write(0, 0, 'Product', headerFormat);
worksheet.write(0, 1, 'Price', headerFormat);
worksheet.write(0, 2, 'Quantity', headerFormat);
worksheet.write(0, 3, 'Total', headerFormat);
worksheet.write(0, 4, 'Discount', headerFormat);
worksheet.write(0, 5, 'In Stock', headerFormat);

// Write data
console.log('Writing data...');
const products = [
  { name: 'Widget', price: 19.99, quantity: 100, discount: 0.1, inStock: true },
  { name: 'Gadget', price: 29.99, quantity: 50, discount: 0.15, inStock: true },
  { name: 'Doohickey', price: 9.99, quantity: 200, discount: 0.05, inStock: false },
  { name: 'Thingamajig', price: 49.99, quantity: 25, discount: 0.2, inStock: true },
];

products.forEach((product, index) => {
  const row = index + 1;

  // Product name
  worksheet.writeString(row, 0, product.name);

  // Price with currency format
  worksheet.write(row, 1, product.price, currencyFormat);

  // Quantity
  worksheet.writeNumber(row, 2, product.quantity);

  // Total (price * quantity) with currency format
  worksheet.write(row, 3, product.price * product.quantity, currencyFormat);

  // Discount with percent format
  worksheet.write(row, 4, product.discount, percentFormat);

  // In stock (boolean)
  worksheet.writeBoolean(row, 5, product.inStock);
});

// Set column widths
console.log('Formatting columns...');
worksheet.setColumnWidth(0, 20);
worksheet.setColumnWidth(1, 12);
worksheet.setColumnWidth(2, 12);
worksheet.setColumnWidth(3, 15);
worksheet.setColumnWidth(4, 12);
worksheet.setColumnWidth(5, 12);

// Add a summary section
const summaryRow = products.length + 2;
const summaryFormat = new Format()
  .setBold()
  .setBackgroundColor('#E7E6E6')
  .setAlign('right');

worksheet.write(summaryRow, 0, 'TOTAL ITEMS:', summaryFormat);
const totalQuantity = products.reduce((sum, p) => sum + p.quantity, 0);
worksheet.writeNumber(summaryRow, 2, totalQuantity);

// Merge cells for a title
const titleFormat = new Format()
  .setBold()
  .setFontSize(16)
  .setAlign('center')
  .setBackgroundColor('#203864')
  .setFontColor('white');

worksheet.mergeRange(0, 7, 0, 9, 'Sales Report 2025', titleFormat);

// Save to file
const filename = 'example-output.xlsx';
console.log(`\nSaving to ${filename}...`);
workbook.save(filename);

// Also demonstrate buffer export
const buffer = workbook.saveToBuffer();
console.log(`✅ Success! File size: ${buffer.length.toLocaleString()} bytes`);
console.log(`\nOpen ${filename} to view the result!`);
