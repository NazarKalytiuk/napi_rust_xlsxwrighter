/**
 * Example demonstrating cell formatting
 */

const { Workbook, Format } = require('../index');

// Create a new workbook
const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Formatted');

// Create various formats
const headerFormat = new Format()
  .setBold()
  .setFontSize(14)
  .setFontColor('white')
  .setBackgroundColor('blue')
  .setAlign('center');

const currencyFormat = new Format()
  .setNumFormat('$#,##0.00');

const percentFormat = new Format()
  .setNumFormat('0.00%');

const dateFormat = new Format()
  .setNumFormat('yyyy-mm-dd');

const borderedFormat = new Format()
  .setBorder('thin');

// Write headers with formatting
worksheet.write(0, 0, 'Product', headerFormat);
worksheet.write(0, 1, 'Price', headerFormat);
worksheet.write(0, 2, 'Discount', headerFormat);
worksheet.write(0, 3, 'Final Price', headerFormat);

// Write data with formatting
worksheet.write(1, 0, 'Widget A', borderedFormat);
worksheet.write(1, 1, 99.99, currencyFormat);
worksheet.write(1, 2, 0.15, percentFormat);
worksheet.write(1, 3, 84.99, currencyFormat);

worksheet.write(2, 0, 'Widget B', borderedFormat);
worksheet.write(2, 1, 149.99, currencyFormat);
worksheet.write(2, 2, 0.20, percentFormat);
worksheet.write(2, 3, 119.99, currencyFormat);

// Adjust column widths
worksheet.setColumnWidth(0, 15);
worksheet.setColumnWidth(1, 12);
worksheet.setColumnWidth(2, 12);
worksheet.setColumnWidth(3, 12);

// Save the workbook
workbook.save('formatting.xlsx');

console.log('Created formatting.xlsx successfully!');
