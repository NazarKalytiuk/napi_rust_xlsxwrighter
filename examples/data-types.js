/**
 * Example demonstrating different data types
 */

const { Workbook, Format } = require('../index');

// Create a new workbook
const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Data Types');

// Create formats
const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('yellow');

const numberFormat = new Format()
  .setNumFormat('#,##0.00');

const integerFormat = new Format()
  .setNumFormat('#,##0');

const percentFormat = new Format()
  .setNumFormat('0.00%');

const scientificFormat = new Format()
  .setNumFormat('0.00E+00');

// Headers
worksheet.write(0, 0, 'Data Type', headerFormat);
worksheet.write(0, 1, 'Value', headerFormat);
worksheet.write(0, 2, 'Notes', headerFormat);

// Strings
worksheet.write(1, 0, 'String');
worksheet.writeString(1, 1, 'Hello, World!');
worksheet.write(1, 2, 'Plain text');

// Numbers
worksheet.write(2, 0, 'Integer');
worksheet.write(2, 1, 42, integerFormat);
worksheet.write(2, 2, 'Whole number');

worksheet.write(3, 0, 'Float');
worksheet.write(3, 1, 3.14159, numberFormat);
worksheet.write(3, 2, 'Decimal number');

worksheet.write(4, 0, 'Large Number');
worksheet.write(4, 1, 1234567.89, numberFormat);
worksheet.write(4, 2, 'With thousands separator');

worksheet.write(5, 0, 'Percentage');
worksheet.write(5, 1, 0.85, percentFormat);
worksheet.write(5, 2, '85%');

worksheet.write(6, 0, 'Scientific');
worksheet.write(6, 1, 1234567890, scientificFormat);
worksheet.write(6, 2, 'Scientific notation');

// Booleans
worksheet.write(7, 0, 'Boolean True');
worksheet.writeBoolean(7, 1, true);
worksheet.write(7, 2, 'TRUE value');

worksheet.write(8, 0, 'Boolean False');
worksheet.writeBoolean(8, 1, false);
worksheet.write(8, 2, 'FALSE value');

// Negative numbers
worksheet.write(9, 0, 'Negative');
worksheet.write(9, 1, -99.99, numberFormat);
worksheet.write(9, 2, 'Negative number');

// Zero
worksheet.write(10, 0, 'Zero');
worksheet.write(10, 1, 0, numberFormat);
worksheet.write(10, 2, 'Zero value');

// Adjust column widths
worksheet.setColumnWidth(0, 15);
worksheet.setColumnWidth(1, 20);
worksheet.setColumnWidth(2, 25);

// Save the workbook
workbook.save('data-types.xlsx');

console.log('Created data-types.xlsx successfully!');
