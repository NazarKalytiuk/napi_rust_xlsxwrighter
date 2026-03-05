/**
 * Basic example showing how to create a workbook and write data
 */

const { Workbook } = require('@nazarkalytiuk/rusc-xlsx');

// Create a new workbook
const workbook = new Workbook();

// Add a worksheet
const worksheet = workbook.addWorksheet('Sheet1');

// Write some data
worksheet.write(0, 0, 'Hello');
worksheet.write(0, 1, 'World');
worksheet.write(1, 0, 'Numbers:');
worksheet.write(1, 1, 42);
worksheet.write(2, 0, 'Boolean:');
worksheet.write(2, 1, true);

// Save the workbook
workbook.save('basic.xlsx');

console.log('Created basic.xlsx successfully!');
