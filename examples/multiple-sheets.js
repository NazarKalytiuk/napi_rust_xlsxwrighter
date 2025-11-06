/**
 * Example showing how to work with multiple worksheets
 */

const { Workbook, Format } = require('../index');

// Create a new workbook
const workbook = new Workbook();

// Create multiple worksheets
const salesSheet = workbook.addWorksheet('Sales');
const expensesSheet = workbook.addWorksheet('Expenses');
const summarySheet = workbook.addWorksheet('Summary');

// Header format
const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('gray');

// Sales sheet
salesSheet.write(0, 0, 'Month', headerFormat);
salesSheet.write(0, 1, 'Revenue', headerFormat);
salesSheet.write(1, 0, 'January');
salesSheet.write(1, 1, 50000);
salesSheet.write(2, 0, 'February');
salesSheet.write(2, 1, 62000);
salesSheet.write(3, 0, 'March');
salesSheet.write(3, 1, 58000);

// Expenses sheet
expensesSheet.write(0, 0, 'Month', headerFormat);
expensesSheet.write(0, 1, 'Costs', headerFormat);
expensesSheet.write(1, 0, 'January');
expensesSheet.write(1, 1, 30000);
expensesSheet.write(2, 0, 'February');
expensesSheet.write(2, 1, 35000);
expensesSheet.write(3, 0, 'March');
expensesSheet.write(3, 1, 32000);

// Summary sheet with merged cells
const titleFormat = new Format()
  .setBold()
  .setFontSize(16)
  .setAlign('center');

summarySheet.mergeRange(0, 0, 0, 2, 'Q1 Financial Summary');
summarySheet.write(0, 0, 'Q1 Financial Summary', titleFormat);

summarySheet.write(2, 0, 'Total Revenue:', headerFormat);
summarySheet.write(2, 1, 170000);

summarySheet.write(3, 0, 'Total Costs:', headerFormat);
summarySheet.write(3, 1, 97000);

summarySheet.write(4, 0, 'Net Profit:', headerFormat);
summarySheet.write(4, 1, 73000);

// Adjust column widths
[salesSheet, expensesSheet, summarySheet].forEach(sheet => {
  sheet.setColumnWidth(0, 15);
  sheet.setColumnWidth(1, 15);
});

// Save the workbook
workbook.save('multiple-sheets.xlsx');

console.log('Created multiple-sheets.xlsx successfully!');
