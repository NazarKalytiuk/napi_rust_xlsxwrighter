/**
 * Advanced Features Example
 *
 * Demonstrates:
 * - Formulas (basic, dynamic array, array formulas)
 * - Data Validation (dropdowns, number ranges, custom formulas)
 * - Hyperlinks (URLs, internal links)
 * - Autofilter
 * - Freeze Panes
 */

const { Workbook, Format, Formula, DataValidation } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();

// ============================================================================
// FORMULAS
// ============================================================================
const formulaSheet = workbook.addWorksheet('Formulas');

// Header
const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('#4472C4')
  .setFontColor('white')
  .setAlign('center');

formulaSheet.write(0, 0, 'Item', headerFormat);
formulaSheet.write(0, 1, 'Quantity', headerFormat);
formulaSheet.write(0, 2, 'Price', headerFormat);
formulaSheet.write(0, 3, 'Total', headerFormat);

// Data
formulaSheet.write(1, 0, 'Apples');
formulaSheet.write(1, 1, 10);
formulaSheet.write(1, 2, 1.5);

formulaSheet.write(2, 0, 'Bananas');
formulaSheet.write(2, 1, 5);
formulaSheet.write(2, 2, 0.75);

formulaSheet.write(3, 0, 'Oranges');
formulaSheet.write(3, 1, 8);
formulaSheet.write(3, 2, 2.0);

// Formula: Calculate total for each row
const currencyFormat = new Format().setNumFormat('$#,##0.00');
formulaSheet.writeFormula(1, 3, new Formula('=B2*C2'), currencyFormat);
formulaSheet.writeFormula(2, 3, new Formula('=B3*C3'), currencyFormat);
formulaSheet.writeFormula(3, 3, new Formula('=B4*C4'), currencyFormat);

// Formula: Grand total
formulaSheet.write(5, 2, 'Grand Total:', new Format().setBold());
formulaSheet.writeFormula(5, 3, new Formula('=SUM(D2:D4)'), currencyFormat.setBold());

// Formula with result (for apps that don't auto-calculate)
formulaSheet.write(7, 0, 'Formula with preset result:');
const formulaWithResult = new Formula('=10*5').setResult('50');
formulaSheet.writeFormula(7, 1, formulaWithResult);

formulaSheet.setColumnWidth(0, 15);
formulaSheet.setColumnWidth(1, 10);
formulaSheet.setColumnWidth(2, 10);
formulaSheet.setColumnWidth(3, 12);

// ============================================================================
// DATA VALIDATION
// ============================================================================
const validationSheet = workbook.addWorksheet('Data Validation');

validationSheet.write(0, 0, 'Validation Examples', headerFormat);

// Dropdown list
validationSheet.write(2, 0, 'Select Status:');
const statusValidation = new DataValidation()
  .allowListStrings(['Pending', 'In Progress', 'Completed', 'Cancelled'])
  .setInputTitle('Status')
  .setInputMessage('Select a status from the dropdown')
  .setErrorStyle('stop')
  .setErrorTitle('Invalid Status')
  .setErrorMessage('Please select a valid status from the list');
validationSheet.addDataValidation(2, 1, 2, 1, statusValidation);

// Number range validation
validationSheet.write(4, 0, 'Enter Age (18-100):');
const ageValidation = new DataValidation()
  .allowWholeNumberBetween(18, 100)
  .setInputTitle('Age')
  .setInputMessage('Enter your age (must be between 18 and 100)')
  .setErrorStyle('warning')
  .setErrorTitle('Invalid Age')
  .setErrorMessage('Age must be between 18 and 100');
validationSheet.addDataValidation(4, 1, 4, 1, ageValidation);

// Decimal number validation
validationSheet.write(6, 0, 'Enter Price (> $0.00):');
const priceValidation = new DataValidation()
  .allowDecimalNumber('greaterThan', 0)
  .setInputTitle('Price')
  .setInputMessage('Enter a positive price')
  .setErrorStyle('stop')
  .setErrorTitle('Invalid Price')
  .setErrorMessage('Price must be greater than $0.00');
validationSheet.addDataValidation(6, 1, 6, 1, priceValidation);

// List from cell range
validationSheet.write(8, 0, 'Countries:', new Format().setBold());
validationSheet.write(9, 0, 'USA');
validationSheet.write(10, 0, 'Canada');
validationSheet.write(11, 0, 'Mexico');
validationSheet.write(12, 0, 'UK');

validationSheet.write(8, 2, 'Select Country:');
const countryValidation = new DataValidation()
  .allowListFormula('=$A$10:$A$13')
  .setInputTitle('Country')
  .setInputMessage('Select from the country list');
validationSheet.addDataValidation(8, 3, 8, 3, countryValidation);

// Custom formula validation (only allow even numbers)
validationSheet.write(14, 0, 'Enter Even Number:');
const evenValidation = new DataValidation()
  .allowCustom('=MOD(B15,2)=0')
  .setInputTitle('Even Number')
  .setInputMessage('Enter an even number')
  .setErrorStyle('information')
  .setErrorTitle('Not Even')
  .setErrorMessage('This number is not even, but you can still proceed');
validationSheet.addDataValidation(14, 1, 14, 1, evenValidation);

validationSheet.setColumnWidth(0, 25);
validationSheet.setColumnWidth(1, 15);
validationSheet.setColumnWidth(2, 18);
validationSheet.setColumnWidth(3, 15);

// ============================================================================
// HYPERLINKS
// ============================================================================
const linksSheet = workbook.addWorksheet('Hyperlinks');

linksSheet.write(0, 0, 'Hyperlink Examples', headerFormat);

// External URL
linksSheet.write(2, 0, 'External Website:');
linksSheet.writeUrl(2, 1, 'https://www.rust-lang.org');

// External URL with custom text
linksSheet.write(4, 0, 'Custom Link Text:');
linksSheet.writeUrlWithText(4, 1, 'https://www.rust-lang.org', 'Learn Rust');

// Another external link
linksSheet.write(6, 0, 'GitHub:');
linksSheet.writeUrlWithText(6, 1, 'https://github.com', 'Visit GitHub');

// Internal link to another sheet
validationSheet.write(8, 0, 'Internal Link:');
linksSheet.writeUrlWithText(8, 1, 'internal:Formulas!A1', 'Go to Formulas sheet');

// Email link
linksSheet.write(10, 0, 'Email:');
linksSheet.writeUrlWithText(10, 1, 'mailto:example@example.com', 'Send Email');

linksSheet.setColumnWidth(0, 20);
linksSheet.setColumnWidth(1, 30);

// ============================================================================
// AUTOFILTER & FREEZE PANES
// ============================================================================
const dataSheet = workbook.addWorksheet('Data Table');

// Freeze the top row (header)
dataSheet.setFreezePanes(1, 0);

// Write headers
const headers = ['Name', 'Department', 'Age', 'Salary', 'Start Date'];
headers.forEach((header, col) => {
  dataSheet.write(0, col, header, headerFormat);
});

// Sample data
const employees = [
  ['John Doe', 'Engineering', 32, 95000, '2020-01-15'],
  ['Jane Smith', 'Marketing', 28, 75000, '2021-03-22'],
  ['Bob Johnson', 'Engineering', 45, 110000, '2018-07-01'],
  ['Alice Williams', 'Sales', 35, 82000, '2019-11-10'],
  ['Charlie Brown', 'Engineering', 29, 88000, '2022-02-14'],
  ['Diana Prince', 'Marketing', 31, 79000, '2020-06-01'],
  ['Eve Anderson', 'Sales', 27, 71000, '2022-01-05'],
  ['Frank Miller', 'Engineering', 38, 102000, '2017-09-12'],
];

const salaryFormat = new Format().setNumFormat('$#,##0');

employees.forEach((employee, rowIndex) => {
  const row = rowIndex + 1;
  dataSheet.write(row, 0, employee[0]);
  dataSheet.write(row, 1, employee[1]);
  dataSheet.write(row, 2, employee[2]);
  dataSheet.write(row, 3, employee[3], salaryFormat);
  dataSheet.write(row, 4, employee[4]);
});

// Add autofilter to the data range (including header)
dataSheet.autofilter(0, 0, employees.length, headers.length - 1);

// Set column widths
dataSheet.setColumnWidth(0, 18);
dataSheet.setColumnWidth(1, 15);
dataSheet.setColumnWidth(2, 8);
dataSheet.setColumnWidth(3, 12);
dataSheet.setColumnWidth(4, 12);

// ============================================================================
// FREEZE PANES DEMO
// ============================================================================
const freezeSheet = workbook.addWorksheet('Freeze Panes Demo');

freezeSheet.write(0, 0, 'Freeze Panes Examples', headerFormat);
freezeSheet.write(2, 0, 'This sheet demonstrates freezing both rows and columns');

// Freeze both top row and first column
freezeSheet.setFreezePanes(3, 1);

// Create a grid to demonstrate
freezeSheet.write(2, 0, 'Fixed');
for (let col = 1; col < 10; col++) {
  freezeSheet.write(2, col, `Col ${col}`, headerFormat);
}

for (let row = 3; row < 20; row++) {
  freezeSheet.write(row, 0, `Row ${row}`, headerFormat);
  for (let col = 1; col < 10; col++) {
    freezeSheet.write(row, col, `R${row}C${col}`);
  }
}

// ============================================================================
// ARRAY FORMULAS
// ============================================================================
const arraySheet = workbook.addWorksheet('Array Formulas');

arraySheet.write(0, 0, 'Array Formula Examples', headerFormat);

// Data for array formula
arraySheet.write(2, 0, 'Values:');
const values = [10, 20, 30, 40, 50];
values.forEach((value, i) => {
  arraySheet.write(3 + i, 0, value);
});

arraySheet.write(2, 1, 'Multiplier:');
const multipliers = [2, 3, 4, 5, 6];
multipliers.forEach((value, i) => {
  arraySheet.write(3 + i, 1, value);
});

// Array formula to multiply two ranges
arraySheet.write(2, 2, 'Result (Array Formula):');
arraySheet.writeArrayFormula(3, 2, 7, 2, new Formula('=A4:A8*B4:B8'));

// Note about the array formula
arraySheet.write(9, 0, 'Note: The array formula multiplies corresponding elements');
arraySheet.setColumnWidth(0, 12);
arraySheet.setColumnWidth(1, 12);
arraySheet.setColumnWidth(2, 25);

// Save the workbook
workbook.save('advanced-features.xlsx');

console.log('✓ Created advanced-features.xlsx');
console.log('');
console.log('Features demonstrated:');
console.log('  ✓ Basic formulas (SUM, multiplication)');
console.log('  ✓ Array formulas');
console.log('  ✓ Data validation (dropdowns, number ranges, custom)');
console.log('  ✓ Hyperlinks (external URLs, internal links, email)');
console.log('  ✓ Autofilter');
console.log('  ✓ Freeze panes');
