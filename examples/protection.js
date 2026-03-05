/**
 * Protection example - demonstrates worksheet protection with granular permissions
 */

const { Workbook, Format, ProtectionOptions } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();

// --- Basic Protection ---
const ws1 = workbook.addWorksheet('Basic Protection');

ws1.setColumnWidth(0, 25);
ws1.setColumnWidth(1, 20);

const headerFmt = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');
const lockedFmt = new Format().setBackgroundColor('#F2F2F2').setBorder('thin');
const unlockedFmt = new Format().setUnlocked().setBackgroundColor('#E2EFDA').setBorder('thin');

ws1.write(0, 0, 'Field', headerFmt);
ws1.write(0, 1, 'Value', headerFmt);

ws1.write(1, 0, 'Name (locked)', lockedFmt);
ws1.write(1, 1, 'John Doe', lockedFmt);

ws1.write(2, 0, 'Email (editable)', lockedFmt);
ws1.writeBlank(2, 1, unlockedFmt); // Users can edit this cell

ws1.write(3, 0, 'Phone (editable)', lockedFmt);
ws1.writeBlank(3, 1, unlockedFmt); // Users can edit this cell

ws1.write(4, 0, 'ID (locked)', lockedFmt);
ws1.write(4, 1, 'EMP-001', lockedFmt);

// Protect - all cells locked by default, unlocked cells are editable
ws1.protect();

// --- Password Protection ---
const ws2 = workbook.addWorksheet('Password Protected');

ws2.setColumnWidth(0, 25);
ws2.setColumnWidth(1, 20);

ws2.write(0, 0, 'This sheet is password-protected');
ws2.write(1, 0, 'Password: "secret"');
ws2.write(3, 0, 'Budget Item', headerFmt);
ws2.write(3, 1, 'Amount', headerFmt);

ws2.write(4, 0, 'Marketing');
ws2.write(4, 1, 50000);
ws2.write(5, 0, 'Engineering');
ws2.write(5, 1, 120000);
ws2.write(6, 0, 'Operations');
ws2.write(6, 1, 75000);

ws2.protectWithPassword('secret');

// --- Granular Permissions ---
const ws3 = workbook.addWorksheet('Custom Permissions');

ws3.setColumnWidth(0, 30);

ws3.write(0, 0, 'Sheet with custom protection options:');
ws3.write(2, 0, 'Allowed: sort, autofilter, format cells');
ws3.write(3, 0, 'Blocked: insert/delete rows/cols, edit objects');

// Write some data to sort/filter
ws3.write(5, 0, 'Name', headerFmt);
ws3.write(5, 1, 'Score', headerFmt);
ws3.write(6, 0, 'Alice');
ws3.write(6, 1, 92);
ws3.write(7, 0, 'Bob');
ws3.write(7, 1, 85);
ws3.write(8, 0, 'Carol');
ws3.write(8, 1, 97);
ws3.write(9, 0, 'Dave');
ws3.write(9, 1, 78);
ws3.autofilter(5, 0, 9, 1);

const options = new ProtectionOptions()
  .setSort(true)
  .setUseAutofilter(true)
  .setFormatCells(true)
  .setFormatColumns(true)
  .setFormatRows(true)
  .setInsertColumns(false)
  .setInsertRows(false)
  .setDeleteColumns(false)
  .setDeleteRows(false);

ws3.protectWithPasswordAndOptions('demo', options);

// --- Hidden Formulas ---
const ws4 = workbook.addWorksheet('Hidden Formulas');

ws4.setColumnWidth(0, 20);
ws4.setColumnWidth(1, 15);

const hiddenFmt = new Format().setHidden();

ws4.write(0, 0, 'This sheet hides formulas in the formula bar');
ws4.write(2, 0, 'Value A');
ws4.write(2, 1, 100);
ws4.write(3, 0, 'Value B');
ws4.write(3, 1, 200);
ws4.write(4, 0, 'Sum (hidden formula)');
ws4.write(4, 1, 300, hiddenFmt);

ws4.protect();

workbook.save('protection.xlsx');
console.log('Created protection.xlsx with 4 protection examples:');
console.log('  - Basic protection (locked/unlocked cells)');
console.log('  - Password protection');
console.log('  - Granular permissions (sort, filter, format)');
console.log('  - Hidden formulas');
