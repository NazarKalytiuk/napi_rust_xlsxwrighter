const { Workbook, Format } = require('@nazarkalytiuk/rusc-xlsx');

console.log('Creating Excel file with advanced styling features...');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Advanced Styling');

// Set column widths
worksheet.setColumnWidth(0, 25);
worksheet.setColumnWidth(1, 30);

let row = 0;

// ========== FONT STYLING ==========
const titleFormat = new Format();
titleFormat.setBold();
titleFormat.setFontSize(16);
titleFormat.setFontName('Arial');
titleFormat.setFontColor('#1F4E78');
worksheet.write(row++, 0, 'FONT STYLING', titleFormat);
row++;

// Underline styles
const underlineFormat = new Format();
underlineFormat.setUnderline('single');
worksheet.write(row++, 0, 'Single Underline', underlineFormat);

const doubleUnderlineFormat = new Format();
doubleUnderlineFormat.setUnderline('double');
worksheet.write(row++, 0, 'Double Underline', doubleUnderlineFormat);

// Strikethrough
const strikeFormat = new Format();
strikeFormat.setStrikethrough();
worksheet.write(row++, 0, 'Strikethrough Text', strikeFormat);

// Different fonts
const timesFormat = new Format();
timesFormat.setFontName('Times New Roman');
timesFormat.setFontSize(14);
worksheet.write(row++, 0, 'Times New Roman Font', timesFormat);

row++;

// ========== TEXT FORMATTING ==========
const sectionFormat = new Format();
sectionFormat.setBold();
sectionFormat.setFontSize(14);
sectionFormat.setFontColor('#4472C4');
worksheet.write(row++, 0, 'TEXT FORMATTING', sectionFormat);
row++;

// Text wrap
const wrapFormat = new Format();
wrapFormat.setTextWrap();
worksheet.setRowHeight(row, 40);
worksheet.write(row++, 0, 'This is a long text that will wrap to multiple lines within the cell', wrapFormat);

// Indent
const indent1 = new Format();
indent1.setIndent(1);
worksheet.write(row++, 0, 'Indent level 1', indent1);

const indent2 = new Format();
indent2.setIndent(2);
worksheet.write(row++, 0, 'Indent level 2', indent2);

// Rotation
const rotateFormat = new Format();
rotateFormat.setRotation(45);
worksheet.setRowHeight(row, 50);
worksheet.write(row++, 0, 'Rotated 45 degrees', rotateFormat);

// Shrink to fit
const shrinkFormat = new Format();
shrinkFormat.setShrink();
worksheet.write(row++, 0, 'This text will shrink to fit the cell width', shrinkFormat);

row++;

// ========== BORDERS & COLORS ==========
worksheet.write(row++, 0, 'BORDERS & COLORS', sectionFormat);
row++;

// Individual borders
const topBorderFormat = new Format();
topBorderFormat.setBorderTop('thick');
topBorderFormat.setBorderTopColor('red');
worksheet.write(row++, 0, 'Thick red top border', topBorderFormat);

const bottomBorderFormat = new Format();
bottomBorderFormat.setBorderBottom('medium');
bottomBorderFormat.setBorderBottomColor('blue');
worksheet.write(row++, 0, 'Medium blue bottom border', bottomBorderFormat);

const leftBorderFormat = new Format();
leftBorderFormat.setBorderLeft('dashed');
leftBorderFormat.setBorderLeftColor('green');
worksheet.write(row++, 0, 'Dashed green left border', leftBorderFormat);

const rightBorderFormat = new Format();
rightBorderFormat.setBorderRight('dotted');
rightBorderFormat.setBorderRightColor('orange');
worksheet.write(row++, 0, 'Dotted orange right border', rightBorderFormat);

// All borders with different colors
const allBordersFormat = new Format();
allBordersFormat.setBorderTop('medium');
allBordersFormat.setBorderTopColor('red');
allBordersFormat.setBorderBottom('medium');
allBordersFormat.setBorderBottomColor('blue');
allBordersFormat.setBorderLeft('medium');
allBordersFormat.setBorderLeftColor('green');
allBordersFormat.setBorderRight('medium');
allBordersFormat.setBorderRightColor('purple');
worksheet.write(row++, 0, 'Each border different color', allBordersFormat);

row++;

// ========== PATTERNS & FILLS ==========
worksheet.write(row++, 0, 'PATTERNS & FILLS', sectionFormat);
row++;

const solidFillFormat = new Format();
solidFillFormat.setPattern('solid');
solidFillFormat.setBackgroundColor('#FFE699');
worksheet.write(row++, 0, 'Solid yellow fill', solidFillFormat);

const grayPatternFormat = new Format();
grayPatternFormat.setPattern('mediumGray');
grayPatternFormat.setBackgroundColor('#D9E1F2');
grayPatternFormat.setForegroundColor('#4472C4');
worksheet.write(row++, 0, 'Medium gray pattern', grayPatternFormat);

row++;

// ========== COMPREHENSIVE EXAMPLE ==========
worksheet.write(row++, 0, 'COMPREHENSIVE STYLING', sectionFormat);
row++;

const fancyFormat = new Format();
fancyFormat.setBold();
fancyFormat.setItalic();
fancyFormat.setFontName('Calibri');
fancyFormat.setFontSize(12);
fancyFormat.setFontColor('white');
fancyFormat.setBackgroundColor('#4472C4');
fancyFormat.setAlign('center');
fancyFormat.setVerticalAlign('center');
fancyFormat.setBorder('medium');
fancyFormat.setBorderColor('white');
worksheet.setRowHeight(row, 30);
worksheet.write(row++, 0, 'Everything Combined!', fancyFormat);

// Number formatting with styling
row++;
worksheet.write(row++, 0, 'NUMBER FORMATTING', sectionFormat);
row++;

const currencyFormat = new Format();
currencyFormat.setNumFormat('$#,##0.00');
currencyFormat.setFontColor('#0B6623');
currencyFormat.setBold();
worksheet.write(row++, 0, 1234.56, currencyFormat);

const percentFormat = new Format();
percentFormat.setNumFormat('0.00%');
percentFormat.setFontColor('#C70039');
percentFormat.setItalic();
worksheet.write(row++, 0, 0.856, percentFormat);

const dateFormat = new Format();
dateFormat.setNumFormat('yyyy-mm-dd');
dateFormat.setFontName('Courier New');
worksheet.write(row++, 0, 45000, dateFormat);

// Save the workbook
workbook.save('advanced-styling.xlsx');

console.log('✓ Created advanced-styling.xlsx');
console.log('');
console.log('New styling features demonstrated:');
console.log('  ✓ Font names (Arial, Times New Roman, Calibri, Courier)');
console.log('  ✓ Underline styles (single, double, accounting)');
console.log('  ✓ Strikethrough text');
console.log('  ✓ Text wrapping');
console.log('  ✓ Text indentation');
console.log('  ✓ Text rotation');
console.log('  ✓ Shrink to fit');
console.log('  ✓ Individual borders (top, bottom, left, right)');
console.log('  ✓ Individual border colors');
console.log('  ✓ Fill patterns');
console.log('  ✓ Foreground and background colors');
console.log('');
console.log('Open advanced-styling.xlsx to see all the features!');
