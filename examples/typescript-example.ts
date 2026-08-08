import {
  Workbook,
  Worksheet,
  Format,
  Note,
  Color,
  HorizontalAlign,
  VerticalAlign,
  BorderStyle,
  UnderlineStyle,
  FillPattern,
  FontScheme
} from '@nazarkalytiuk/rusc-xlsx';

// Type-safe example demonstrating IntelliSense support with fluent API

const workbook = new Workbook();
workbook.useExcel2023Theme();
const worksheet: Worksheet = workbook.addWorksheet('TypeScript Demo');

// Font styling with type safety and fluent API
const headerFormat = new Format()
  .setBold()
  .setItalic()
  .setFontSize(16)
  .setFontName('Arial');
const fontScheme: FontScheme = 'none';
headerFormat.setFontScheme(fontScheme);

// Color type ensures valid values
const validColor: Color = '#4472C4';
headerFormat.setFontColor(validColor);
headerFormat.setBackgroundColor('blue'); // Named colors also work

// Alignment with type safety
const centerAlign: HorizontalAlign = 'center';
headerFormat.setAlign(centerAlign);

const verticalCenter: VerticalAlign = 'center';
headerFormat.setVerticalAlign(verticalCenter);

// Border styles with type safety
const borderStyle: BorderStyle = 'medium';
headerFormat.setBorder(borderStyle);
headerFormat.setBorderColor('white');

// Individual borders with different styles
headerFormat.setBorderTop('thick');
headerFormat.setBorderTopColor('red');
headerFormat.setBorderBottom('dashed');
headerFormat.setBorderBottomColor('blue');

// Write header with formatting
worksheet.setRowHeight(0, 30);
worksheet.write(0, 0, 'Product', headerFormat);
worksheet.write(0, 1, 'Price', headerFormat);
worksheet.write(0, 2, 'Stock', headerFormat);

// Underline with type safety and fluent API
const underlineStyle: UnderlineStyle = 'double';
const underlineFormat = new Format()
  .setUnderline(underlineStyle)
  .setFontColor('purple');

worksheet.write(1, 0, 'Special Item', underlineFormat);

// Pattern fills with type safety and fluent API
const pattern: FillPattern = 'mediumGray';
const patternFormat = new Format()
  .setPattern(pattern)
  .setBackgroundColor('#D9E1F2')
  .setForegroundColor('#4472C4');

worksheet.write(2, 0, 'Patterned Cell', patternFormat);

// Text formatting with fluent API
const textFormat = new Format()
  .setTextWrap()
  .setIndent(2)
  .setRotation(45);

worksheet.setRowHeight(3, 50);
worksheet.write(3, 0, 'Wrapped and rotated text', textFormat);

// Number formatting with fluent API
const currencyFormat = new Format()
  .setNumFormat('$#,##0.00')
  .setFontColor('green')
  .setBold();

worksheet.write(1, 1, 1234.56, currencyFormat);

const percentFormat = new Format()
  .setNumFormat('0.00%');

worksheet.write(2, 2, 0.85, percentFormat);

// Notes with type safety and fluent API
const note = new Note('This is an important note')
  .setAuthor('TypeScript User')
  .setWidth(200)
  .setHeight(100)
  .setBackgroundColor('#FFFFE1') // Hex color
  .setVisible(true)
  .setAltText('Important note for accessibility');

worksheet.insertNote(1, 0, note);

// Cell protection with fluent API
const protectedFormat = new Format()
  .setUnlocked()
  .setHidden();

worksheet.write(4, 0, 'Protected cell', protectedFormat);

// Column and row sizing
worksheet.setColumnWidth(0, 20);
worksheet.setColumnWidth(1, 15);
worksheet.setColumnWidth(2, 15);

// Merge cells with fluent API
const mergeFormat = new Format()
  .setBold()
  .setAlign('center')
  .setVerticalAlign('center')
  .setBackgroundColor('#E7E6E6');

worksheet.setRowHeight(5, 25);
worksheet.mergeRange(5, 0, 5, 2, 'Merged Header', mergeFormat);

worksheet.setAutofitMaxRow(200);
worksheet.setAutofitMaxWidth(300);
worksheet.autofit();

// Multiple worksheets
const dataSheet = workbook.addWorksheet('Data');
dataSheet.setDefaultNoteAuthor('Data Team');

// Memory-optimized worksheets
const largeDataSheet = workbook.addWorksheetWithConstantMemory('Large Data');
const repeatedDataSheet = workbook.addWorksheetWithLowMemory('Repeated Data');

// Save to file
workbook.save('typescript-demo.xlsx');

// Or save to buffer
const buffer: Buffer = workbook.saveToBuffer();

console.log('TypeScript example completed!');
console.log('File: typescript-demo.xlsx');
console.log('Buffer size:', buffer.length, 'bytes');
console.log('');
console.log('✓ Full type safety with IntelliSense support');
console.log('✓ Type unions for valid values (colors, alignments, borders, etc.)');
console.log('✓ JSDoc documentation for all methods');
console.log('✓ Parameter validation at compile time');
