/**
 * TypeScript Type Safety Demonstration
 *
 * This file demonstrates that the TypeScript definitions provide
 * proper type checking and IntelliSense support.
 */

import { Workbook, Format, Note, Color, BorderStyle, HorizontalAlign } from '@nazarkalytiuk/rusc-xlsx';

// ✅ VALID CODE - These should compile without errors

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Demo');

// Valid color values
const validColor1: Color = 'red';
const validColor2: Color = '#FF0000';
const validColor3: Color = 'blue';

// Valid border styles
const validBorder1: BorderStyle = 'thin';
const validBorder2: BorderStyle = 'dashDot';

// Valid alignment
const validAlign: HorizontalAlign = 'center';

// Fluent API with method chaining
const format = new Format()
  .setBold()                          // Returns Format
  .setItalic()                        // Returns Format
  .setFontColor(validColor1)          // Type-safe color
  .setBorder(validBorder1)            // Type-safe border
  .setAlign(validAlign)               // Type-safe alignment
  .setFontSize(14)                    // Number
  .setRotation(45);                   // Number with validation

// Note with fluent API
const note = new Note('Test')
  .setAuthor('Author')                // Returns Note
  .setWidth(200)                      // Returns Note
  .setVisible(true);                  // Returns Note

// Write with type-safe values
worksheet.write(0, 0, 'String');      // string
worksheet.write(0, 1, 123);           // number
worksheet.write(0, 2, true);          // boolean
worksheet.write(0, 3, 'Formatted', format);

worksheet.insertNote(0, 0, note);

workbook.save('type-safe.xlsx');

// ❌ INVALID CODE - These should cause TypeScript errors
// Uncomment to see type checking in action:

// const invalidColor: Color = 'purple'; // Error: not a valid color
// const invalidBorder: BorderStyle = 'wavy'; // Error: not a valid border
// format.setFontColor('invalid-color'); // Error: string is not assignable to Color
// format.setBorder('invalid-border');   // Error: string is not assignable to BorderStyle
// worksheet.write(0, 0, []);            // Error: array is not string | number | boolean
// note.setWidth('200');                 // Error: string is not assignable to number

console.log('✓ Type safety demo completed successfully!');
console.log('✓ All type checks passed');
console.log('✓ Method chaining works with proper return types');
