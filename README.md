# rusc-xlsx

Fast Excel XLSX library for Node.js powered by [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) and NAPI-RS.

## Status

**Build Status**: ✅ STABLE
**Test Coverage**: ✅ **248/248 tests passing** (+38 new tests)
**Phase 0 (Core Foundation)**: ✅ 100% Complete
**Advanced Features**: ✅ **100% COVERAGE** for Conditional Formatting (21/21) & Page Setup (20/20)
**Last Updated**: 2025-11-06 (v0.7.0)

> **🎯 MILESTONE: 100% coverage achieved for Conditional Formatting and Page Setup!**
>
> All core features plus 21 conditional formatting types and 20 page setup methods are fully implemented and tested. See [FEATURES.md](FEATURES.md) for detailed breakdown.

## Features

- **Fast Performance**: Native Rust performance for Excel file generation
- **Memory Efficient**: Support for constant and low memory modes for large datasets
- **Simple API**: Clean, intuitive JavaScript API
- **TypeScript Support**: Full TypeScript type definitions included
- **Core Excel Features**:
  - Create workbooks with multiple worksheets
  - Write strings, numbers, and booleans
  - **Advanced Cell Formatting**:
    - Font styling (bold, italic, underline, strikethrough, font name, size, color)
    - Text alignment (horizontal, vertical, rotation, indent, wrap, shrink-to-fit)
    - Borders (individual sides with styles and colors)
    - Fills (solid colors, patterns, foreground/background)
    - Number formats (currency, percentage, dates, custom)
  - Column widths and row heights
  - Merge cells
  - Cell notes (comments) with customization
  - Cell protection (locked/unlocked, hidden formulas)
  - **Formulas**: Standard, array, and dynamic formulas
  - **Data Validation**: Dropdowns, number ranges, custom rules
  - **Conditional Formatting**: Color scales, data bars, cell/text rules
  - **Hyperlinks**: External URLs and internal sheet links
  - **Page Setup**: Margins, orientation, headers/footers, print area
  - **View Controls**: Autofilters and freeze panes
  - Save to file or buffer
- **Memory Modes**:
  - Constant memory mode for minimal memory usage (~0.02 MB regardless of dataset size)
  - Low memory mode for balanced memory/file-size optimization
  - Handle millions of rows efficiently

## Installation

```bash
npm install rusc-xlsx
```

### Requirements

- Node.js >= 16
- Rust toolchain (for building from source)

## TypeScript Support

This library includes comprehensive TypeScript type definitions with:

- ✅ Full IntelliSense support in VS Code and other editors
- ✅ Type-safe enums for colors, alignments, borders, and patterns
- ✅ JSDoc documentation for all methods
- ✅ Compile-time parameter validation

```typescript
import { RuscWorkbook, RuscFormat, Color, BorderStyle } from 'rusc-xlsx';

const workbook = new RuscWorkbook();
const worksheet = workbook.addWorksheet('Data');

const format = new RuscFormat();
format.setFontColor('#FF0000');  // Hex colors supported
format.setAlign('center');        // IntelliSense shows valid values
format.setBorder('medium');       // Type-safe border styles

worksheet.write(0, 0, 'Hello', format);
workbook.save('output.xlsx');
```

See `examples/typescript-example.ts` for a complete TypeScript example.

## Quick Start

```javascript
const { Workbook, Format } = require('rusc-xlsx');

// Create a new workbook
const workbook = new Workbook();

// Add a worksheet
const worksheet = workbook.addWorksheet('Sheet1');

// Write some data
worksheet.write(0, 0, 'Hello');
worksheet.write(0, 1, 'World');
worksheet.write(1, 0, 'Number:');
worksheet.write(1, 1, 42);

// Save the workbook
workbook.save('output.xlsx');
```

## Memory Modes for Large Datasets

rusc-xlsx supports three memory modes for handling datasets of different sizes:

### Standard Mode (Default)

```javascript
const worksheet = workbook.addWorksheet('Sheet1');
```

- **Use when**: Working with typical datasets (< 100,000 rows)
- **Memory**: Scales with data size
- **Features**: Full Excel feature support
- **File Size**: Optimized

### Constant Memory Mode

```javascript
const worksheet = workbook.addWorksheetWithConstantMemory('Large');
```

- **Use when**: Processing millions of rows with memory constraints
- **Memory**: Constant ~0.02 MB regardless of dataset size
- **Features**: Limited (see below)
- **File Size**: Larger (strings stored inline)

**Important Limitations:**
- Data MUST be written sequentially row-by-row
- Cannot write to previous rows
- `mergeRange()` only works for current row
- `addTable()` not supported

**Memory Comparison (50% strings / 50% numbers):**

| Rows | Standard | Low Memory | Constant Memory |
|------|----------|------------|-----------------|
| 100K | 18.0 MB | 3.0 MB | 0.0215 MB |
| 1M | 216.8 MB | 41.7 MB | 0.0215 MB |

**Example:**

```javascript
const { Workbook } = require('rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheetWithConstantMemory('VeryLarge');

// Write 1 million rows sequentially
for (let row = 0; row < 1000000; row++) {
  worksheet.write(row, 0, `Row ${row}`);
  worksheet.write(row, 1, row * 100);
  worksheet.write(row, 2, row % 2 === 0);
}

workbook.save('large.xlsx');
```

### Low Memory Mode

```javascript
const worksheet = workbook.addWorksheetWithLowMemory('Large');
```

- **Use when**: Large datasets with repeated strings (categories, status values)
- **Memory**: Scales with unique strings
- **Features**: Limited (same as constant memory)
- **File Size**: Smaller (uses shared string table)

**Best for datasets with:**
- Categorical data
- Status fields
- Repeated product names
- Region/location data

**Example:**

```javascript
const workbook = new Workbook();
const worksheet = workbook.addWorksheetWithLowMemory('Sales');

const regions = ['North', 'South', 'East', 'West'];
const products = ['Widget A', 'Widget B', 'Widget C'];

// These repeated strings are stored efficiently
for (let row = 0; row < 500000; row++) {
  worksheet.write(row, 0, row);
  worksheet.write(row, 1, regions[row % regions.length]);
  worksheet.write(row, 2, products[row % products.length]);
}

workbook.save('sales.xlsx');
```

### Custom Temp Directory

When using memory modes, you can set a custom temporary directory:

```javascript
const workbook = new Workbook();

// Must be called BEFORE adding worksheets
workbook.setTempdir('./temp');

const worksheet = workbook.addWorksheetWithConstantMemory();
```

### Choosing the Right Mode

| Scenario | Recommended Mode |
|----------|------------------|
| < 100K rows, full features needed | Standard |
| Large dataset, unique strings | Constant Memory |
| Large dataset, repeated strings | Low Memory |
| Memory is critical | Constant Memory |
| File size is critical | Low Memory |
| Need tables/complex merges | Standard |

## API Reference

### Workbook

#### `new Workbook()`

Creates a new workbook.

```javascript
const workbook = new Workbook();
```

#### `workbook.addWorksheet([name])`

Adds a worksheet to the workbook. Returns a `Worksheet` instance.

- `name` (optional): Name for the worksheet

```javascript
const sheet1 = workbook.addWorksheet();
const sheet2 = workbook.addWorksheet('Sales Data');
```

#### `workbook.addWorksheetWithConstantMemory([name])`

Adds a worksheet with constant memory mode (minimal memory footprint).

- `name` (optional): Name for the worksheet
- Returns: `Worksheet` instance in constant memory mode

**Limitations**: Sequential row writing only, limited merge support, no tables.

```javascript
const worksheet = workbook.addWorksheetWithConstantMemory('Large');
```

#### `workbook.addWorksheetWithLowMemory([name])`

Adds a worksheet with low memory mode (balanced optimization).

- `name` (optional): Name for the worksheet
- Returns: `Worksheet` instance in low memory mode

**Limitations**: Sequential row writing only, limited merge support, no tables.

```javascript
const worksheet = workbook.addWorksheetWithLowMemory('Sales');
```

#### `workbook.setTempdir(path)`

Sets custom temporary directory for constant/low memory modes.

- `path`: Directory path for temporary files

**Important**: Must be called BEFORE adding memory mode worksheets.

```javascript
workbook.setTempdir('./temp');
```

#### `workbook.save(filename)`

Saves the workbook to a file.

- `filename`: Path where the file should be saved

```javascript
workbook.save('output.xlsx');
```

#### `workbook.saveToBuffer()`

Saves the workbook to a Buffer. Returns a `Buffer`.

```javascript
const buffer = workbook.saveToBuffer();
```

### Worksheet

#### `worksheet.write(row, col, value, [format])`

Writes a value to a cell. Auto-detects the type (string, number, or boolean).

- `row`: Row number (0-indexed)
- `col`: Column number (0-indexed)
- `value`: Value to write
- `format` (optional): `Format` object to apply

```javascript
worksheet.write(0, 0, 'Text');
worksheet.write(0, 1, 42);
worksheet.write(0, 2, true);

const format = new Format().setBold();
worksheet.write(1, 0, 'Bold Text', format);
```

#### `worksheet.writeString(row, col, value)`

Writes a string to a cell.

```javascript
worksheet.writeString(0, 0, 'Hello');
```

#### `worksheet.writeNumber(row, col, value)`

Writes a number to a cell.

```javascript
worksheet.writeNumber(0, 0, 42.5);
```

#### `worksheet.writeBoolean(row, col, value)`

Writes a boolean to a cell.

```javascript
worksheet.writeBoolean(0, 0, true);
```

#### `worksheet.setColumnWidth(col, width)`

Sets the width of a column.

- `col`: Column number (0-indexed)
- `width`: Width in Excel units (approximately character width)

```javascript
worksheet.setColumnWidth(0, 20);
```

#### `worksheet.setRowHeight(row, height)`

Sets the height of a row.

- `row`: Row number (0-indexed)
- `height`: Height in points

```javascript
worksheet.setRowHeight(0, 30);
```

#### `worksheet.setName(name)`

Sets the worksheet name.

```javascript
worksheet.setName('Sales Q1');
```

#### `worksheet.mergeRange(firstRow, firstCol, lastRow, lastCol, value)`

Merges a range of cells and writes a value to it.

```javascript
worksheet.mergeRange(0, 0, 0, 3, 'Merged Header');
```

#### `worksheet.insertNote(row, col, note)`

Inserts a note (post-it style comment) to a cell.

- `row`: Row number (0-indexed)
- `col`: Column number (0-indexed)
- `note`: `Note` object to insert

```javascript
const note = new Note('This is a note');
note.setAuthor('John Doe');
worksheet.insertNote(0, 0, note);
```

#### `worksheet.setDefaultNoteAuthor(name)`

Sets the default author name for all notes in the worksheet.

- `name`: Author name (max 52 characters)

```javascript
worksheet.setDefaultNoteAuthor('John Doe');
```

### Format

#### `new Format()`

Creates a new format object. Format methods can be chained.

```javascript
const format = new Format();
```

#### `format.setBold()`

Sets bold font. Returns `this` for chaining.

```javascript
format.setBold();
```

#### `format.setItalic()`

Sets italic font. Returns `this` for chaining.

```javascript
format.setItalic();
```

#### `format.setFontSize(size)`

Sets font size in points.

```javascript
format.setFontSize(14);
```

#### `format.setFontColor(color)`

Sets font color. Accepts color names or hex codes.

- Color names: `'black'`, `'blue'`, `'brown'`, `'cyan'`, `'gray'`, `'green'`, `'lime'`, `'magenta'`, `'navy'`, `'orange'`, `'pink'`, `'purple'`, `'red'`, `'silver'`, `'white'`, `'yellow'`
- Hex codes: `'#RRGGBB'` format

```javascript
format.setFontColor('red');
format.setFontColor('#FF0000');
```

#### `format.setBackgroundColor(color)`

Sets background color. Same color format as `setFontColor()`.

```javascript
format.setBackgroundColor('yellow');
```

#### `format.setAlign(align)`

Sets horizontal alignment.

- Values: `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'centerAcross'`, `'distributed'`

```javascript
format.setAlign('center');
```

#### `format.setVerticalAlign(align)`

Sets vertical alignment.

- Values: `'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'`

```javascript
format.setVerticalAlign('center');
```

#### `format.setBorder(style)`

Sets border style for all borders.

- Values: `'thin'`, `'medium'`, `'thick'`, `'double'`, `'dotted'`, `'dashed'`

```javascript
format.setBorder('thin');
```

#### `format.setNumFormat(format)`

Sets number format using Excel format strings.

```javascript
format.setNumFormat('#,##0.00');      // Thousands separator with 2 decimals
format.setNumFormat('$#,##0.00');     // Currency format
format.setNumFormat('0.00%');         // Percentage format
format.setNumFormat('yyyy-mm-dd');    // Date format
```

#### `format.setFontName(name)`

Sets font name.

```javascript
format.setFontName('Arial');
format.setFontName('Times New Roman');
format.setFontName('Calibri');
```

#### `format.setUnderline(style)`

Sets text underline style.

- Values: `'single'`, `'double'`, `'singleAccounting'`, `'doubleAccounting'`

```javascript
format.setUnderline('single');
format.setUnderline('double');
```

#### `format.setStrikethrough()`

Sets strikethrough text.

```javascript
format.setStrikethrough();
```

#### `format.setTextWrap()`

Enables text wrapping in the cell.

```javascript
format.setTextWrap();
```

#### `format.setIndent(level)`

Sets text indentation level (0-15).

```javascript
format.setIndent(1);  // Indent by 1 level
format.setIndent(3);  // Indent by 3 levels
```

#### `format.setRotation(degrees)`

Sets text rotation in degrees.

- Values: `-90` to `90`, or `270` for vertical text

```javascript
format.setRotation(45);   // Rotate 45 degrees
format.setRotation(-30);  // Rotate -30 degrees
format.setRotation(270);  // Vertical text
```

#### `format.setShrink()`

Shrinks text to fit cell width.

```javascript
format.setShrink();
```

#### `format.setBorderTop(style)`

Sets top border style.

- Values: `'thin'`, `'medium'`, `'thick'`, `'double'`, `'dotted'`, `'dashed'`, `'dashDot'`, `'dashDotDot'`, `'slantDashDot'`, `'mediumDashed'`, `'mediumDashDot'`, `'mediumDashDotDot'`

```javascript
format.setBorderTop('thick');
```

#### `format.setBorderBottom(style)`

Sets bottom border style. Same values as `setBorderTop()`.

```javascript
format.setBorderBottom('medium');
```

#### `format.setBorderLeft(style)`

Sets left border style. Same values as `setBorderTop()`.

```javascript
format.setBorderLeft('dashed');
```

#### `format.setBorderRight(style)`

Sets right border style. Same values as `setBorderTop()`.

```javascript
format.setBorderRight('dotted');
```

#### `format.setBorderColor(color)`

Sets color for all borders. Same color format as `setFontColor()`.

```javascript
format.setBorderColor('red');
format.setBorderColor('#FF0000');
```

#### `format.setBorderTopColor(color)`

Sets top border color.

```javascript
format.setBorderTopColor('blue');
```

#### `format.setBorderBottomColor(color)`

Sets bottom border color.

```javascript
format.setBorderBottomColor('green');
```

#### `format.setBorderLeftColor(color)`

Sets left border color.

```javascript
format.setBorderLeftColor('orange');
```

#### `format.setBorderRightColor(color)`

Sets right border color.

```javascript
format.setBorderRightColor('purple');
```

#### `format.setPattern(pattern)`

Sets fill pattern.

- Values: `'solid'`, `'mediumGray'`, `'darkGray'`, `'lightGray'`, `'darkHorizontal'`, `'darkVertical'`, `'darkDown'`, `'darkUp'`, `'darkGrid'`, `'darkTrellis'`, `'lightHorizontal'`, `'lightVertical'`, `'lightDown'`, `'lightUp'`, `'lightGrid'`, `'lightTrellis'`, `'gray125'`, `'gray0625'`

```javascript
format.setPattern('solid');
format.setPattern('mediumGray');
```

#### `format.setForegroundColor(color)`

Sets foreground color (for patterns). Same color format as `setFontColor()`.

```javascript
format.setForegroundColor('#4472C4');
```

#### `format.setUnlocked()`

Unlocks cells (for use with worksheet protection).

```javascript
format.setUnlocked();
```

#### `format.setHidden()`

Hides formulas in cells (for use with worksheet protection).

```javascript
format.setHidden();
```

### Method Chaining

Format methods can be chained for convenience:

```javascript
const headerFormat = new Format()
  .setBold()
  .setFontSize(14)
  .setFontColor('white')
  .setBackgroundColor('blue')
  .setAlign('center');
```

### Note

#### `new Note(text)`

Creates a new note (cell comment) with the given text.

```javascript
const note = new Note('Remember to review this cell');
```

#### `note.setAuthor(name)`

Sets the author name for the note.

- `name`: Author name (max 52 characters)

```javascript
note.setAuthor('Jane Smith');
```

#### `note.setWidth(width)`

Sets the width of the note in pixels.

```javascript
note.setWidth(200);
```

#### `note.setHeight(height)`

Sets the height of the note in pixels.

```javascript
note.setHeight(100);
```

#### `note.setBackgroundColor(color)`

Sets the background color of the note. Same color format as `Format.setBackgroundColor()`.

```javascript
note.setBackgroundColor('yellow');
note.setBackgroundColor('#FFFFE1');
```

#### `note.setVisible(visible)`

Makes the note visible by default (normally notes are hidden until you hover over the cell).

```javascript
note.setVisible(true);
```

#### `note.setAltText(text)`

Sets alt text for accessibility.

```javascript
note.setAltText('Note about important deadline');
```

## Examples

### Basic Usage

```javascript
const { Workbook } = require('rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Data');

worksheet.write(0, 0, 'Name');
worksheet.write(0, 1, 'Age');
worksheet.write(1, 0, 'Alice');
worksheet.write(1, 1, 30);
worksheet.write(2, 0, 'Bob');
worksheet.write(2, 1, 25);

workbook.save('people.xlsx');
```

### Formatting

```javascript
const { Workbook, Format } = require('rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Formatted');

// Create formats
const headerFormat = new Format()
  .setBold()
  .setFontColor('white')
  .setBackgroundColor('blue')
  .setAlign('center');

const currencyFormat = new Format()
  .setNumFormat('$#,##0.00');

// Write headers
worksheet.write(0, 0, 'Product', headerFormat);
worksheet.write(0, 1, 'Price', headerFormat);

// Write data
worksheet.write(1, 0, 'Widget');
worksheet.write(1, 1, 29.99, currencyFormat);

workbook.save('formatted.xlsx');
```

### Multiple Worksheets

```javascript
const { Workbook } = require('rusc-xlsx');

const workbook = new Workbook();

// Create multiple sheets
const sales = workbook.addWorksheet('Sales');
const expenses = workbook.addWorksheet('Expenses');
const summary = workbook.addWorksheet('Summary');

sales.write(0, 0, 'Q1 Sales');
sales.write(1, 0, 100000);

expenses.write(0, 0, 'Q1 Expenses');
expenses.write(1, 0, 50000);

summary.write(0, 0, 'Profit');
summary.write(1, 0, 50000);

workbook.save('report.xlsx');
```

### Save to Buffer

```javascript
const { Workbook } = require('rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet();
worksheet.write(0, 0, 'Hello Buffer');

const buffer = workbook.saveToBuffer();

// Use buffer (e.g., send via HTTP, save to database, etc.)
console.log(`Generated ${buffer.length} bytes`);
```

### Cell Notes

```javascript
const { Workbook, Note } = require('rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Tasks');

// Set default author for all notes in this worksheet
worksheet.setDefaultNoteAuthor('Project Manager');

// Write some data
worksheet.write(0, 0, 'Task');
worksheet.write(0, 1, 'Status');

// Add a basic note
const note1 = new Note('Deadline: Friday');
note1.setAuthor('PM');
worksheet.insertNote(1, 0, note1);

// Add a custom styled note
const note2 = new Note('URGENT: Review this ASAP!');
note2.setAuthor('Team Lead');
note2.setWidth(250);
note2.setHeight(120);
note2.setBackgroundColor('#FFB6C1');
note2.setVisible(true); // Show by default
worksheet.insertNote(2, 0, note2);

// Add a note with alt text for accessibility
const note3 = new Note('Remember to update documentation');
note3.setAltText('Documentation reminder note');
worksheet.insertNote(3, 0, note3);

workbook.save('tasks-with-notes.xlsx');
```

### Advanced Styling

```javascript
const { Workbook, Format } = require('rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Styled');

// Create a fancy header format
const headerFormat = new Format()
    .setBold()
    .setFontName('Arial')
    .setFontSize(14)
    .setFontColor('white')
    .setBackgroundColor('#4472C4')
    .setAlign('center')
    .setVerticalAlign('center')
    .setBorder('medium')
    .setBorderColor('white');

worksheet.setRowHeight(0, 30);
worksheet.write(0, 0, 'Product', headerFormat);
worksheet.write(0, 1, 'Price', headerFormat);
worksheet.write(0, 2, 'Status', headerFormat);

// Format with underline and custom borders
const priceFormat = new Format()
    .setNumFormat('$#,##0.00')
    .setFontColor('#0B6623')
    .setBorderBottom('double')
    .setBorderBottomColor('green');

worksheet.write(1, 1, 29.99, priceFormat);

// Text with rotation
const rotatedFormat = new Format()
    .setRotation(45)
    .setBold()
    .setFontColor('red');

worksheet.setRowHeight(2, 50);
worksheet.write(2, 0, 'Rotated Text', rotatedFormat);

// Text wrap example
const wrapFormat = new Format()
    .setTextWrap()
    .setAlign('left')
    .setVerticalAlign('top')
    .setIndent(1);

worksheet.setColumnWidth(1, 30);
worksheet.setRowHeight(3, 50);
worksheet.write(3, 1, 'This is a long text that will wrap to multiple lines', wrapFormat);

// Individual border colors
const borderFormat = new Format()
    .setBorderTop('medium')
    .setBorderTopColor('red')
    .setBorderBottom('medium')
    .setBorderBottomColor('blue')
    .setBorderLeft('medium')
    .setBorderLeftColor('green')
    .setBorderRight('medium')
    .setBorderRightColor('purple');

worksheet.write(4, 0, 'Rainbow Borders', borderFormat);

workbook.save('advanced-styling.xlsx');
```

## Building from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/rusc-xlsx.git
cd rusc-xlsx

# Install dependencies
npm install

# Build the native module
npm run build

# Run tests
npm test

# Run examples
node examples/basic.js
```

## Development

### Project Structure

```
rusc-xlsx/
├── src/
│   └── lib.rs          # Rust bindings implementation
├── examples/           # Example usage files
│   ├── basic.js
│   ├── formatting.js
│   ├── multiple-sheets.js
│   └── data-types.js
├── tests/             # Test suite
│   ├── workbook.test.js
│   ├── worksheet.test.js
│   └── format.test.js
├── index.js           # JavaScript API wrapper
├── index.d.ts         # TypeScript type definitions
├── Cargo.toml         # Rust configuration
└── package.json       # Node.js configuration
```

### Running Tests

```bash
npm test
```

### Running Examples

```bash
node examples/basic.js
node examples/formatting.js
node examples/multiple-sheets.js
node examples/data-types.js
node examples/constant-memory.js  # Large dataset with constant memory
node examples/low-memory.js       # Large dataset with low memory
```

## Performance

rusc-xlsx leverages native Rust performance through Neon bindings, making it significantly faster than pure JavaScript Excel libraries for large datasets.

### Memory Efficiency

With constant memory mode, you can generate Excel files with millions of rows using only ~0.02 MB of memory:

```
Dataset Size    | Standard Mode | Low Memory Mode | Constant Memory Mode
100,000 rows    | ~18 MB        | ~3 MB           | ~0.02 MB
1,000,000 rows  | ~217 MB       | ~42 MB          | ~0.02 MB
10,000,000 rows | ~2.2 GB       | ~420 MB         | ~0.02 MB
```

This makes it possible to generate massive Excel reports even in memory-constrained environments.

## Limitations

This library currently focuses on core Excel writing functionality. Advanced features like:
- ❌ Reading existing Excel files (rust_xlsxwriter is write-only)
- ❌ Charts and images
- ❌ Excel Tables
- ❌ Sparklines
- ❌ Rich text (multi-format strings)

✅ **Now Implemented:**
- ✅ Formulas (standard, array, dynamic)
- ✅ Data validation
- ✅ Conditional formatting (color scales, data bars, cell/text rules)
- ✅ Page setup & printing

See [FEATURES.md](FEATURES.md) for a complete breakdown of implemented vs. not-yet-implemented features.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT

## Acknowledgments

- [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) - The underlying Rust library
- [Neon](https://neon-bindings.com/) - For making Rust/Node.js bindings possible

## Links

- [rust_xlsxwriter Documentation](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/)
- [Neon Documentation](https://neon-bindings.com/)
- [GitHub Repository](https://github.com/yourusername/rusc-xlsx)
