# rusc-xlsx

[![Test](https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter/actions/workflows/test.yml/badge.svg)](https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter/actions/workflows/test.yml)
[![npm version](https://img.shields.io/npm/v/@nazarkalytiuk/rusc-xlsx.svg)]([https://www.npmjs.com/package/rusc-xlsx](https://www.npmjs.com/package/@nazarkalytiuk/rusc-xlsx))
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Fast, feature-complete Excel XLSX library for Node.js powered by [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) and NAPI-RS.

## Features

### Core Excel Features
- **Workbook & Worksheets**: Multiple sheets, naming, read-only mode
- **Data Types**: Strings, numbers, booleans, dates/times
- **Formatting**:
  - Font styling (bold, italic, underline, strikethrough, size, color, script)
  - Alignment (horizontal, vertical, rotation, indent, wrap, shrink)
  - Borders (12 styles, individual sides with colors, diagonal)
  - Fills (18 patterns, foreground/background colors)
  - Number formats (currency, percentage, dates, custom)
  - Cell protection (locked/unlocked, hidden formulas)
- **Layout**: Column widths, row heights, merged cells
- **Formulas**: Standard, array, and dynamic formulas
- **Data Validation**: Dropdowns, number ranges, custom rules with messages
- **Conditional Formatting**:
  - Color scales (2-color, 3-color)
  - Data bars
  - Cell rules (greaterThan, lessThan, between, etc.)
  - Text rules (contains, beginsWith, endsWith)
  - Top/Bottom N values
  - Above/Below average
  - Duplicate/Unique values
  - Icon sets (18 types)
  - Date-based rules (10 types)
  - Formula-based rules
- **Charts**: 25 chart types (area, bar, column, line, pie, scatter, radar, stock)
- **Images**: PNG, JPG, GIF, BMP with scaling, positioning, and decorative mode
- **Tables**: Excel tables with custom columns, total rows, styles (60 styles)
- **Sparklines**: Line, column, win/lose sparklines with 36 styles
- **Rich Text**: Multi-format strings within cells
- **Notes**: Cell comments with custom styling
- **Hyperlinks**: External URLs and internal sheet links
- **Page Setup**: Margins, orientation, paper size, headers/footers, print area, scaling
- **View Controls**: Autofilters, freeze panes, zoom, tab colors, hide/show sheets

### Advanced Features
- **Named Ranges**: Global and worksheet-scoped named ranges
- **VBA Macros**: Add VBA projects from binary files
- **Form Controls**: Interactive buttons with macro assignments
- **Shapes**: Textbox annotations with positioning
- **Document Metadata**: Author, title, subject, company, manager, keywords, etc.
- **Protection**: Worksheet protection with granular permissions
- **Memory Modes**: Constant and low memory modes for massive datasets

### Performance & Memory
- **Fast**: Native Rust performance via NAPI-RS
- **Memory Efficient**: Three memory modes:
  - **Standard**: Full features, optimized for typical datasets
  - **Constant Memory**: ~0.02 MB regardless of size (millions of rows)
  - **Low Memory**: Optimized for repeated strings
- **TypeScript**: Complete type definitions with IntelliSense support

## Installation

```bash
npm install @nazarkalytiuk/rusc-xlsx
```

### Requirements

- Node.js >= 16
- Rust toolchain (for building from source)

## Quick Start

```javascript
const { Workbook, Format } = require('@nazarkalytiuk/rusc-xlsx');

// Create workbook and worksheet
const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

// Write data with formatting
const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('blue')
  .setFontColor('white');

worksheet.write(0, 0, 'Name', headerFormat);
worksheet.write(0, 1, 'Age', headerFormat);
worksheet.write(1, 0, 'Alice');
worksheet.write(1, 1, 30);

// Save file
workbook.save('output.xlsx');
```

## Memory Modes for Large Datasets

### Standard Mode (Default)

```javascript
const worksheet = workbook.addWorksheet('Sheet1');
```

- **Use when**: < 100,000 rows
- **Memory**: Scales with data size
- **Features**: Full Excel feature support

### Constant Memory Mode

```javascript
const worksheet = workbook.addWorksheetWithConstantMemory('Large');
```

- **Use when**: Millions of rows with memory constraints
- **Memory**: Constant ~0.02 MB (regardless of size!)
- **Limitations**: Sequential row writing only, no tables

> **Important:** For large exports, use `save(filename)` or `await saveToStream(writable)` instead of `saveToBuffer()`. `saveToStream()` is temp-file-backed: it saves to the operating system's temporary directory, pipes the file into the writable, waits for completion, and removes the temporary file. It avoids buffering the entire XLSX file in JavaScript memory, but it does use temporary disk I/O.

**Example: 1 million rows**
```javascript
const workbook = new Workbook();
const worksheet = workbook.addWorksheetWithConstantMemory('Data');
worksheet.setDefaultRowHeight(15.001); // Set once instead of once per row

for (let row = 0; row < 1000000; row++) {
  worksheet.write(row, 0, `Row ${row}`);
  worksheet.write(row, 1, row * 100);
}

workbook.save('large.xlsx'); // Use save(), not saveToBuffer()
```

### Low Memory Mode

```javascript
const worksheet = workbook.addWorksheetWithLowMemory('Sales');
```

- **Use when**: Large datasets with repeated strings
- **Memory**: Scales with unique strings
- **Best for**: Categorical data, status fields, repeated names

## TypeScript Support

Full TypeScript definitions included:

```typescript
import { Workbook, Format } from '@nazarkalytiuk/rusc-xlsx';

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Data');

const format = new Format();
format.setFontColor('#FF0000');  // Hex colors supported
format.setAlign('center');        // IntelliSense shows valid values
format.setBorder('medium');       // Type-safe

worksheet.write(0, 0, 'Hello', format);
workbook.save('output.xlsx');
```

## API Overview

### Workbook
- `new Workbook()` - Create workbook
- `addWorksheet([name])` - Add worksheet
- `addWorksheetWithConstantMemory([name])` - Constant memory mode
- `addWorksheetWithLowMemory([name])` - Low memory mode
- `setTempdir(path)` - Set temp directory for memory modes
- `save(filename)` - Save to file
- `saveToBuffer()` - Save to Buffer
- `saveToStream(writable)` - Save through a temp file and await stream completion
- `defineName(name, formula)` - Define named range
- `readOnlyRecommended()` - Mark as read-only recommended
- `addVbaProject(path)` - Add VBA macros
- `setProperties(props)` - Set document metadata

### Worksheet
- **Writing**: `write()`, `writeString()`, `writeNumber()`, `writeBoolean()`, `writeBlank()`
- **DateTime**: `writeDatetime()`, `writeDatetimeFromTimestamp()`, `writeDatetimeFromYmd()`, etc.
- **Formulas**: `writeFormula()`, `writeArrayFormula()`, `writeDynamicFormula()`
- **Layout**: `setColumnWidth()`, `setRowHeight()`, `setDefaultRowHeight()`, `mergeRange()`, `setName()`
- **Features**: `insertNote()`, `addDataValidation()`, `insertChart()`, `insertImage()`, `addTable()`, `insertButton()`, `insertShape()`
- **Conditional Formatting**: 14 methods for all rule types
- **Page Setup**: 26 methods for printing, margins, headers, etc.
- **View**: `autofilter()`, `setFreezePanes()`, `setZoom()`, `setTabColor()`, `hideSheet()`
- **Hyperlinks**: `writeUrl()`
- **Protection**: `protect()`, `protectWithPassword()`
- **Rich Text**: `writeRichString()`

### Format (39 methods)
- **Font**: `setBold()`, `setItalic()`, `setFontSize()`, `setFontColor()`, `setFontName()`, `setUnderline()`, `setStrikethrough()`, `setFontScript()`
- **Alignment**: `setAlign()`, `setVerticalAlign()`, `setRotation()`, `setIndent()`, `setTextWrap()`, `setShrink()`
- **Borders**: `setBorder()`, `setBorderTop/Bottom/Left/Right()`, `setBorderColor()`, `setBorderDiagonal()`
- **Fill**: `setBackgroundColor()`, `setForegroundColor()`, `setPattern()`
- **Number**: `setNumFormat()`
- **Protection**: `setLocked()`, `setUnlocked()`, `setHidden()`
- **Other**: `setReadingDirection()`, `setQuotePrefix()`

### Supporting Classes
- **Chart**: 25 chart types with series, titles, legends, 48 styles
- **Image**: Insert with scaling, positioning, alt text, URLs
- **Table**: Custom columns, total rows, 60 styles
- **Sparkline**: 3 types, markers, colors, 36 styles
- **DataValidation**: All validation types with custom messages
- **Note**: Comments with author, size, color, visibility
- **Formula**: Standard, array, dynamic with preset results
- **RichText**: Multi-format text segments
- **ProtectionOptions**: 16 granular permissions
- **DocProperties**: 10 metadata fields
- **Button**: Form controls with VBA macros
- **Shape**: Textbox shapes for annotations

## Examples

See the `examples/` directory for complete examples:
- `basic.js` - Simple workbook creation
- `formatting.js` - Cell formatting and number formats
- `data-types.js` - All supported data types
- `datetime.js` - Date/time writing methods and formats
- `multiple-sheets.js` - Working with multiple worksheets
- `advanced-styling.js` - Fonts, borders, patterns, fills
- `advanced-features.js` - Formulas, data validation, hyperlinks, freeze panes
- `charts.js` - Column, line, pie, bar, and scatter charts
- `tables.js` - Excel tables with totals, formulas, and styles
- `images.js` - Inserting images from buffers with scaling and links
- `sparklines.js` - Line, column, and win/lose sparklines
- `rich-text.js` - Multi-format text within cells
- `notes.js` - Cell comments with custom styling
- `conditional-formatting-and-page-setup.js` - Conditional formatting and print setup
- `protection.js` - Worksheet protection with granular permissions
- `doc-properties.js` - Document metadata (author, title, etc.)
- `constant-memory.js` - Large datasets with constant memory mode
- `low-memory.js` - Large datasets with low memory mode
- `save-to-buffer.js` - Saving to Buffer for HTTP responses
- `typescript-example.ts` - Full TypeScript usage
- `type-safety-demo.ts` - TypeScript type safety features

## Building from Source

```bash
# Clone repository
git clone https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter.git
cd napi_rust_xlsxwrighter

# Install dependencies
npm install

# Build native module
npm run build

# Run tests (577 tests)
npm test

# Run examples
node examples/basic.js
```

## Performance

### Fast Native Performance

Built with Rust and NAPI-RS for native-level performance:
- **Fast execution**: Native Rust performance via NAPI bindings
- **Low memory usage**: Optimized memory footprint across all modes
- **Test performance**: All 577 tests complete in < 1 second

### Memory Modes

**Memory Efficiency** (50% strings / 50% numbers):

| Rows | Standard | Low Memory | Constant Memory |
|------|----------|------------|-----------------|
| 100K | 18.0 MB | 3.0 MB | 0.0215 MB |
| 1M | 216.8 MB | 41.7 MB | 0.0215 MB |
| 10M | ~2.2 GB | ~420 MB | 0.0215 MB |

**Mode Recommendations:**
- **< 1K rows**: Default mode (fastest)
- **1K-100K rows**: Low memory mode (best balance)
- **> 100K rows**: Constant memory mode (minimal footprint)
- **Serverless/Cloud**: Low or constant memory modes

### Pure Rust vs Node.js Benchmark

Run an apples-to-apples comparison between direct `rust_xlsxwriter` calls and
this package through N-API:

```bash
npm run benchmark
```

The default workload generates exactly 200,000 rows × 20 columns (4 million
cells), split evenly between strings and numbers, in constant-memory mode. It
builds both implementations in release mode, alternates their execution order
across three runs, reports median throughput and peak RSS, and verifies that
both worksheets have identical XML size and CRC.

Override the workload or reuse existing release builds:

```bash
npm run benchmark -- --rows 100000 --columns 20 --runs 5
npm run benchmark -- --skip-build
```

The default comparison uses `worksheet.write()` once per cell. To measure the
lower-overhead bulk APIs, use row-by-row or bounded batches:

```bash
npm run benchmark -- --node-write-mode row
npm run benchmark -- --node-write-mode batch --batch-rows 1000
```

`writeRows()` in bounded batches is the fastest option for large datasets.
Larger batches reduce JavaScript-to-Rust calls but temporarily retain more
JavaScript values, so tune `--batch-rows` for the available memory.

For database cursors and other row-by-row streams, use `writeRow()` so each
cursor item crosses the JavaScript-to-Rust boundary once:

```javascript
let rowIndex = 0
for await (const item of cursor) {
  worksheet.writeRow(rowIndex++, 0, [
    item.id,
    item.name,
    item.amount,
    item.status,
  ])
}
```

This keeps memory bounded to the current cursor item while retaining
constant-memory worksheet generation.

## Limitations

- ❌ **Read-only**: Cannot read existing Excel files (rust_xlsxwriter is write-only)
- ❌ **Memory mode constraints**: Constant/low memory modes require sequential row writing

## Contributing

Contributions are welcome! Please submit a Pull Request.

## License

MIT

## Acknowledgments

- [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) - The excellent underlying Rust library
- [NAPI-RS](https://napi.rs/) - For making Rust/Node.js bindings seamless

## Links

- [GitHub Repository](https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter)
- [rust_xlsxwriter Documentation](https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/)
- [NAPI-RS Documentation](https://napi.rs/)
