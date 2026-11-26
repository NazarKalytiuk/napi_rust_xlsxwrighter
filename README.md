# rusc-xlsx

[![Test](https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter/actions/workflows/test.yml/badge.svg)](https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter/actions/workflows/test.yml)
[![npm version](https://img.shields.io/npm/v/@nazarkalytiuk/rusc-xlsx.svg)]([https://www.npmjs.com/package/rusc-xlsx](https://www.npmjs.com/package/@nazarkalytiuk/rusc-xlsx))
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Fast, feature-complete Excel XLSX library for Node.js and browsers, powered by [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) and NAPI-RS. In browsers it runs on WebAssembly; see [Browsers](#browsers-webassembly).

## Features

### Core Excel Features
- **Workbook & Worksheets**: Multiple sheets, naming, read-only mode
- **Data Types**: Strings, numbers, booleans, dates/times
- **Formatting**:
  - Font styling (bold, italic, underline, strikethrough, size, color, script)
  - Colors: named, `#RRGGBB`, or theme palette colors such as `'theme:4:2'` (formats and tab colors)
  - Alignment (horizontal, vertical, rotation, indent, wrap, shrink)
  - Borders (13 styles including hair, individual sides with colors, diagonal)
  - Fills (18 patterns, foreground/background colors)
  - Number formats (currency, percentage, dates, custom)
  - Cell protection (locked/unlocked, hidden formulas)
- **Layout**: Column widths, row heights, column/row formats, hidden rows and columns, merged cells
- **Outlines**: Grouped rows and columns, nested and collapsed
- **Formatting after writing**: Format cells and ranges (with an outside border) without rewriting their data
- **Formulas**: Standard, array, and dynamic formulas
- **Data Validation**: Dropdowns, number ranges, text length, custom rules with messages
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
- **Notes**: Cell comments with custom styling and fonts
- **Hyperlinks**: External URLs and internal sheet links
- **Page Setup**: Margins, orientation, paper size, headers/footers, print area, scaling
- **View Controls**: Autofilters, freeze panes, zoom, tab colors, screen gridlines, hide/show sheets

### Advanced Features
- **Named Ranges**: Global and worksheet-scoped named ranges
- **VBA Macros**: Add VBA projects from binary files
- **Form Controls**: Interactive buttons with macro assignments
- **Shapes**: Textboxes with solid or pattern fills, border lines, fonts, alignment and rotated text
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
- **Browsers**: The same API on WebAssembly, in the page or in a Web Worker

## Installation

```bash
npm install @nazarkalytiuk/rusc-xlsx
```

### Requirements

- Node.js >= 22.13
- Rust >= 1.95 (only when building from source)

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

## Browsers (WebAssembly)

Bundlers that follow the `browser` field of `package.json` (Vite, webpack,
esbuild, Rollup with `@rollup/plugin-node-resolve`) pick a browser build of
the same API. It runs on the WebAssembly package
`@nazarkalytiuk/rusc-xlsx-wasm32-wasip1`, which npm installs with the
library. Vite needs no configuration, in development or in production builds.

```javascript
import { Format, Workbook } from '@nazarkalytiuk/rusc-xlsx';

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Report');
worksheet.writeRow(0, 0, ['Name', 'Amount'], new Format().setBold());
worksheet.writeRow(1, 0, ['Alice', 42]);

const blob = new Blob([workbook.saveToBuffer()], {
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
});
const link = document.createElement('a');
link.href = URL.createObjectURL(blob);
link.download = 'report.xlsx';
link.click();
```

What differs from Node.js:

- Browsers have no file system: use `saveToBuffer()` instead of `save()` or
  `saveToStream()`, and `Image.fromBuffer(new Uint8Array(await response.arrayBuffer()))`
  instead of `new Image(path)`. Methods that take a path (`save`,
  `useCustomTheme`, `addVbaProject`, `setTempdir`) throw, and so do
  `addWorksheetWithConstantMemory()` and `addWorksheetWithLowMemory()`, which
  need temporary files; use the default mode.
- `saveToBuffer()` returns a `Buffer` from the
  [`buffer`](https://www.npmjs.com/package/buffer) package, a `Uint8Array`
  subclass with the Node.js `Buffer` methods.
- The module loads with top-level await, so the bundle must target browsers
  that support it (current Chrome, Edge, Firefox and Safari do, and so do
  Vite's default build targets).
- It runs on any page, without cross-origin isolation headers, and in Web
  Workers, which keeps a large export off the main thread.
- TypeScript: the declarations mention Node.js types (`Buffer`, `Writable`),
  so a browser project needs `@types/node` installed or `skipLibCheck: true`
  (the default in Vite's templates).

On Node.js the native addon is used. Where no prebuilt binary exists for the
platform, Node.js falls back to the same WebAssembly package automatically;
set `NAPI_RS_WASI_FLAVOR=wasm32-wasip1` to choose it on purpose. Use Node.js 24
or newer for that: on Node.js 22 the process crashes (SIGSEGV in `node:wasi`'s
`fd_write`) when the WebAssembly build writes a file, as `save()` and the
constant and low memory modes do.

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
- **Limitations**: Sequential row writing only, no tables, no row groups (`groupRows()` throws)

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
- **Limitations**: Sequential row writing only, no row groups (`groupRows()` throws)

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
- `setDefaultFormat(format, rowHeight, columnWidth)` - Set workbook-wide default formatting before adding worksheets
- `useExcel2023Theme()` / `useCustomTheme(path)` - Select the workbook theme before adding worksheets
- `setCalculationMode(mode)` - Formula recalculation: `'automatic'`, `'manual'` or `'partial'`
- `setZipCompressionLevel(level)` - Trade file size for speed (1-9, default 6)
- `useZipLargeFile(enable)` - Allow files over the 4 GB zip limit (ZIP64)
- `setTabRatio(ratio)` - Width of the sheet tab bar vs. the horizontal scroll bar (0-100)
- `save(filename)` - Save to file
- `saveToBuffer()` - Save to Buffer
- `saveToStream(writable)` - Save through a temp file and await stream completion
- `defineName(name, formula)` - Define named range
- `readOnlyRecommended()` - Mark as read-only recommended
- `addVbaProject(path)` - Add VBA macros
- `setProperties(props)` - Set document metadata

### Worksheet
- **Writing**: `write()`, `writeRow()`, `writeRows()` (strings, numbers, booleans and `Date`s), `writeString()`, `writeNumber()`, `writeBoolean()`, `writeBlank()`
- **DateTime**: `writeDatetime()`, `writeDatetimeFromTimestamp()`, `writeDatetimeFromYmd()`, etc. A datetime written without a number format is shown with Excel's built-in date, date-time or time format, in the viewer's locale
- **Formulas**: `writeFormula()`, `writeArrayFormula()`, `writeDynamicFormula()`
- **Layout**: `setColumnWidth()`, `setColumnWidthPixels()`, `setColumnRangeWidth()`, `setColumnAutofitWidth()`, `setRowHeight()`, `setRowHeightPixels()`, `setDefaultRowHeight()`, `mergeRange()`, `setName()`, `autofit()`, `setAutofitMaxRow()`, `setAutofitMaxWidth()`
- **Rows & columns**: `setColumnFormat()`, `setColumnRangeFormat()`, `setColumnHidden()`, `setColumnRangeHidden()`, `setRowFormat()`, `setRowHidden()`, `setRowUnhidden()`, `hideUnusedRows()`
- **Outlines**: `groupRows()`, `groupRowsCollapsed()`, `groupColumns()`, `groupColumnsCollapsed()`, `groupSymbolsAbove()`, `groupSymbolsToLeft()`
- **Cell formatting**: `setCellFormat()`, `setRangeFormat()`, `setRangeFormatWithBorder()`, `clearCell()`, `clearCellFormat()`
- **Features**: `insertNote()`, `addDataValidation()`, `insertChart()`, `insertImage()`, `insertImageFitToCellCentered()`, `addTable()`, `insertButton()`, `insertShape()`
- **Conditional Formatting**: 14 methods for all rule types
- **Page Setup**: 26 methods for printing, margins, headers, etc.
- **View**: `autofilter()`, `setFreezePanes()`, `setZoom()`, `setTabColor()`, `setScreenGridlines()`, `setHiddenZeroes()`, `setHidden()`
- **Hyperlinks**: `writeUrl()`
- **Warnings**: `ignoreError()`, `ignoreErrorRange()` turn off Excel warnings such as "number stored as text"
- **Protection**: `protect()`, `protectWithPassword()`
- **Rich Text**: `writeRichString()`

### Format (40 methods)
- **Font**: `setBold()`, `setItalic()`, `setFontSize()`, `setFontColor()`, `setFontName()`, `setFontScheme()`, `setUnderline()`, `setStrikethrough()`, `setFontScript()`
- **Alignment**: `setAlign()`, `setVerticalAlign()`, `setRotation()`, `setIndent()`, `setTextWrap()`, `setShrink()`
- **Borders**: `setBorder()`, `setBorderTop/Bottom/Left/Right()`, `setBorderColor()`, `setBorderDiagonal()`
- **Fill**: `setBackgroundColor()`, `setForegroundColor()`, `setPattern()`
- **Number**: `setNumFormat()`
- **Protection**: `setLocked()`, `setUnlocked()`, `setHidden()`
- **Other**: `setReadingDirection()`, `setQuotePrefix()`

### Supporting Classes
- **Chart**: 25 chart types with series, titles, legends, 48 styles
- **Image**: Insert with scaling, positioning, alt text, URLs
- **Table**: Custom columns, total rows, alternative text, 60 styles
- **Sparkline**: 3 types, markers, colors, 36 styles
- **DataValidation**: Lists, whole and decimal numbers, text length (`allowTextLength()`, `allowTextLengthBetween()`, `allowTextLengthFormula()`), custom formulas, with messages
- **Note**: Comments with author, size, color, font (`setFontName()`, `setFontSize()`), visibility
- **Formula**: Standard, array, dynamic with preset results
- **RichText**: Multi-format text segments
- **ProtectionOptions**: 16 granular permissions
- **DocProperties**: 10 metadata fields
- **Button**: Form controls with VBA macros
- **Shape**: Textboxes with `setFormat()`, `setFont()` and `setTextOptions()`
- **ShapeFormat**: `setSolidFill()`, `setPatternFill()`, `setNoFill()`, `setLine()`, `setNoLine()`
- **ShapeSolidFill** / **ShapePatternFill** / **ShapeLine**: colors (named, `#RRGGBB`, `theme:<color>:<shade>`), transparency, 48 patterns, line width and 8 dash types
- **ShapeFont**: name, size, bold, italic, underline, color
- **ShapeText**: horizontal and vertical alignment, text direction (`rotate90`, `rotate270`, `stacked`, ...)

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

# Install dependencies (also builds wrapper.js + wrapper.d.ts from wrapper.ts via tsdown)
npm install

# Build native module
npm run build

# Rebuild the JavaScript wrapper after editing wrapper.ts
npm run build:wrapper

# Build the WebAssembly module and run the suite on it, and in browsers
npm run build:wasm
npm run test:wasm
npx playwright install chromium firefox webkit
npm run test:browser

# Run tests (Vitest)
npm test
npm run test:watch      # re-run on change
npm run test:coverage   # V8 coverage report for wrapper.ts
npm run test:coverage:rust  # line coverage of src/*.rs via an instrumented addon;
                            # needs cargo-llvm-cov and `rustup component add llvm-tools`

# Format, lint (oxfmt, oxlint, rustfmt, clippy) and typecheck — same as CI
npm run check
npm run fmt             # apply oxfmt + rustfmt

# rust_xlsxwriter coverage checklist (rust-xlsxwriter-coverage-checklist.md)
npm run coverage:check  # verify test call sites, counts and the pinned version
npm run coverage:sync   # refresh them after editing tests or row statuses; after a
                        # rust_xlsxwriter upgrade it documents the new release with
                        # rustdoc JSON (`rustup toolchain install nightly-2026-11-13`)

# Run examples
node examples/basic.js
```

## Performance

### Fast Native Performance

Built with Rust and NAPI-RS for native-level performance:
- **Fast execution**: Native Rust performance via NAPI bindings
- **Low memory usage**: Optimized memory footprint across all modes
- **Test performance**: All 601 tests complete in < 1 second

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

### Node.js vs Bun Benchmarks

The runtime benchmark executes the same package workload through the same
release N-API binary on Node.js and Bun:

```bash
npm run benchmark:runtimes
```

Bun must be installed and available in `PATH`. The default workload generates
200,000 rows × 20 columns through `writeRow()` in constant-memory mode. The
runner alternates runtime order across three runs, reports both in-process
generation time and end-to-end process wall time, and verifies identical
worksheet XML size and CRC for every Node.js/Bun pair.

Override the workload or package API, or reuse the existing release build:

```bash
npm run benchmark:runtimes -- --rows 100000 --columns 20 --runs 5
npm run benchmark:runtimes -- --write-mode cell
npm run benchmark:runtimes -- --write-mode batch --batch-rows 1000
npm run benchmark:runtimes -- --skip-build
```

#### ExcelJS runtime comparison

To isolate the runtime difference for a JavaScript XLSX implementation, run the
same dataset through the ExcelJS streaming writer:

```bash
npm run benchmark:runtimes:exceljs
```

This profile uses ExcelJS `stream.xlsx.WorkbookWriter`, disables styles and
shared strings, fixes ZIP DEFLATE compression at level 6, and commits every
row. It uses the same alternating order, timing and memory metrics, and
worksheet-equivalence checks as the native-package profile. Override the
default workload when needed:

```bash
npm run benchmark:runtimes:exceljs -- --rows 100000 --columns 20 --runs 5
```

### Node.js vs Go, C#, and Python Benchmark

The cross-language benchmark compares concrete streaming XLSX libraries, not
language runtimes in isolation:

| Runtime | XLSX implementation | Bounded-memory API |
|---------|---------------------|--------------------|
| Node.js | `@nazarkalytiuk/rusc-xlsx` | Constant-memory `writeRow()` |
| Go | Excelize 2.11.0 | `StreamWriter.SetRow()` |
| C# | Open XML SDK 3.5.1 | Forward-only worksheet XML writer |
| Python | XlsxWriter 3.2.9 | `constant_memory` + `write_row()` |

Install Rust, Go 1.25+, the .NET 8 SDK, and Python 3.10+, then run:

```bash
npm run benchmark:languages
```

The runner builds release artifacts and creates an isolated Python virtual
environment under `benchmarks/python-xlsxwriter/.venv`. The default workload is
the same 200,000 rows × 20 columns, 50/50 string-and-number dataset used by the
Rust comparison. Before measurement, each implementation generates a small
workbook that is checked for its worksheet name, complete row/cell shape, and
representative values. Measured execution order is balanced across runs, and
the report includes median generation time, end-to-end process wall time,
throughput, peak RSS, and XLSX size.

Override the workload, select available runtimes, or reuse existing artifacts:

```bash
npm run benchmark:languages -- --rows 100000 --columns 20 --runs 5
npm run benchmark:languages -- --implementations node,go,python
npm run benchmark:languages -- --skip-build
```

Use `--node-write-mode cell|row|batch` to measure a different Node.js API;
`row` is the cross-language default because all four implementations consume
the workload sequentially by row.

## Limitations

- ❌ **Read-only**: Cannot read existing Excel files (rust_xlsxwriter is write-only)
- ❌ **Memory mode constraints**: Constant/low memory modes require sequential row writing and don't support row groups (rust_xlsxwriter 0.104.0 writes their rows without an outline level)
- ❌ **Shape text strikethrough**: rust_xlsxwriter 0.104.0 doesn't write it, so `ShapeFont` has no `setStrikethrough()`

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
