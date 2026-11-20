// Fluent, chainable JavaScript API over the native bindings in ./index.js.
// tsdown compiles this file into the published wrapper.js and wrapper.d.ts,
// and into wrapper.browser.mjs, the browser entry, which loads the WebAssembly
// build through ./browser.js instead of ./index.js.

import type { Writable } from 'node:stream';

import {
  __napiBindingTarget,
  RuscButton,
  RuscChart,
  RuscDataValidation,
  RuscDocProperties,
  RuscFormat,
  RuscFormula,
  type RuscFormulaData,
  RuscImage,
  RuscNote,
  RuscProtectionOptions,
  RuscRichText,
  RuscShape,
  RuscSparkline,
  RuscTable,
  RuscTableColumn,
  RuscWorkbook,
  RuscWorksheet,
} from './index.js';

/**
 * Returns `process` on Node.js and `undefined` in browsers.
 *
 * Node.js built-in modules are looked up with `process.getBuiltinModule()` when
 * a method that needs them runs, instead of being imported, so this module also
 * loads in browsers. It returns the same module objects an import would, which
 * keeps their functions replaceable (tests stub `os.tmpdir`, for example).
 */
function nodeProcess(): NodeJS.Process | undefined {
  const candidate = globalThis.process;
  return typeof candidate?.getBuiltinModule === 'function' ? candidate : undefined;
}

/**
 * Prepares a file path for the native binding.
 *
 * The WebAssembly build reaches the file system through WASI, which resolves a
 * relative path against `/` instead of the working directory, so on Node.js the
 * path is made absolute first. That gives relative paths the same meaning in
 * both builds. The native addon gets the path unchanged.
 */
function bindingPath(path: string): string {
  if (__napiBindingTarget === 'native') {
    return path;
  }
  return nodeProcess()?.getBuiltinModule('node:path').resolve(path) ?? path;
}

function isErrnoException(error: unknown): error is NodeJS.ErrnoException {
  return error instanceof Error && 'code' in error;
}

/** Color value - can be a named color or hex code */
export type Color =
  | 'black'
  | 'blue'
  | 'brown'
  | 'cyan'
  | 'gray'
  | 'green'
  | 'lime'
  | 'magenta'
  | 'navy'
  | 'orange'
  | 'pink'
  | 'purple'
  | 'red'
  | 'silver'
  | 'white'
  | 'yellow'
  | `#${string}`; // Hex colors like #FF0000

/** Horizontal alignment options */
export type HorizontalAlign =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerAcross'
  | 'distributed';

/** Vertical alignment options */
export type VerticalAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/** Border style options */
export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'double'
  | 'dotted'
  | 'dashed'
  | 'dashDot'
  | 'dashDotDot'
  | 'slantDashDot'
  | 'mediumDashed'
  | 'mediumDashDot'
  | 'mediumDashDotDot';

/** Underline style options */
export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/** Fill pattern options */
export type FillPattern =
  | 'solid'
  | 'mediumGray'
  | 'darkGray'
  | 'lightGray'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis'
  | 'gray125'
  | 'gray0625';

/** Font theme scheme */
export type FontScheme = 'body' | 'headings' | 'none';

/** Data validation rule type */
export type ValidationRule =
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

/** Data validation error style */
export type ValidationErrorStyle = 'stop' | 'warning' | 'information';

/**
 * A value that {@link Worksheet.write}, {@link Worksheet.writeRow} and
 * {@link Worksheet.writeRows} write to a cell. A `Date` is written as its UTC
 * date and time; `null` and `undefined` are no value.
 */
export type CellValue = string | number | boolean | Date | null | undefined;

/**
 * When Excel recalculates formulas: `automatic` (the default), `manual` (only on
 * request, e.g. F9), or `partial` (automatic except for data tables).
 */
export type CalculationMode = 'automatic' | 'manual' | 'partial';

/**
 * An Excel warning (the green triangle in a cell's corner) that
 * {@link Worksheet.ignoreError} turns off:
 * - `numberStoredAsText`: a number stored as text
 * - `formulaError`: a formula that evaluates to an error, such as `#DIV/0!`
 * - `formulaDiffers`: a formula that differs from the formulas around it
 * - `formulaRefersToEmptyCells`: a formula that refers to empty cells
 * - `formulaOmitsCells`: a formula that omits cells next to its range
 * - `dataValidationError`: a value that breaks the cell's data validation
 * - `unlockedCellsWithFormula`: an unlocked cell that contains a formula
 * - `inconsistentColumnFormula`: a table cell whose formula differs from its
 *   column's formula
 *
 * Excel's two-digit text year warning is not available: rust_xlsxwriter
 * 0.102.0 writes its rule with an attribute name Excel doesn't define.
 */
export type IgnoreErrorType =
  | 'numberStoredAsText'
  | 'formulaError'
  | 'formulaDiffers'
  | 'formulaRefersToEmptyCells'
  | 'formulaOmitsCells'
  | 'dataValidationError'
  | 'unlockedCellsWithFormula'
  | 'inconsistentColumnFormula';

/**
 * Format class for cell styling with fluent API
 * All methods return `this` for method chaining
 *
 * @example
 * const format = new Format()
 *   .setBold()
 *   .setFontColor('red')
 *   .setAlign('center');
 */
export class Format {
  /** Internal native binding */
  _native: RuscFormat;

  /** Create a new format */
  constructor() {
    this._native = new RuscFormat();
  }

  // Font styling methods (return this for chaining)
  /** Set bold text */
  setBold(): this {
    this._native.setBold();
    return this;
  }
  /** Set italic text */
  setItalic(): this {
    this._native.setItalic();
    return this;
  }
  /** Set font size in points */
  setFontSize(size: number): this {
    this._native.setFontSize(size);
    return this;
  }
  /** Set font name (e.g., 'Arial', 'Calibri', 'Times New Roman') */
  setFontName(fontName: string): this {
    this._native.setFontName(fontName);
    return this;
  }
  /** Set whether the font belongs to the body, headings, or no theme scheme */
  setFontScheme(scheme: FontScheme): this {
    this._native.setFontScheme(scheme);
    return this;
  }
  /** Set font color using named color or hex code */
  setFontColor(color: Color): this {
    this._native.setFontColor(color);
    return this;
  }
  /** Set text underline style */
  setUnderline(underline: UnderlineStyle): this {
    this._native.setUnderline(underline);
    return this;
  }
  /** Set strikethrough text */
  setStrikethrough(): this {
    this._native.setStrikethrough();
    return this;
  }

  // Fill and background methods
  /** Set background color using named color or hex code */
  setBackgroundColor(color: Color): this {
    this._native.setBackgroundColor(color);
    return this;
  }
  /** Set foreground color (for patterns) using named color or hex code */
  setForegroundColor(color: Color): this {
    this._native.setForegroundColor(color);
    return this;
  }
  /** Set fill pattern */
  setPattern(pattern: FillPattern): this {
    this._native.setPattern(pattern);
    return this;
  }

  // Alignment and text formatting methods
  /** Set horizontal alignment */
  setAlign(align: HorizontalAlign): this {
    this._native.setAlign(align);
    return this;
  }
  /** Set vertical alignment */
  setVerticalAlign(align: VerticalAlign): this {
    this._native.setVerticalAlign(align);
    return this;
  }
  /** Enable text wrapping in the cell */
  setTextWrap(): this {
    this._native.setTextWrap();
    return this;
  }
  /** Set text indentation level (0-15) */
  setIndent(indent: number): this {
    this._native.setIndent(indent);
    return this;
  }
  /** Set text rotation in degrees (-90 to 90, or 270 for vertical text) */
  setRotation(rotation: number): this {
    this._native.setRotation(rotation);
    return this;
  }
  /** Shrink text to fit cell */
  setShrink(): this {
    this._native.setShrink();
    return this;
  }

  // Border methods
  /** Set border style for all borders */
  setBorder(border: BorderStyle): this {
    this._native.setBorder(border);
    return this;
  }
  /** Set top border style */
  setBorderTop(border: BorderStyle): this {
    this._native.setBorderTop(border);
    return this;
  }
  /** Set bottom border style */
  setBorderBottom(border: BorderStyle): this {
    this._native.setBorderBottom(border);
    return this;
  }
  /** Set left border style */
  setBorderLeft(border: BorderStyle): this {
    this._native.setBorderLeft(border);
    return this;
  }
  /** Set right border style */
  setBorderRight(border: BorderStyle): this {
    this._native.setBorderRight(border);
    return this;
  }
  /** Set color for all borders */
  setBorderColor(color: Color): this {
    this._native.setBorderColor(color);
    return this;
  }
  /** Set top border color */
  setBorderTopColor(color: Color): this {
    this._native.setBorderTopColor(color);
    return this;
  }
  /** Set bottom border color */
  setBorderBottomColor(color: Color): this {
    this._native.setBorderBottomColor(color);
    return this;
  }
  /** Set left border color */
  setBorderLeftColor(color: Color): this {
    this._native.setBorderLeftColor(color);
    return this;
  }
  /** Set right border color */
  setBorderRightColor(color: Color): this {
    this._native.setBorderRightColor(color);
    return this;
  }

  // Number format
  /**
   * Set number format using Excel format strings
   * @example
   * format.setNumFormat('#,##0.00');      // Thousands separator
   * format.setNumFormat('$#,##0.00');     // Currency
   * format.setNumFormat('0.00%');         // Percentage
   * format.setNumFormat('yyyy-mm-dd');    // Date
   */
  setNumFormat(format: string): this {
    this._native.setNumFormat(format);
    return this;
  }

  // Cell protection
  /** Unlock cell (for use with worksheet protection) */
  setUnlocked(): this {
    this._native.setUnlocked();
    return this;
  }
  /** Lock cell (cell will be protected when worksheet is protected) */
  setLocked(): this {
    this._native.setLocked();
    return this;
  }
  /** Hide formulas in cell (for use with worksheet protection) */
  setHidden(): this {
    this._native.setHidden();
    return this;
  }

  /**
   * Set font script (superscript or subscript)
   * @param script - 'superscript' for X² or 'subscript' for H₂O
   * @example
   * format.setFontScript('superscript'); // X²
   * format.setFontScript('subscript');   // H₂O
   */
  setFontScript(script: 'superscript' | 'subscript'): this {
    this._native.setFontScript(script);
    return this;
  }

  /**
   * Set diagonal border style
   * @param border - Border style
   * @example
   * format.setBorderDiagonal('thin');
   */
  setBorderDiagonal(border: BorderStyle): this {
    this._native.setBorderDiagonal(border);
    return this;
  }

  /**
   * Set diagonal border color
   * @param color - Color for diagonal border
   */
  setBorderDiagonalColor(color: Color): this {
    this._native.setBorderDiagonalColor(color);
    return this;
  }

  /**
   * Set diagonal border type
   * @param type - 'up' for / border, 'down' for \ border, 'both' for X
   * @example
   * format.setBorderDiagonalType('up');   // /
   * format.setBorderDiagonalType('down'); // \
   * format.setBorderDiagonalType('both'); // X
   */
  setBorderDiagonalType(type: 'up' | 'down' | 'both'): this {
    this._native.setBorderDiagonalType(type);
    return this;
  }

  /**
   * Set reading direction for RTL languages
   * @param direction - 0=context, 1=left-to-right, 2=right-to-left
   * @example
   * format.setReadingDirection(2); // Right-to-left (Arabic, Hebrew)
   */
  setReadingDirection(direction: 0 | 1 | 2): this {
    this._native.setReadingDirection(direction);
    return this;
  }

  /**
   * Force string interpretation with leading apostrophe
   * Useful for numbers that should be treated as text
   * @example
   * format.setQuotePrefix(); // '00123 will display as 00123
   */
  setQuotePrefix(): this {
    this._native.setQuotePrefix();
    return this;
  }
}

/**
 * Note class for Excel cell notes (comments) with fluent API
 * All methods return `this` for method chaining
 *
 * Notes are post-it style comments that appear when hovering over a cell.
 * They show as a small red triangle in the cell corner.
 *
 * @example
 * const note = new Note('Important: Review this data')
 *   .setAuthor('John Doe')
 *   .setWidth(200)
 *   .setBackgroundColor('#FFFFE1');
 */
export class Note {
  /** Internal native binding */
  _native: RuscNote;

  /**
   * Create a new note with text
   * @param text - The text content of the note
   */
  constructor(text: string) {
    this._native = new RuscNote(text);
  }

  /**
   * Set the author name for the note (max 52 characters)
   * @param name - Author name
   */
  setAuthor(name: string): this {
    this._native.setAuthor(name);
    return this;
  }

  /**
   * Set the width of the note in pixels
   * @param width - Width in pixels
   */
  setWidth(width: number): this {
    this._native.setWidth(width);
    return this;
  }

  /**
   * Set the height of the note in pixels
   * @param height - Height in pixels
   */
  setHeight(height: number): this {
    this._native.setHeight(height);
    return this;
  }

  /**
   * Set the background color of the note
   * @param color - Color name or hex code
   */
  setBackgroundColor(color: Color): this {
    this._native.setBackgroundColor(color);
    return this;
  }

  /**
   * Make the note visible by default (normally hidden until hover)
   * @param visible - True to show by default, false to show on hover
   */
  setVisible(visible: boolean): this {
    this._native.setVisible(visible);
    return this;
  }

  /**
   * Set alt text for accessibility
   * @param altText - Accessibility description
   */
  setAltText(altText: string): this {
    this._native.setAltText(altText);
    return this;
  }
}

/**
 * Formula class for Excel formulas with fluent API
 *
 * @example
 * const formula = new Formula('=SUM(A1:A10)');
 * const formulaWithResult = new Formula('=2+2').setResult('4');
 */
export class Formula {
  /**
   * Internal formula data passed to the native writer. A `Formula` is plain
   * JavaScript data, so creating one per row allocates no Rust memory.
   */
  readonly _data: RuscFormulaData;

  /**
   * Create a new formula
   * @param formula - Excel formula string (e.g., '=SUM(A1:A10)')
   */
  constructor(formula: string) {
    this._data = { formula };
  }

  /**
   * Set the result value for applications that don't auto-calculate formulas
   * @param result - The calculated result as a string
   */
  setResult(result: string): this {
    this._data.result = result;
    return this;
  }
}

/** The data the native formula writers take for a `Formula` or a formula string. */
function formulaData(formula: Formula | string): RuscFormulaData {
  return typeof formula === 'string' ? { formula } : formula._data;
}

/**
 * DataValidation class for restricting cell input with fluent API
 *
 * Allows you to create dropdown lists, number ranges, date validation, etc.
 *
 * @example
 * // Dropdown list
 * const validation = new DataValidation()
 *   .allowListStrings(['Option 1', 'Option 2', 'Option 3'])
 *   .setInputTitle('Select an option')
 *   .setInputMessage('Choose from the dropdown');
 *
 * @example
 * // Number range
 * const validation = new DataValidation()
 *   .allowWholeNumberBetween(1, 100)
 *   .setErrorTitle('Invalid number')
 *   .setErrorMessage('Please enter a number between 1 and 100');
 */
export class DataValidation {
  /** Internal native binding */
  _native: RuscDataValidation;

  /** Create a new data validation */
  constructor() {
    this._native = new RuscDataValidation();
  }

  /**
   * Allow whole number validation with a rule
   * @param rule - Validation rule type
   * @param value - Comparison value
   */
  allowWholeNumber(rule: ValidationRule, value: number): this {
    this._native.allowWholeNumber(rule, value);
    return this;
  }

  /**
   * Allow whole number validation between min and max
   * @param min - Minimum value (inclusive)
   * @param max - Maximum value (inclusive)
   */
  allowWholeNumberBetween(min: number, max: number): this {
    this._native.allowWholeNumberBetween(min, max);
    return this;
  }

  /**
   * Allow decimal number validation with a rule
   * @param rule - Validation rule type
   * @param value - Comparison value
   */
  allowDecimalNumber(rule: ValidationRule, value: number): this {
    this._native.allowDecimalNumber(rule, value);
    return this;
  }

  /**
   * Allow decimal number validation between min and max
   * @param min - Minimum value (inclusive)
   * @param max - Maximum value (inclusive)
   */
  allowDecimalNumberBetween(min: number, max: number): this {
    this._native.allowDecimalNumberBetween(min, max);
    return this;
  }

  /**
   * Allow list validation with an array of strings (creates dropdown)
   * @param list - Array of allowed values (max 255 chars total)
   */
  allowListStrings(list: string[]): this {
    this._native.allowListStrings(list);
    return this;
  }

  /**
   * Allow list validation from a cell range formula
   * @param formula - Cell range formula (e.g., '=$A$1:$A$10')
   */
  allowListFormula(formula: string): this {
    this._native.allowListFormula(formula);
    return this;
  }

  /**
   * Allow custom formula validation
   * @param formula - Excel formula that returns true/false
   */
  allowCustom(formula: string): this {
    this._native.allowCustom(formula);
    return this;
  }

  /**
   * Allow any value (mainly used for displaying input messages)
   */
  allowAnyValue(): this {
    this._native.allowAnyValue();
    return this;
  }

  /**
   * Set whether to ignore blank cells
   * @param enable - True to allow blank cells (default: true)
   */
  ignoreBlank(enable: boolean): this {
    this._native.ignoreBlank(enable);
    return this;
  }

  /**
   * Set whether to show dropdown arrow for list validations
   * @param enable - True to show dropdown (default: true)
   */
  showDropdown(enable: boolean): this {
    this._native.showDropdown(enable);
    return this;
  }

  /**
   * Set whether to show input message when cell is selected
   * @param enable - True to show message (default: true)
   */
  showInputMessage(enable: boolean): this {
    this._native.showInputMessage(enable);
    return this;
  }

  /**
   * Set whether to show error message on invalid input
   * @param enable - True to show message (default: true)
   */
  showErrorMessage(enable: boolean): this {
    this._native.showErrorMessage(enable);
    return this;
  }

  /**
   * Set the input message title (shown when cell is selected)
   * @param title - Title text (max 32 characters)
   */
  setInputTitle(title: string): this {
    this._native.setInputTitle(title);
    return this;
  }

  /**
   * Set the input message text (shown when cell is selected)
   * @param message - Message text (max 255 characters)
   */
  setInputMessage(message: string): this {
    this._native.setInputMessage(message);
    return this;
  }

  /**
   * Set the error message title (shown on invalid input)
   * @param title - Title text (max 32 characters)
   */
  setErrorTitle(title: string): this {
    this._native.setErrorTitle(title);
    return this;
  }

  /**
   * Set the error message text (shown on invalid input)
   * @param message - Message text (max 255 characters)
   */
  setErrorMessage(message: string): this {
    this._native.setErrorMessage(message);
    return this;
  }

  /**
   * Set the error dialog style
   * @param style - Error style: 'stop' (red X), 'warning' (yellow !), or 'information' (blue i)
   */
  setErrorStyle(style: ValidationErrorStyle): this {
    this._native.setErrorStyle(style);
    return this;
  }
}

/**
 * Protection Options for worksheet protection
 * Configure granular permissions for protected worksheets
 *
 * @example
 * const options = new ProtectionOptions()
 *   .setInsertRows(true)
 *   .setInsertColumns(true)
 *   .setFormatCells(false);
 * worksheet.protectWithOptions(options);
 */
export class ProtectionOptions {
  /** Internal native binding */
  _native: RuscProtectionOptions;

  /** Create a new ProtectionOptions with default settings */
  constructor() {
    this._native = new RuscProtectionOptions();
  }

  /**
   * Allow or prevent selecting locked cells (default: true)
   * @param allow - true to allow, false to prevent
   */
  setSelectLockedCells(allow: boolean): this {
    this._native.setSelectLockedCells(allow);
    return this;
  }

  /**
   * Allow or prevent selecting unlocked cells (default: true)
   * @param allow - true to allow, false to prevent
   */
  setSelectUnlockedCells(allow: boolean): this {
    this._native.setSelectUnlockedCells(allow);
    return this;
  }

  /**
   * Allow or prevent formatting cells (default: false)
   * @param allow - true to allow, false to prevent
   */
  setFormatCells(allow: boolean): this {
    this._native.setFormatCells(allow);
    return this;
  }

  /**
   * Allow or prevent formatting columns (default: false)
   * @param allow - true to allow, false to prevent
   */
  setFormatColumns(allow: boolean): this {
    this._native.setFormatColumns(allow);
    return this;
  }

  /**
   * Allow or prevent formatting rows (default: false)
   * @param allow - true to allow, false to prevent
   */
  setFormatRows(allow: boolean): this {
    this._native.setFormatRows(allow);
    return this;
  }

  /**
   * Allow or prevent inserting columns (default: false)
   * @param allow - true to allow, false to prevent
   */
  setInsertColumns(allow: boolean): this {
    this._native.setInsertColumns(allow);
    return this;
  }

  /**
   * Allow or prevent inserting rows (default: false)
   * @param allow - true to allow, false to prevent
   */
  setInsertRows(allow: boolean): this {
    this._native.setInsertRows(allow);
    return this;
  }

  /**
   * Allow or prevent inserting hyperlinks (default: false)
   * @param allow - true to allow, false to prevent
   */
  setInsertLinks(allow: boolean): this {
    this._native.setInsertLinks(allow);
    return this;
  }

  /**
   * Allow or prevent deleting columns (default: false)
   * @param allow - true to allow, false to prevent
   */
  setDeleteColumns(allow: boolean): this {
    this._native.setDeleteColumns(allow);
    return this;
  }

  /**
   * Allow or prevent deleting rows (default: false)
   * @param allow - true to allow, false to prevent
   */
  setDeleteRows(allow: boolean): this {
    this._native.setDeleteRows(allow);
    return this;
  }

  /**
   * Allow or prevent sorting data (default: false)
   * @param allow - true to allow, false to prevent
   */
  setSort(allow: boolean): this {
    this._native.setSort(allow);
    return this;
  }

  /**
   * Allow or prevent using autofilters (default: false)
   * @param allow - true to allow, false to prevent
   */
  setUseAutofilter(allow: boolean): this {
    this._native.setUseAutofilter(allow);
    return this;
  }

  /**
   * Allow or prevent using pivot tables (default: false)
   * @param allow - true to allow, false to prevent
   */
  setUsePivotTables(allow: boolean): this {
    this._native.setUsePivotTables(allow);
    return this;
  }

  /**
   * Allow or prevent editing scenarios (default: false)
   * @param allow - true to allow, false to prevent
   */
  setEditScenarios(allow: boolean): this {
    this._native.setEditScenarios(allow);
    return this;
  }

  /**
   * Allow or prevent editing objects like images (default: false)
   * @param allow - true to allow, false to prevent
   */
  setEditObjects(allow: boolean): this {
    this._native.setEditObjects(allow);
    return this;
  }

  /**
   * Allow or prevent editing contents in chartsheets (default: true)
   * @param allow - true to allow, false to prevent
   */
  setContents(allow: boolean): this {
    this._native.setContents(allow);
    return this;
  }
}

/**
 * Image for inserting images into worksheets
 * Supports PNG, JPG, GIF, and BMP formats
 *
 * @example
 * const image = new Image('logo.png');
 * image.setWidth(200).setHeight(100);
 * worksheet.insertImage(0, 0, image);
 */
export class Image {
  /** Internal native binding */
  _native: RuscImage;

  /**
   * Create a new Image from a file path
   * @param path - Path to the image file (PNG, JPG, GIF, or BMP)
   * @throws Error if file cannot be read or format is unsupported
   */
  constructor(path: string) {
    this._native = RuscImage.new(bindingPath(path));
  }

  /**
   * Create an Image from a buffer
   * @param buffer - Buffer or Uint8Array containing image data
   * @returns Image instance
   * @example
   * const buffer = fs.readFileSync('logo.png');
   * const image = Image.fromBuffer(buffer);
   *
   * @example
   * // In a browser
   * const response = await fetch('/logo.png');
   * const image = Image.fromBuffer(new Uint8Array(await response.arrayBuffer()));
   */
  static fromBuffer(buffer: Uint8Array): Image {
    // Bypasses the path-based constructor; the native image comes from the buffer.
    const image: Image = Object.create(Image.prototype);
    image._native = RuscImage.newFromBuffer(buffer);
    return image;
  }

  /**
   * Set the width of the image in pixels
   * @param width - Width in pixels
   * @returns this for method chaining
   */
  setWidth(width: number): this {
    this._native.setWidth(width);
    return this;
  }

  /**
   * Set the height of the image in pixels
   * @param height - Height in pixels
   * @returns this for method chaining
   */
  setHeight(height: number): this {
    this._native.setHeight(height);
    return this;
  }

  /**
   * Scale the width of the image as a percentage
   * @param scale - Scale percentage (e.g., 1.5 for 150%)
   * @returns this for method chaining
   */
  setScaleWidth(scale: number): this {
    this._native.setScaleWidth(scale);
    return this;
  }

  /**
   * Scale the height of the image as a percentage
   * @param scale - Scale percentage (e.g., 0.5 for 50%)
   * @returns this for method chaining
   */
  setScaleHeight(scale: number): this {
    this._native.setScaleHeight(scale);
    return this;
  }

  /**
   * Scale the image to a specific size
   * @param width - Target width
   * @param height - Target height
   * @param keepAspectRatio - Whether to maintain aspect ratio (default: true)
   * @returns this for method chaining
   * @example
   * image.setScaleToSize(300, 200, true);
   */
  setScaleToSize(width: number, height: number, keepAspectRatio: boolean = true): this {
    this._native.setScaleToSize(width, height, keepAspectRatio);
    return this;
  }

  /**
   * Set alt text for the image (for accessibility)
   * @param altText - Descriptive text for the image
   * @returns this for method chaining
   */
  setAltText(altText: string): this {
    this._native.setAltText(altText);
    return this;
  }

  /**
   * Mark the image as decorative (no alt text required)
   * @param enable - true to mark as decorative
   * @returns this for method chaining
   */
  setDecorative(enable: boolean): this {
    this._native.setDecorative(enable);
    return this;
  }

  /**
   * Set how the image moves and sizes with cells
   * @param option - 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'
   * @returns this for method chaining
   * @example
   * image.setObjectMovement('moveOnly'); // Image moves but doesn't resize with cells
   */
  setObjectMovement(option: 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'): this {
    this._native.setObjectMovement(option);
    return this;
  }

  /**
   * Add a URL link to the image
   * @param url - URL to link to
   * @returns this for method chaining
   * @example
   * image.setUrl('https://example.com');
   */
  setUrl(url: string): this {
    this._native.setUrl(url);
    return this;
  }
}

/**
 * Rich Text for multi-format text in a single cell
 *
 * @example
 * const richText = new RichText()
 *   .addSegment('This is ', defaultFormat)
 *   .addSegment('bold', boldFormat)
 *   .addSegment(' and this is ', defaultFormat)
 *   .addSegment('italic', italicFormat);
 * worksheet.writeRichString(0, 0, richText);
 */
export class RichText {
  /**
   * Internal `[text, format]` segments passed to the native writer. A
   * `RichText` is plain JavaScript data, so creating one per row allocates no
   * Rust memory.
   */
  readonly _segments: [text: string, format: RuscFormat | undefined][] = [];

  /**
   * Add a text segment with formatting. The format is read when the rich text
   * is written, as the format of any other write is.
   * @param text - The text content
   * @param format - Optional format to apply to this segment; without one the
   *   segment uses the default cell font
   * @returns this for method chaining
   * @example
   * richText.addSegment('Hello ', normalFormat)
   *         .addSegment('World', boldFormat);
   */
  addSegment(text: string, format?: Format): this {
    this._segments.push([text, format?._native]);
    return this;
  }
}

/**
 * Worksheet class for Excel sheet manipulation
 *
 * @example
 * const worksheet = workbook.addWorksheet('Data');
 * worksheet.write(0, 0, 'Name');
 * worksheet.write(0, 1, 'Age');
 * worksheet.setColumnWidth(0, 20);
 */
export class Worksheet {
  /** Internal native binding */
  _native: RuscWorksheet;

  /** @internal */
  constructor(nativeWorksheet: RuscWorksheet) {
    this._native = nativeWorksheet;
  }

  /**
   * Write a value to a cell (auto-detects type)
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Value to write (string, number, boolean or Date); `null` or
   *   `undefined` writes nothing, or a blank cell when a format is given. A Date
   *   is written as its UTC date and time and, unless the format has a number
   *   format of its own, shown with Excel's built-in date or date-time format.
   * @param format - Optional format to apply
   * @throws If the value is of another type, such as an object, or an invalid
   *   Date or one outside Excel's 1900-9999 range
   * @example
   * worksheet.write(0, 0, new Date('2024-01-15')); // shown as a date, e.g. 1/15/2024
   */
  write(row: number, col: number, value: CellValue, format?: Format): void {
    if (format) {
      return this._native.write(row, col, value, format._native);
    }

    switch (typeof value) {
      case 'string':
        return this._native.writeString(row, col, value);
      case 'number':
        return this._native.writeNumber(row, col, value);
      case 'boolean':
        return this._native.writeBoolean(row, col, value);
      default:
        return this._native.write(row, col, value);
    }
  }

  /**
   * Write an array of values as a row starting at the given column.
   * Batches all writes into a single native call — much faster than calling write() in a loop.
   * Each value is written as {@link Worksheet.write} writes it: `null`, `undefined` and
   * array holes leave the cell empty, or write a blank cell when a format is given.
   *
   * @param row - Row number (0-indexed)
   * @param startCol - Starting column number (0-indexed)
   * @param values - Array of values to write
   * @param format - Optional format to apply to all cells in the row
   * @throws If a value is of another type, such as an object, or an invalid
   *   Date; the values before it are already written
   * @example
   * worksheet.writeRow(0, 0, ['Name', 'Age', 'Active'], headerFormat);
   * worksheet.writeRow(1, 0, ['Alice', 30, true]);
   */
  writeRow(row: number, startCol: number, values: CellValue[], format?: Format): void {
    return this._native.writeRow(row, startCol, values, format ? format._native : undefined);
  }

  /**
   * Write a 2D array of values starting at the given position.
   * Batches all writes — much faster than nested write() loops.
   * Each value is written as {@link Worksheet.write} writes it: `null`, `undefined` and
   * array holes leave the cell empty, or write a blank cell when a format is given.
   *
   * @param startRow - Starting row number (0-indexed)
   * @param startCol - Starting column number (0-indexed)
   * @param rows - 2D array of values to write
   * @param format - Optional format to apply to all cells
   * @throws If a row is not an array or a value is of another type, such as an
   *   object, or an invalid Date; the values before it are already written
   * @example
   * worksheet.writeRows(0, 0, [
   *   ['Name', 'Age', 'Active'],
   *   ['Alice', 30, true],
   *   ['Bob', 25, false],
   * ]);
   */
  writeRows(startRow: number, startCol: number, rows: CellValue[][], format?: Format): void {
    return this._native.writeRows(startRow, startCol, rows, format ? format._native : undefined);
  }

  /**
   * Write a string to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - String value to write
   */
  writeString(row: number, col: number, value: string): void {
    return this._native.writeString(row, col, value);
  }

  /**
   * Write a number to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Number value to write
   */
  writeNumber(row: number, col: number, value: number): void {
    return this._native.writeNumber(row, col, value);
  }

  /**
   * Write a boolean to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Boolean value to write
   */
  writeBoolean(row: number, col: number, value: boolean): void {
    return this._native.writeBoolean(row, col, value);
  }

  /**
   * Write a blank cell with formatting
   * Useful for creating styled empty cells or clearing cells while keeping formatting
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param format - Format to apply to the blank cell
   * @example
   * const grayFormat = new Format().setBackgroundColor('gray');
   * worksheet.writeBlank(0, 0, grayFormat); // Gray empty cell
   */
  writeBlank(row: number, col: number, format: Format): void {
    return this._native.writeBlank(row, col, format._native);
  }

  /**
   * Write a datetime from a Unix timestamp (seconds since 1970-01-01 UTC)
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param timestamp - Unix timestamp in seconds
   * @param format - Optional format; unless it has a number format of its own,
   *   the value is shown with Excel's built-in date, date-time or time format
   * @example
   * const dateFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm');
   * worksheet.writeDatetimeFromTimestamp(0, 0, 1706184600, dateFormat);
   */
  writeDatetimeFromTimestamp(row: number, col: number, timestamp: number, format?: Format): void {
    return this._native.writeDatetimeFromTimestamp(
      row,
      col,
      timestamp,
      format ? format._native : undefined,
    );
  }

  /**
   * Write a datetime from year, month, day components
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param year - Year (1900-9999)
   * @param month - Month (1-12)
   * @param day - Day (1-31)
   * @param format - Optional format; unless it has a number format of its own,
   *   the value is shown with Excel's built-in date format
   * @example
   * const dateFormat = new Format().setNumFormat('yyyy-mm-dd');
   * worksheet.writeDatetimeFromYmd(0, 0, 2023, 1, 25, dateFormat);
   */
  writeDatetimeFromYmd(
    row: number,
    col: number,
    year: number,
    month: number,
    day: number,
    format?: Format,
  ): void {
    return this._native.writeDatetimeFromYmd(
      row,
      col,
      year,
      month,
      day,
      format ? format._native : undefined,
    );
  }

  /**
   * Write a datetime from year, month, day, hour, minute, second components
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param year - Year (1900-9999)
   * @param month - Month (1-12)
   * @param day - Day (1-31)
   * @param hour - Hour (0-23)
   * @param min - Minute (0-59)
   * @param sec - Second (0-59.999...)
   * @param format - Optional format; unless it has a number format of its own,
   *   the value is shown with Excel's built-in date, date-time or time format
   * @example
   * const datetimeFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm:ss');
   * worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 1, 25, 12, 30, 45.5, datetimeFormat);
   */
  writeDatetimeFromYmdHms(
    row: number,
    col: number,
    year: number,
    month: number,
    day: number,
    hour: number,
    min: number,
    sec: number,
    format?: Format,
  ): void {
    return this._native.writeDatetimeFromYmdHms(
      row,
      col,
      year,
      month,
      day,
      hour,
      min,
      sec,
      format ? format._native : undefined,
    );
  }

  /**
   * Write a datetime from an ISO 8601 string
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param datetimeStr - ISO 8601 date/time string (e.g., "2023-01-25T12:30:00" or "2023-01-25")
   * @param format - Optional format; unless it has a number format of its own,
   *   the value is shown with Excel's built-in date, date-time or time format
   * @example
   * const datetimeFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm');
   * worksheet.writeDatetimeFromString(0, 0, '2023-01-25T12:30:00', datetimeFormat);
   */
  writeDatetimeFromString(row: number, col: number, datetimeStr: string, format?: Format): void {
    return this._native.writeDatetimeFromString(
      row,
      col,
      datetimeStr,
      format ? format._native : undefined,
    );
  }

  /**
   * Write a datetime (convenience method that auto-detects input type)
   * Accepts: Date object (written as its UTC date and time), Unix timestamp in seconds
   * (number), ISO string, or object with date components
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param datetime - Date/time value in various formats
   * @param format - Optional format; unless it has a number format of its own,
   *   the value is shown with Excel's built-in date, date-time or time format
   * @example
   * // Using JavaScript Date
   * const dateFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm');
   * worksheet.writeDatetime(0, 0, new Date(), dateFormat);
   *
   * @example
   * // Using timestamp
   * worksheet.writeDatetime(1, 0, 1706184600, dateFormat);
   *
   * @example
   * // Using ISO string
   * worksheet.writeDatetime(2, 0, '2023-01-25T12:30:00', dateFormat);
   *
   * @example
   * // Using object
   * worksheet.writeDatetime(3, 0, { year: 2023, month: 1, day: 25, hour: 12, min: 30 }, dateFormat);
   */
  writeDatetime(
    row: number,
    col: number,
    datetime:
      | Date
      | number
      | string
      | { year: number; month: number; day: number; hour?: number; min?: number; sec?: number },
    format?: Format,
  ): void {
    // Convenience method that handles multiple input types
    if (datetime instanceof Date) {
      // A Date is a cell value of its own, written with its milliseconds
      return this._native.write(row, col, datetime, format ? format._native : undefined);
    } else if (typeof datetime === 'number') {
      // Unix timestamp
      return this.writeDatetimeFromTimestamp(row, col, datetime, format);
    } else if (typeof datetime === 'string') {
      // ISO string
      return this.writeDatetimeFromString(row, col, datetime, format);
    } else if (typeof datetime === 'object') {
      // Object with year, month, day, etc.
      if (datetime.hour !== undefined || datetime.min !== undefined || datetime.sec !== undefined) {
        return this.writeDatetimeFromYmdHms(
          row,
          col,
          datetime.year,
          datetime.month,
          datetime.day,
          datetime.hour || 0,
          datetime.min || 0,
          datetime.sec || 0,
          format,
        );
      } else {
        return this.writeDatetimeFromYmd(
          row,
          col,
          datetime.year,
          datetime.month,
          datetime.day,
          format,
        );
      }
    } else {
      throw new Error(
        'Invalid datetime input: must be Date, number (timestamp), string (ISO), or object',
      );
    }
  }

  /**
   * Write a value with formatting to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Value to write (string, number, boolean or Date); `null` or
   *   `undefined` writes a blank cell with the format
   * @param format - Format to apply; a Date shown with Excel's built-in date
   *   format unless the format has a number format of its own
   * @throws If the value is of another type, such as an object, or an invalid Date
   */
  writeWithFormat(row: number, col: number, value: CellValue, format: Format): void {
    return this._native.writeWithFormat(row, col, value, format._native);
  }

  /**
   * Set the width of a column
   * @param col - Column number (0-indexed)
   * @param width - Width in Excel units (approximately character width)
   */
  setColumnWidth(col: number, width: number): void {
    return this._native.setColumnWidth(col, width);
  }

  /**
   * Set the height of a row
   * @param row - Row number (0-indexed)
   * @param height - Height in points
   */
  setRowHeight(row: number, height: number): void {
    return this._native.setRowHeight(row, height);
  }

  /**
   * Set the default height for all worksheet rows without per-row metadata
   * @param height - Height in points
   */
  setDefaultRowHeight(height: number): void {
    return this._native.setDefaultRowHeight(height);
  }

  /**
   * Set the width of a column in pixels
   * @param col - Column number (0-indexed)
   * @param width - Width in pixels
   */
  setColumnWidthPixels(col: number, width: number): void {
    return this._native.setColumnWidthPixels(col, width);
  }

  /**
   * Set the width of a range of columns
   * @param firstCol - First column (0-indexed)
   * @param lastCol - Last column (0-indexed), inclusive
   * @param width - Width in Excel units (approximately character width)
   */
  setColumnRangeWidth(firstCol: number, lastCol: number, width: number): void {
    return this._native.setColumnRangeWidth(firstCol, lastCol, width);
  }

  /**
   * Set the width of a range of columns in pixels
   * @param firstCol - First column (0-indexed)
   * @param lastCol - Last column (0-indexed), inclusive
   * @param width - Width in pixels
   */
  setColumnRangeWidthPixels(firstCol: number, lastCol: number, width: number): void {
    return this._native.setColumnRangeWidthPixels(firstCol, lastCol, width);
  }

  /**
   * Set the width of a column in pixels in auto-fit mode. As with a column that
   * Excel auto-fitted, Excel widens it when the user enters a number with more
   * digits; a width from {@link Worksheet.setColumnWidthPixels} stays fixed.
   * @param col - Column number (0-indexed)
   * @param width - Width in pixels
   */
  setColumnAutofitWidth(col: number, width: number): void {
    return this._native.setColumnAutofitWidth(col, width);
  }

  /**
   * Set the default format of a column, used by its cells that have no format
   * of their own
   * @param col - Column number (0-indexed)
   * @param format - Format for the column
   * @example
   * worksheet.setColumnFormat(2, new Format().setNumFormat('0.00%'));
   */
  setColumnFormat(col: number, format: Format): void {
    return this._native.setColumnFormat(col, format._native);
  }

  /**
   * Set the default format of a range of columns
   * @param firstCol - First column (0-indexed)
   * @param lastCol - Last column (0-indexed), inclusive
   * @param format - Format for the columns
   */
  setColumnRangeFormat(firstCol: number, lastCol: number, format: Format): void {
    return this._native.setColumnRangeFormat(firstCol, lastCol, format._native);
  }

  /**
   * Hide a column
   * @param col - Column number (0-indexed)
   */
  setColumnHidden(col: number): void {
    return this._native.setColumnHidden(col);
  }

  /**
   * Hide a range of columns
   * @param firstCol - First column (0-indexed)
   * @param lastCol - Last column (0-indexed), inclusive
   */
  setColumnRangeHidden(firstCol: number, lastCol: number): void {
    return this._native.setColumnRangeHidden(firstCol, lastCol);
  }

  /**
   * Set the height of a row in pixels
   * @param row - Row number (0-indexed)
   * @param height - Height in pixels
   */
  setRowHeightPixels(row: number, height: number): void {
    return this._native.setRowHeightPixels(row, height);
  }

  /**
   * Set the default height of all worksheet rows in pixels
   * @param height - Height in pixels
   */
  setDefaultRowHeightPixels(height: number): void {
    return this._native.setDefaultRowHeightPixels(height);
  }

  /**
   * Set the default format of a row, used by its cells that have no format of
   * their own
   * @param row - Row number (0-indexed)
   * @param format - Format for the row
   * @example
   * worksheet.setRowFormat(0, new Format().setBold());
   */
  setRowFormat(row: number, format: Format): void {
    return this._native.setRowFormat(row, format._native);
  }

  /**
   * Hide a row
   * @param row - Row number (0-indexed)
   */
  setRowHidden(row: number): void {
    return this._native.setRowHidden(row);
  }

  /**
   * Unhide a row hidden with {@link Worksheet.setRowHidden}
   * @param row - Row number (0-indexed)
   */
  setRowUnhidden(row: number): void {
    return this._native.setRowUnhidden(row);
  }

  /**
   * Hide every row that has no data, format or row settings. It uses a single
   * worksheet setting, so it is much smaller and quicker than hiding each row.
   * @param enable - True to hide the unused rows
   */
  hideUnusedRows(enable: boolean): void {
    return this._native.hideUnusedRows(enable);
  }

  /**
   * Group a range of rows into an outline level. A range inside an existing
   * group adds a nested level; groups nest up to 7 deep.
   * @param firstRow - First row (0-indexed)
   * @param lastRow - Last row (0-indexed), inclusive
   * @example
   * worksheet.groupRows(1, 10);
   * worksheet.groupRows(2, 5); // nested level
   */
  groupRows(firstRow: number, lastRow: number): void {
    return this._native.groupRows(firstRow, lastRow);
  }

  /**
   * Group a range of rows into a collapsed outline level, with the rows hidden
   * @param firstRow - First row (0-indexed)
   * @param lastRow - Last row (0-indexed), inclusive
   */
  groupRowsCollapsed(firstRow: number, lastRow: number): void {
    return this._native.groupRowsCollapsed(firstRow, lastRow);
  }

  /**
   * Group a range of columns into an outline level. A range inside an existing
   * group adds a nested level; groups nest up to 7 deep.
   * @param firstCol - First column (0-indexed)
   * @param lastCol - Last column (0-indexed), inclusive
   */
  groupColumns(firstCol: number, lastCol: number): void {
    return this._native.groupColumns(firstCol, lastCol);
  }

  /**
   * Group a range of columns into a collapsed outline level, with the columns
   * hidden
   * @param firstCol - First column (0-indexed)
   * @param lastCol - Last column (0-indexed), inclusive
   */
  groupColumnsCollapsed(firstCol: number, lastCol: number): void {
    return this._native.groupColumnsCollapsed(firstCol, lastCol);
  }

  /**
   * Show the row outline expand/collapse symbols above each group instead of
   * below it
   * @param enable - True to show the symbols above
   */
  groupSymbolsAbove(enable: boolean): void {
    return this._native.groupSymbolsAbove(enable);
  }

  /**
   * Show the column outline expand/collapse symbols to the left of each group
   * instead of to the right
   * @param enable - True to show the symbols to the left
   */
  groupSymbolsToLeft(enable: boolean): void {
    return this._native.groupSymbolsToLeft(enable);
  }

  /**
   * Set the format of a cell without changing its data; an empty cell becomes a
   * formatted blank cell. In constant memory mode only cells of the row being
   * written can be formatted: rows already written are left unchanged.
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param format - Format for the cell
   */
  setCellFormat(row: number, col: number, format: Format): void {
    return this._native.setCellFormat(row, col, format._native);
  }

  /**
   * Set the format of every cell in a range without changing their data. In
   * constant memory mode only cells of the row being written can be formatted.
   * @param firstRow - First row (0-indexed)
   * @param firstCol - First column (0-indexed)
   * @param lastRow - Last row (0-indexed), inclusive
   * @param lastCol - Last column (0-indexed), inclusive
   * @param format - Format for the cells
   * @example
   * worksheet.writeRows(0, 0, rows);
   * worksheet.setRangeFormat(0, 0, 0, 3, headerFormat);
   */
  setRangeFormat(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    format: Format,
  ): void {
    return this._native.setRangeFormat(firstRow, firstCol, lastRow, lastCol, format._native);
  }

  /**
   * Set the format of every cell in a range without changing their data, and
   * draw the border of `borderFormat` around the outside of the range. In
   * constant memory mode only cells of the row being written can be formatted.
   * @param firstRow - First row (0-indexed)
   * @param firstCol - First column (0-indexed)
   * @param lastRow - Last row (0-indexed), inclusive
   * @param lastCol - Last column (0-indexed), inclusive
   * @param cellFormat - Format for the cells
   * @param borderFormat - Format whose border is drawn around the range
   * @example
   * worksheet.setRangeFormatWithBorder(0, 0, 9, 3, new Format(), new Format().setBorder('thick'));
   */
  setRangeFormatWithBorder(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    cellFormat: Format,
    borderFormat: Format,
  ): void {
    return this._native.setRangeFormatWithBorder(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      cellFormat._native,
      borderFormat._native,
    );
  }

  /**
   * Remove the data and the format of a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   */
  clearCell(row: number, col: number): void {
    return this._native.clearCell(row, col);
  }

  /**
   * Remove the format of a cell without changing its data
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   */
  clearCellFormat(row: number, col: number): void {
    return this._native.clearCellFormat(row, col);
  }

  /**
   * Turn off an Excel warning, such as "number stored as text", for a cell.
   * Excel allows one rule per cell or range, so a cell or range that already
   * has one can't get another.
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param errorType - The warning to turn off
   * @throws If the type is unknown, or the cell already has a rule
   * @example
   * worksheet.write(0, 0, '0123');
   * worksheet.ignoreError(0, 0, 'numberStoredAsText');
   */
  ignoreError(row: number, col: number, errorType: IgnoreErrorType): void {
    return this._native.ignoreError(row, col, errorType);
  }

  /**
   * Turn off an Excel warning, such as "number stored as text", for a range of
   * cells. Excel allows one rule per cell or range, so a range that already has
   * one can't get another.
   * @param firstRow - First row (0-indexed)
   * @param firstCol - First column (0-indexed)
   * @param lastRow - Last row (0-indexed), inclusive
   * @param lastCol - Last column (0-indexed), inclusive
   * @param errorType - The warning to turn off
   * @throws If the type is unknown, the range is inverted, or the range already
   *   has a rule
   * @example
   * worksheet.ignoreErrorRange(1, 0, 1000, 0, 'numberStoredAsText');
   */
  ignoreErrorRange(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    errorType: IgnoreErrorType,
  ): void {
    return this._native.ignoreErrorRange(firstRow, firstCol, lastRow, lastCol, errorType);
  }

  /**
   * Set the exclusive upper row bound inspected by autofit
   * @param maxRow - Rows at or above this zero-indexed bound are skipped
   */
  setAutofitMaxRow(maxRow: number): void {
    return this._native.setAutofitMaxRow(maxRow);
  }

  /**
   * Limit the maximum autofitted column width
   * @param maxWidth - Maximum width in pixels
   */
  setAutofitMaxWidth(maxWidth: number): void {
    return this._native.setAutofitMaxWidth(maxWidth);
  }

  /** Autofit worksheet columns based on cell contents */
  autofit(): void {
    return this._native.autofit();
  }

  /**
   * Set the worksheet name
   * @param name - Worksheet name
   */
  setName(name: string): void {
    return this._native.setName(name);
  }

  /**
   * Merge a range of cells and write a value to it
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param value - Value to write to merged cell
   * @param format - Optional format to apply
   */
  mergeRange(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    value: string,
    format?: Format,
  ): void {
    return this._native.mergeRange(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      value,
      format ? format._native : undefined,
    );
  }

  /**
   * Merge a range of cells with formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param value - Value to write to merged cell
   * @param format - Format to apply
   */
  mergeRangeWithFormat(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    value: string,
    format: Format,
  ): void {
    return this._native.mergeRangeWithFormat(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      value,
      format._native,
    );
  }

  /**
   * Insert a note (post-it style comment) to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param note - Note object to insert
   */
  insertNote(row: number, col: number, note: Note): void {
    return this._native.insertNote(row, col, note._native);
  }

  /**
   * Set the default author name for all notes in the worksheet
   * @param name - Author name (max 52 characters)
   */
  setDefaultNoteAuthor(name: string): void {
    return this._native.setDefaultNoteAuthor(name);
  }

  /**
   * Write a formula to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param formula - Formula object or formula string (e.g., '=SUM(A1:A10)')
   * @param format - Optional format to apply
   * @example
   * worksheet.writeFormula(0, 2, '=A1+B1');
   * worksheet.writeFormula(1, 2, new Formula('=A2+B2').setResult('42'));
   */
  writeFormula(row: number, col: number, formula: Formula | string, format?: Format): void {
    return this._native.writeFormula(
      row,
      col,
      formulaData(formula),
      format ? format._native : undefined,
    );
  }

  /**
   * Write a dynamic array formula to a cell (Excel 365+)
   * Dynamic arrays automatically spill into adjacent cells
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param formula - Formula object or formula string with a dynamic array
   *   function (FILTER, SORT, UNIQUE, etc.)
   * @param format - Optional format to apply
   */
  writeDynamicFormula(row: number, col: number, formula: Formula | string, format?: Format): void {
    return this._native.writeDynamicFormula(
      row,
      col,
      formulaData(formula),
      format ? format._native : undefined,
    );
  }

  /**
   * Write an array formula to a range of cells
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param formula - Formula object or formula string
   * @param format - Optional format to apply
   */
  writeArrayFormula(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    formula: Formula | string,
    format?: Format,
  ): void {
    return this._native.writeArrayFormula(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      formulaData(formula),
      format ? format._native : undefined,
    );
  }

  /**
   * Add data validation to a cell or range
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param validation - DataValidation object with rules
   */
  addDataValidation(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    validation: DataValidation,
  ): void {
    return this._native.addDataValidation(firstRow, firstCol, lastRow, lastCol, validation._native);
  }

  /**
   * Write a URL/hyperlink to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param url - URL string (http://, https://, file://, or internal:Sheet1!A1)
   * @param format - Optional format to apply
   */
  writeUrl(row: number, col: number, url: string, format?: Format): void {
    return this._native.writeUrl(row, col, url, format ? format._native : undefined);
  }

  /**
   * Write a URL/hyperlink to a cell with custom display text
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param url - URL string
   * @param text - Display text for the link
   * @param format - Optional format to apply
   */
  writeUrlWithText(row: number, col: number, url: string, text: string, format?: Format): void {
    return this._native.writeUrlWithText(row, col, url, text, format ? format._native : undefined);
  }

  /**
   * Add autofilter dropdown buttons to a range
   * @param firstRow - First row of range (0-indexed, typically header row)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   */
  autofilter(firstRow: number, firstCol: number, lastRow: number, lastCol: number): void {
    return this._native.autofilter(firstRow, firstCol, lastRow, lastCol);
  }

  /**
   * Freeze panes at the specified row and column
   * @param row - Row to freeze above (0 = no horizontal freeze)
   * @param col - Column to freeze left of (0 = no vertical freeze)
   * @example
   * worksheet.setFreezePanes(1, 0); // Freeze top row
   * worksheet.setFreezePanes(0, 1); // Freeze first column
   * worksheet.setFreezePanes(1, 1); // Freeze both top row and first column
   */
  setFreezePanes(row: number, col: number): void {
    return this._native.setFreezePanes(row, col);
  }

  /**
   * Add 2-color scale conditional formatting (gradient between two colors)
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param minColor - Color for minimum value
   * @param maxColor - Color for maximum value
   */
  addConditionalFormat2ColorScale(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    minColor: Color,
    maxColor: Color,
  ): void {
    return this._native.addConditionalFormat2ColorScale(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      minColor,
      maxColor,
    );
  }

  /**
   * Add 3-color scale conditional formatting (gradient across three colors)
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param minColor - Color for minimum value
   * @param midColor - Color for middle value
   * @param maxColor - Color for maximum value
   */
  addConditionalFormat3ColorScale(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    minColor: Color,
    midColor: Color,
    maxColor: Color,
  ): void {
    return this._native.addConditionalFormat3ColorScale(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      minColor,
      midColor,
      maxColor,
    );
  }

  /**
   * Add data bar conditional formatting (horizontal bar visualization in cells)
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param barColor - Color for the data bars
   */
  addConditionalFormatDataBar(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    barColor: Color,
  ): void {
    return this._native.addConditionalFormatDataBar(firstRow, firstCol, lastRow, lastCol, barColor);
  }

  /**
   * Add cell-based conditional formatting (highlight cells based on value)
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param rule - Comparison rule ('greaterThan', 'lessThan', 'equal', 'notEqual', 'greaterThanOrEqual', 'lessThanOrEqual')
   * @param value - Comparison value
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatCell(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    rule:
      | 'greaterThan'
      | 'lessThan'
      | 'equal'
      | 'notEqual'
      | 'greaterThanOrEqual'
      | 'lessThanOrEqual',
    value: number,
    format: Format,
  ): void {
    return this._native.addConditionalFormatCell(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      rule,
      value,
      format._native,
    );
  }

  /**
   * Add text-based conditional formatting (highlight cells based on text content)
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param rule - Text matching rule ('contains', 'notContains', 'beginsWith', 'endsWith')
   * @param text - Text to match
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatText(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    rule: 'contains' | 'notContains' | 'beginsWith' | 'endsWith',
    text: string,
    format: Format,
  ): void {
    return this._native.addConditionalFormatText(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      rule,
      text,
      format._native,
    );
  }

  /**
   * Add average/standard deviation conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param rule - Average rule ('aboveAverage', 'belowAverage', 'equalOrAboveAverage', 'equalOrBelowAverage')
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatAverage(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    rule: 'aboveAverage' | 'belowAverage' | 'equalOrAboveAverage' | 'equalOrBelowAverage',
    format: Format,
  ): void {
    return this._native.addConditionalFormatAverage(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      rule,
      format._native,
    );
  }

  /**
   * Add top/bottom N values conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param rule - Top/bottom rule ('top', 'bottom', 'topPercent', 'bottomPercent')
   * @param value - Number of items or percentage
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatTop(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    rule: 'top' | 'bottom' | 'topPercent' | 'bottomPercent',
    value: number,
    format: Format,
  ): void {
    return this._native.addConditionalFormatTop(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      rule,
      value,
      format._native,
    );
  }

  /**
   * Add duplicate/unique values conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param invert - If true, highlight unique values instead of duplicates
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatDuplicate(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    invert: boolean,
    format: Format,
  ): void {
    return this._native.addConditionalFormatDuplicate(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      invert,
      format._native,
    );
  }

  /**
   * Add blank/non-blank cells conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param invert - If true, highlight non-blank cells instead of blank
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatBlank(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    invert: boolean,
    format: Format,
  ): void {
    return this._native.addConditionalFormatBlank(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      invert,
      format._native,
    );
  }

  /**
   * Add error/non-error cells conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param invert - If true, highlight non-error cells instead of errors
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatError(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    invert: boolean,
    format: Format,
  ): void {
    return this._native.addConditionalFormatError(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      invert,
      format._native,
    );
  }

  /**
   * Add date-based conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param rule - Date rule ('yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek', 'lastMonth', 'thisMonth', 'nextMonth')
   * @param format - Format to apply to matching cells
   */
  addConditionalFormatDate(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    rule:
      | 'yesterday'
      | 'today'
      | 'tomorrow'
      | 'last7Days'
      | 'lastWeek'
      | 'thisWeek'
      | 'nextWeek'
      | 'lastMonth'
      | 'thisMonth'
      | 'nextMonth',
    format: Format,
  ): void {
    return this._native.addConditionalFormatDate(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      rule,
      format._native,
    );
  }

  /**
   * Add formula-based conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param formula - Excel formula (e.g., '=MOD(ROW(),2)=0' for even rows)
   * @param format - Format to apply when formula returns true
   */
  addConditionalFormatFormula(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    formula: string,
    format: Format,
  ): void {
    return this._native.addConditionalFormatFormula(
      firstRow,
      firstCol,
      lastRow,
      lastCol,
      formula,
      format._native,
    );
  }

  /**
   * Add icon set conditional formatting
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param iconType - Icon set type (e.g., 'threeArrows', 'threeTrafficLights', 'fiveStars', etc.)
   */
  addConditionalFormatIconSet(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    iconType:
      | 'threeArrows'
      | 'threeArrowsGray'
      | 'threeFlags'
      | 'threeTrafficLights'
      | 'threeTrafficLightsWithRim'
      | 'threeSigns'
      | 'threeSymbols'
      | 'threeSymbolsCircled'
      | 'threeStars'
      | 'threeTriangles'
      | 'fourArrows'
      | 'fourArrowsGray'
      | 'fourRedToBlack'
      | 'fourHistograms'
      | 'fourTrafficLights'
      | 'fiveArrows'
      | 'fiveArrowsGray'
      | 'fiveHistograms'
      | 'fiveQuadrants'
      | 'fiveBoxes',
  ): void {
    return this._native.addConditionalFormatIconSet(firstRow, firstCol, lastRow, lastCol, iconType);
  }

  /**
   * Set page margins for printing (in inches)
   * @param left - Left margin
   * @param right - Right margin
   * @param top - Top margin
   * @param bottom - Bottom margin
   * @param header - Header margin
   * @param footer - Footer margin
   */
  setMargins(
    left: number,
    right: number,
    top: number,
    bottom: number,
    header: number,
    footer: number,
  ): void {
    return this._native.setMargins(left, right, top, bottom, header, footer);
  }

  /**
   * Set page orientation to portrait
   */
  setPortrait(): void {
    return this._native.setPortrait();
  }

  /**
   * Set page orientation to landscape
   */
  setLandscape(): void {
    return this._native.setLandscape();
  }

  /**
   * Set paper size for printing
   * @param paperSize - Paper size code (1=Letter, 5=Legal, 9=A4, 11=A5, etc.)
   */
  setPaperSize(paperSize: number): void {
    return this._native.setPaperSize(paperSize);
  }

  /**
   * Set header text for printing
   * @param header - Header text (use &L, &C, &R for left/center/right, &P for page number, &D for date)
   * @example
   * worksheet.setHeader('&CReport Title'); // Centered header
   * worksheet.setHeader('&LCompany&CTitle&RPage &P'); // Left, center, and right
   */
  setHeader(header: string): void {
    return this._native.setHeader(header);
  }

  /**
   * Set footer text for printing
   * @param footer - Footer text (use &L, &C, &R for left/center/right, &P for page number, &D for date)
   * @example
   * worksheet.setFooter('&CPage &P of &N'); // Centered page number
   * worksheet.setFooter('&L&D&RPage &P'); // Date left, page right
   */
  setFooter(footer: string): void {
    return this._native.setFooter(footer);
  }

  /**
   * Set print area (only this range will be printed)
   * @param firstRow - First row of print area (0-indexed)
   * @param firstCol - First column of print area (0-indexed)
   * @param lastRow - Last row of print area (0-indexed)
   * @param lastCol - Last column of print area (0-indexed)
   */
  setPrintArea(firstRow: number, firstCol: number, lastRow: number, lastCol: number): void {
    return this._native.setPrintArea(firstRow, firstCol, lastRow, lastCol);
  }

  /**
   * Show or hide gridlines when printing
   * @param enable - True to show gridlines, false to hide
   */
  setPrintGridlines(enable: boolean): void {
    return this._native.setPrintGridlines(enable);
  }

  /**
   * Center worksheet horizontally on the page when printing
   * @param enable - True to center horizontally
   */
  setPrintCenterHorizontally(enable: boolean): void {
    return this._native.setPrintCenterHorizontally(enable);
  }

  /**
   * Center worksheet vertically on the page when printing
   * @param enable - True to center vertically
   */
  setPrintCenterVertically(enable: boolean): void {
    return this._native.setPrintCenterVertically(enable);
  }

  /**
   * Fit worksheet to N pages wide by M pages tall when printing
   * @param width - Number of pages wide (0 = no limit)
   * @param height - Number of pages tall (0 = no limit)
   * @example
   * worksheet.setPrintFitToPages(1, 0); // Fit to 1 page wide, unlimited height
   */
  setPrintFitToPages(width: number, height: number): void {
    return this._native.setPrintFitToPages(width, height);
  }

  /**
   * Set print scale as a percentage
   * @param scale - Scale percentage (10-400)
   * @example
   * worksheet.setPrintScale(75); // Print at 75% scale
   */
  setPrintScale(scale: number): void {
    return this._native.setPrintScale(scale);
  }

  /**
   * Set the first page number for printing
   * @param pageNumber - Starting page number
   */
  setPrintFirstPageNumber(pageNumber: number): void {
    return this._native.setPrintFirstPageNumber(pageNumber);
  }

  /**
   * Enable/disable black and white printing
   * @param enable - True to print in black and white
   */
  setPrintBlackAndWhite(enable: boolean): void {
    return this._native.setPrintBlackAndWhite(enable);
  }

  /**
   * Enable/disable draft quality printing
   * @param enable - True for draft quality
   */
  setPrintDraft(enable: boolean): void {
    return this._native.setPrintDraft(enable);
  }

  /**
   * Enable/disable printing of row and column headings
   * @param enable - True to print headings
   */
  setPrintHeadings(enable: boolean): void {
    return this._native.setPrintHeadings(enable);
  }

  /**
   * Repeat rows at the top of each printed page
   * @param firstRow - First row to repeat (0-indexed)
   * @param lastRow - Last row to repeat (0-indexed)
   * @example
   * worksheet.setRepeatRows(0, 0); // Repeat first row on each page
   * worksheet.setRepeatRows(0, 2); // Repeat first 3 rows on each page
   */
  setRepeatRows(firstRow: number, lastRow: number): void {
    return this._native.setRepeatRows(firstRow, lastRow);
  }

  /**
   * Repeat columns at the left of each printed page
   * @param firstCol - First column to repeat (0-indexed)
   * @param lastCol - Last column to repeat (0-indexed)
   * @example
   * worksheet.setRepeatColumns(0, 0); // Repeat first column on each page
   * worksheet.setRepeatColumns(0, 1); // Repeat first 2 columns on each page
   */
  setRepeatColumns(firstCol: number, lastCol: number): void {
    return this._native.setRepeatColumns(firstCol, lastCol);
  }

  /**
   * Scale header and footer with document
   * @param enable - True to scale with document
   */
  setHeaderFooterScaleWithDoc(enable: boolean): void {
    return this._native.setHeaderFooterScaleWithDoc(enable);
  }

  /**
   * Align header and footer with page margins
   * @param enable - True to align with page margins
   */
  setHeaderFooterAlignWithPage(enable: boolean): void {
    return this._native.setHeaderFooterAlignWithPage(enable);
  }

  /**
   * Insert a chart at the specified cell position
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param chart - Chart object to insert
   * @example
   * const chart = new Chart('column');
   * chart.addSeries({ values: 'Sheet1!$B$2:$B$7', categories: 'Sheet1!$A$2:$A$7' });
   * worksheet.insertChart(1, 3, chart);
   */
  insertChart(row: number, col: number, chart: Chart): void {
    return this._native.insertChart(row, col, chart._native);
  }

  /**
   * Insert a chart at the specified cell position with pixel offsets
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param chart - Chart object to insert
   * @param xOffset - Horizontal offset in pixels
   * @param yOffset - Vertical offset in pixels
   * @example
   * const chart = new Chart('line');
   * chart.addSeries({ values: 'Sheet1!$B$2:$B$7' });
   * worksheet.insertChartWithOffset(1, 3, chart, 25, 10);
   */
  insertChartWithOffset(
    row: number,
    col: number,
    chart: Chart,
    xOffset: number,
    yOffset: number,
  ): void {
    return this._native.insertChartWithOffset(row, col, chart._native, xOffset, yOffset);
  }

  /**
   * Insert an image at the specified cell position
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param image - Image instance
   * @example
   * const image = new Image('logo.png');
   * worksheet.insertImage(0, 0, image);
   */
  insertImage(row: number, col: number, image: Image): void {
    return this._native.insertImage(row, col, image._native);
  }

  /**
   * Insert an image at the specified cell position with pixel offsets
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param image - Image instance
   * @param xOffset - Horizontal offset in pixels
   * @param yOffset - Vertical offset in pixels
   * @example
   * const image = new Image('logo.png');
   * worksheet.insertImageWithOffset(1, 2, image, 10, 20);
   */
  insertImageWithOffset(
    row: number,
    col: number,
    image: Image,
    xOffset: number,
    yOffset: number,
  ): void {
    return this._native.insertImageWithOffset(row, col, image._native, xOffset, yOffset);
  }

  /**
   * Insert an image that fits to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param image - Image instance
   * @param keepAspectRatio - Whether to maintain aspect ratio (default: true)
   * @example
   * const image = new Image('photo.jpg');
   * worksheet.insertImageFitToCell(0, 0, image, true);
   */
  insertImageFitToCell(
    row: number,
    col: number,
    image: Image,
    keepAspectRatio: boolean = true,
  ): void {
    return this._native.insertImageFitToCell(row, col, image._native, keepAspectRatio);
  }

  /**
   * Insert an image scaled proportionally and centered within a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param image - Image instance
   */
  insertImageFitToCellCentered(row: number, col: number, image: Image): void {
    return this._native.insertImageFitToCellCentered(row, col, image._native);
  }

  /**
   * Add a table to the worksheet
   * @param firstRow - First row of the table (0-indexed)
   * @param firstCol - First column of the table (0-indexed)
   * @param lastRow - Last row of the table (0-indexed)
   * @param lastCol - Last column of the table (0-indexed)
   * @param table - Table instance
   * @example
   * const table = new Table();
   * table.setName('Sales').setStyle(9);
   * worksheet.addTable(0, 0, 10, 4, table);
   */
  addTable(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    table: Table,
  ): void {
    return this._native.addTable(firstRow, firstCol, lastRow, lastCol, table._native);
  }

  /**
   * Add a sparkline to a single cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param sparkline - Sparkline instance
   * @example
   * const sparkline = new Sparkline();
   * sparkline.setRange('A1:E1').setType('column');
   * worksheet.addSparkline(0, 5, sparkline);
   */
  addSparkline(row: number, col: number, sparkline: Sparkline): void {
    return this._native.addSparkline(row, col, sparkline._native);
  }

  /**
   * Add a sparkline group to a range of cells
   * @param firstRow - First row of the sparkline range (0-indexed)
   * @param firstCol - First column of the sparkline range (0-indexed)
   * @param lastRow - Last row of the sparkline range (0-indexed)
   * @param lastCol - Last column of the sparkline range (0-indexed)
   * @param sparkline - Sparkline instance
   * @example
   * const sparkline = new Sparkline();
   * sparkline.setRange('A1:E3').setType('line');
   * worksheet.addSparklineGroup(0, 5, 2, 5, sparkline);
   */
  addSparklineGroup(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    sparkline: Sparkline,
  ): void {
    return this._native.addSparklineGroup(firstRow, firstCol, lastRow, lastCol, sparkline._native);
  }

  /**
   * Set the worksheet tab color
   * @param color - Color name or hex code
   * @example
   * worksheet.setTabColor('red');
   * worksheet.setTabColor('#FF0000');
   */
  setTabColor(color: Color): void {
    return this._native.setTabColor(color);
  }

  /**
   * Hide or show the worksheet
   * @param enable - True to hide worksheet, false to show
   * @example
   * worksheet.setHidden(true); // Hide worksheet
   */
  setHidden(enable: boolean): void {
    return this._native.setHidden(enable);
  }

  /**
   * Set the worksheet zoom level
   * @param zoom - Zoom percentage (10-400)
   * @example
   * worksheet.setZoom(150); // 150% zoom
   */
  setZoom(zoom: number): void {
    return this._native.setZoom(zoom);
  }

  /**
   * Hide zero values in worksheet cells, the equivalent of turning off the Excel
   * option "Show a zero in cells that have zero value"
   * @param enable - True to hide zero values, false to show them (the default)
   * @example
   * worksheet.setHiddenZeroes(true);
   */
  setHiddenZeroes(enable: boolean): void {
    return this._native.setHiddenZeroes(enable);
  }

  /**
   * Set the default cell selection when opening the worksheet
   * @param firstRow - First row of selection (0-indexed)
   * @param firstCol - First column of selection (0-indexed)
   * @param lastRow - Last row of selection (0-indexed)
   * @param lastCol - Last column of selection (0-indexed)
   * @example
   * worksheet.setSelection(0, 0, 5, 5); // Select A1:F6
   */
  setSelection(firstRow: number, firstCol: number, lastRow: number, lastCol: number): void {
    return this._native.setSelection(firstRow, firstCol, lastRow, lastCol);
  }

  /**
   * Set which cell appears in the top-left corner when opening the worksheet
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @example
   * worksheet.setTopLeftCell(10, 0); // Start view at row 11
   */
  setTopLeftCell(row: number, col: number): void {
    return this._native.setTopLeftCell(row, col);
  }

  /**
   * Protect the worksheet from editing
   * Cells must be unlocked via format.setUnlocked() to remain editable
   * @example
   * worksheet.protect(); // Protect all cells
   */
  protect(): void {
    return this._native.protect();
  }

  /**
   * Protect the worksheet with a password
   * @param password - Password to protect worksheet
   * @example
   * worksheet.protectWithPassword('secret123');
   */
  protectWithPassword(password: string): void {
    return this._native.protectWithPassword(password);
  }

  /**
   * Protect the worksheet with granular protection options
   * @param options - ProtectionOptions instance with desired settings
   * @example
   * const options = new ProtectionOptions()
   *   .setInsertRows(true)
   *   .setInsertColumns(true);
   * worksheet.protectWithOptions(options);
   */
  protectWithOptions(options: ProtectionOptions): void {
    return this._native.protectWithOptions(options._native);
  }

  /**
   * Protect the worksheet with a password and granular protection options
   * @param password - Password to protect worksheet
   * @param options - ProtectionOptions instance with desired settings
   * @example
   * const options = new ProtectionOptions()
   *   .setInsertRows(true);
   * worksheet.protectWithPasswordAndOptions('secret123', options);
   */
  protectWithPasswordAndOptions(password: string, options: ProtectionOptions): void {
    return this._native.protectWithPasswordAndOptions(password, options._native);
  }

  /**
   * Write a rich text string (multi-format text) to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param richText - RichText instance with formatted segments
   * @example
   * const richText = new RichText()
   *   .addSegment('This is ', normal)
   *   .addSegment('bold', bold);
   * worksheet.writeRichString(0, 0, richText);
   */
  writeRichString(row: number, col: number, richText: RichText): void {
    return this._native.writeRichString(row, col, richText._segments);
  }

  /**
   * Write a rich text string with an additional cell format
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param richText - RichText instance with formatted segments
   * @param cellFormat - Format to apply to the entire cell
   * @example
   * const richText = new RichText()
   *   .addSegment('Hello ', normal)
   *   .addSegment('World', bold);
   * const cellFormat = new Format().setAlign('center');
   * worksheet.writeRichStringWithFormat(0, 0, richText, cellFormat);
   */
  writeRichStringWithFormat(
    row: number,
    col: number,
    richText: RichText,
    cellFormat: Format,
  ): void {
    return this._native.writeRichStringWithFormat(row, col, richText._segments, cellFormat._native);
  }

  /**
   * Insert a button at the specified cell position
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param button - Button instance
   * @example
   * const button = new Button();
   * button.setCaption('Run Macro').setMacro('MyMacro');
   * worksheet.insertButton(0, 0, button);
   */
  insertButton(row: number, col: number, button: Button): void {
    return this._native.insertButton(row, col, button._native);
  }

  /**
   * Insert a button at the specified cell position with pixel offsets
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param button - Button instance
   * @param xOffset - Horizontal offset in pixels
   * @param yOffset - Vertical offset in pixels
   * @example
   * const button = new Button();
   * button.setCaption('Click Me');
   * worksheet.insertButtonWithOffset(0, 0, button, 10, 5);
   */
  insertButtonWithOffset(
    row: number,
    col: number,
    button: Button,
    xOffset: number,
    yOffset: number,
  ): void {
    return this._native.insertButtonWithOffset(row, col, button._native, xOffset, yOffset);
  }

  /**
   * Insert a shape at the specified cell position
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param shape - Shape instance
   * @example
   * const shape = Shape.textbox();
   * shape.setText('Important: Read carefully');
   * worksheet.insertShape(0, 0, shape);
   */
  insertShape(row: number, col: number, shape: Shape): void {
    return this._native.insertShape(row, col, shape._native);
  }

  /**
   * Insert a shape at the specified cell position with pixel offsets
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param shape - Shape instance
   * @param xOffset - Horizontal offset in pixels
   * @param yOffset - Vertical offset in pixels
   * @example
   * const shape = Shape.textbox();
   * shape.setText('Note');
   * worksheet.insertShapeWithOffset(1, 1, shape, 15, 10);
   */
  insertShapeWithOffset(
    row: number,
    col: number,
    shape: Shape,
    xOffset: number,
    yOffset: number,
  ): void {
    return this._native.insertShapeWithOffset(row, col, shape._native, xOffset, yOffset);
  }
}

/** Chart type options */
export type ChartType =
  | 'area'
  | 'areaStacked'
  | 'area_stacked'
  | 'areaPercentStacked'
  | 'area_percent_stacked'
  | 'bar'
  | 'barStacked'
  | 'bar_stacked'
  | 'barPercentStacked'
  | 'bar_percent_stacked'
  | 'column'
  | 'columnStacked'
  | 'column_stacked'
  | 'columnPercentStacked'
  | 'column_percent_stacked'
  | 'doughnut'
  | 'line'
  | 'lineStacked'
  | 'line_stacked'
  | 'linePercentStacked'
  | 'line_percent_stacked'
  | 'pie'
  | 'radar'
  | 'radarWithMarkers'
  | 'radar_with_markers'
  | 'radarFilled'
  | 'radar_filled'
  | 'scatter'
  | 'scatterStraight'
  | 'scatter_straight'
  | 'scatterStraightWithMarkers'
  | 'scatter_straight_with_markers'
  | 'scatterSmooth'
  | 'scatter_smooth'
  | 'scatterSmoothWithMarkers'
  | 'scatter_smooth_with_markers'
  | 'stock';

/** Chart series configuration options */
export interface ChartSeriesOptions {
  /** Data values range (e.g., 'Sheet1!$B$2:$B$7') */
  values: string;
  /** Category labels range (optional, e.g., 'Sheet1!$A$2:$A$7') */
  categories?: string;
  /** Series name (optional, e.g., 'Sales Q1' or 'Sheet1!$B$1') */
  name?: string;
}

/**
 * Chart class for creating Excel charts
 *
 * @example
 * const chart = new Chart('column');
 * chart.addSeries({
 *   categories: 'Sheet1!$A$2:$A$7',
 *   values: 'Sheet1!$B$2:$B$7',
 *   name: 'Q1 Sales'
 * });
 * chart.setTitle('Sales Report');
 * chart.setXAxisName('Month');
 * chart.setYAxisName('Revenue');
 * chart.setStyle(11);
 * worksheet.insertChart(0, 2, chart);
 */
export class Chart {
  /** Internal native binding */
  _native: RuscChart;

  /**
   * Create a new chart
   * @param chartType - Type of chart to create
   * @example
   * const chart = new Chart('column');
   * const pieChart = new Chart('pie');
   * const lineChart = new Chart('line');
   */
  constructor(chartType: ChartType) {
    this._native = new RuscChart(chartType);
  }

  /**
   * Add a data series to the chart
   * @param options - Series configuration options
   * @example
   * chart.addSeries({
   *   categories: 'Sheet1!$A$2:$A$7',
   *   values: 'Sheet1!$B$2:$B$7',
   *   name: 'Sales'
   * });
   */
  addSeries(options: ChartSeriesOptions): void {
    const { categories, values, name } = options;
    return this._native.addSeries(categories || null, values, name || null);
  }

  /**
   * Set the chart title
   * @param title - Chart title text
   * @example
   * chart.setTitle('Q1 Sales Results');
   */
  setTitle(title: string): this {
    this._native.setTitle(title);
    return this;
  }

  /**
   * Set the X-axis label
   * @param name - X-axis label text
   * @example
   * chart.setXAxisName('Month');
   */
  setXAxisName(name: string): this {
    this._native.setXAxisName(name);
    return this;
  }

  /**
   * Set the Y-axis label
   * @param name - Y-axis label text
   * @example
   * chart.setYAxisName('Revenue ($)');
   */
  setYAxisName(name: string): this {
    this._native.setYAxisName(name);
    return this;
  }

  /**
   * Set the chart style (1-48)
   * @param style - Excel chart style number (1-48)
   * @example
   * chart.setStyle(11); // Apply Excel style 11
   */
  setStyle(style: number): this {
    this._native.setStyle(style);
    return this;
  }

  /**
   * Hide the chart legend
   * @example
   * chart.setLegendHidden();
   */
  setLegendHidden(): this {
    this._native.setLegendHidden();
    return this;
  }
}

/**
 * Table class for creating Excel tables
 *
 * @example
 * const table = new Table();
 * table.setName('SalesData')
 *      .setStyle(9)
 *      .setTotalRow(true);
 * worksheet.addTable(0, 0, 10, 4, table);
 */
export class Table {
  /** Internal native binding */
  _native: RuscTable;

  /** Create a new table */
  constructor() {
    this._native = RuscTable.new();
  }

  /**
   * Set whether the table has a header row
   * @param show - Show header row (default: true)
   */
  setHeaderRow(show: boolean): this {
    this._native.setHeaderRow(show);
    return this;
  }

  /**
   * Set whether the table has a total row
   * @param show - Show total row (default: false)
   */
  setTotalRow(show: boolean): this {
    this._native.setTotalRow(show);
    return this;
  }

  /**
   * Set whether the table has banded rows
   * @param show - Show banded rows (default: true)
   */
  setBandedRows(show: boolean): this {
    this._native.setBandedRows(show);
    return this;
  }

  /**
   * Set whether the table has banded columns
   * @param show - Show banded columns (default: false)
   */
  setBandedColumns(show: boolean): this {
    this._native.setBandedColumns(show);
    return this;
  }

  /**
   * Set whether the first column is formatted
   * @param show - Format first column (default: false)
   */
  setFirstColumn(show: boolean): this {
    this._native.setFirstColumn(show);
    return this;
  }

  /**
   * Set whether the last column is formatted
   * @param show - Format last column (default: false)
   */
  setLastColumn(show: boolean): this {
    this._native.setLastColumn(show);
    return this;
  }

  /**
   * Set the name of the table
   * @param name - Table name (must be unique)
   */
  setName(name: string): this {
    this._native.setName(name);
    return this;
  }

  /** Set alternative text for screen readers */
  setAltText(altText: string): this {
    this._native.setAltText(altText);
    return this;
  }

  /** Set the alternative text title for screen readers */
  setAltTextTitle(title: string): this {
    this._native.setAltTextTitle(title);
    return this;
  }

  /**
   * Set the style of the table
   * @param styleId - Table style ID (0: no style, 1-21: light, 22-49: medium, 50-60: dark)
   * @throws If the ID is outside 0-60
   * @example
   * table.setStyle(30); // Medium style 9, Excel's default
   */
  setStyle(styleId: number): this {
    this._native.setStyle(styleId);
    return this;
  }

  /**
   * Set columns for the table
   * @param columns - Array of TableColumn objects
   */
  setColumns(columns: TableColumn[]): this {
    const nativeColumns = columns.map((col) => col._native);
    this._native.setColumns(nativeColumns);
    return this;
  }
}

/**
 * TableColumn class for configuring table columns
 *
 * @example
 * const col = new TableColumn();
 * col.setHeader('Sales')
 *    .setTotalFunction('sum');
 */
export class TableColumn {
  /** Internal native binding */
  _native: RuscTableColumn;

  /** Create a new table column */
  constructor() {
    this._native = RuscTableColumn.new();
  }

  /**
   * Set the header name for the column
   * @param header - Column header text
   */
  setHeader(header: string): this {
    this._native.setHeader(header);
    return this;
  }

  /**
   * Set the total function for the column
   * @param func - Function: 'average', 'count', 'countNumbers', 'max', 'min', 'stdDev', 'sum', 'var'
   */
  setTotalFunction(
    func: 'average' | 'count' | 'countNumbers' | 'max' | 'min' | 'stdDev' | 'sum' | 'var',
  ): this {
    this._native.setTotalFunction(func);
    return this;
  }

  /**
   * Set a custom total label for the column
   * @param label - Total row label text
   */
  setTotalLabel(label: string): this {
    this._native.setTotalLabel(label);
    return this;
  }

  /**
   * Set a custom formula for the column
   * @param formula - Excel formula (without '=')
   */
  setFormula(formula: string): this {
    this._native.setFormula(formula);
    return this;
  }
}

/**
 * Sparkline class for creating Excel sparklines
 *
 * @example
 * const sparkline = new Sparkline();
 * sparkline.setRange('Sheet1!A1:E1')
 *          .setType('column')
 *          .setStyle(12);
 * worksheet.addSparkline(0, 5, sparkline);
 */
export class Sparkline {
  /** Internal native binding */
  _native: RuscSparkline;

  /** Create a new sparkline */
  constructor() {
    this._native = RuscSparkline.new();
  }

  /**
   * Set the data range for the sparkline
   * @param range - Range string (e.g., 'Sheet1!A1:E1' or 'A1:E1'); a range without a
   *   sheet name refers to the worksheet the sparkline is added to
   */
  setRange(range: string): this {
    this._native.setRange(range);
    return this;
  }

  /**
   * Set the type of sparkline
   * @param type - Sparkline type: 'line', 'column', or 'winLose'
   */
  setType(type: 'line' | 'column' | 'winLose'): this {
    this._native.setType(type);
    return this;
  }

  /**
   * Show markers for the highest point
   * @param enable - Enable high point marker
   */
  showHighPoint(enable: boolean): this {
    this._native.showHighPoint(enable);
    return this;
  }

  /**
   * Show markers for the lowest point
   * @param enable - Enable low point marker
   */
  showLowPoint(enable: boolean): this {
    this._native.showLowPoint(enable);
    return this;
  }

  /**
   * Show markers for the first point
   * @param enable - Enable first point marker
   */
  showFirstPoint(enable: boolean): this {
    this._native.showFirstPoint(enable);
    return this;
  }

  /**
   * Show markers for the last point
   * @param enable - Enable last point marker
   */
  showLastPoint(enable: boolean): this {
    this._native.showLastPoint(enable);
    return this;
  }

  /**
   * Show markers for negative points
   * @param enable - Enable negative point markers
   */
  showNegativePoints(enable: boolean): this {
    this._native.showNegativePoints(enable);
    return this;
  }

  /**
   * Show markers for all points
   * @param enable - Enable all point markers
   */
  showMarkers(enable: boolean): this {
    this._native.showMarkers(enable);
    return this;
  }

  /**
   * Show the horizontal axis
   * @param enable - Enable axis display
   */
  showAxis(enable: boolean): this {
    this._native.showAxis(enable);
    return this;
  }

  /**
   * Show data from hidden rows/columns
   * @param enable - Enable hidden data display
   */
  showHiddenData(enable: boolean): this {
    this._native.showHiddenData(enable);
    return this;
  }

  /**
   * Display sparkline right to left
   * @param enable - Enable right-to-left direction
   */
  setRightToLeft(enable: boolean): this {
    this._native.setRightToLeft(enable);
    return this;
  }

  /**
   * Set the color of the sparkline
   * @param color - Hex color (e.g., '#FF0000')
   */
  setSparklineColor(color: string): this {
    this._native.setSparklineColor(color);
    return this;
  }

  /**
   * Set the color of the high point marker
   * @param color - Hex color (e.g., '#00FF00')
   */
  setHighPointColor(color: string): this {
    this._native.setHighPointColor(color);
    return this;
  }

  /**
   * Set the color of the low point marker
   * @param color - Hex color (e.g., '#FF0000')
   */
  setLowPointColor(color: string): this {
    this._native.setLowPointColor(color);
    return this;
  }

  /**
   * Set the color of the first point marker
   * @param color - Hex color (e.g., '#0000FF')
   */
  setFirstPointColor(color: string): this {
    this._native.setFirstPointColor(color);
    return this;
  }

  /**
   * Set the color of the last point marker
   * @param color - Hex color (e.g., '#0000FF')
   */
  setLastPointColor(color: string): this {
    this._native.setLastPointColor(color);
    return this;
  }

  /**
   * Set the color of negative point markers
   * @param color - Hex color (e.g., '#FFA500')
   */
  setNegativePointsColor(color: string): this {
    this._native.setNegativePointsColor(color);
    return this;
  }

  /**
   * Set the color of all point markers
   * @param color - Hex color (e.g., '#000000')
   */
  setMarkersColor(color: string): this {
    this._native.setMarkersColor(color);
    return this;
  }

  /**
   * Set the line weight/width
   * @param weight - Line weight (default: 0.75)
   */
  setLineWeight(weight: number): this {
    this._native.setLineWeight(weight);
    return this;
  }

  /**
   * Set the custom maximum value for vertical axis
   * @param max - Maximum value
   */
  setCustomMax(max: number): this {
    this._native.setCustomMax(max);
    return this;
  }

  /**
   * Set the custom minimum value for vertical axis
   * @param min - Minimum value
   */
  setCustomMin(min: number): this {
    this._native.setCustomMin(min);
    return this;
  }

  /**
   * Set maximum value based on group maximum
   * @param enable - Use group maximum
   */
  setGroupMax(enable: boolean): this {
    this._native.setGroupMax(enable);
    return this;
  }

  /**
   * Set minimum value based on group minimum
   * @param enable - Use group minimum
   */
  setGroupMin(enable: boolean): this {
    this._native.setGroupMin(enable);
    return this;
  }

  /**
   * Set a predefined style
   * @param style - Style number (1-36)
   */
  setStyle(style: number): this {
    this._native.setStyle(style);
    return this;
  }

  /**
   * Set how empty cells are displayed
   * @param option - Display option: 'gaps', 'zero', or 'connected'
   */
  showEmptyCellsAs(option: 'gaps' | 'zero' | 'connected'): this {
    this._native.showEmptyCellsAs(option);
    return this;
  }
}

/**
 * DocProperties class for setting document metadata
 *
 * @example
 * const props = new DocProperties();
 * props.setAuthor('John Doe')
 *      .setTitle('Sales Report')
 *      .setCompany('ACME Corp');
 * workbook.setProperties(props);
 */
export class DocProperties {
  /** Internal native binding */
  _native: RuscDocProperties;

  /** Create a new DocProperties instance */
  constructor() {
    this._native = RuscDocProperties.new();
  }

  /**
   * Set the document author
   * @param author - Author name
   */
  setAuthor(author: string): this {
    this._native.setAuthor(author);
    return this;
  }

  /**
   * Set the document title
   * @param title - Document title
   */
  setTitle(title: string): this {
    this._native.setTitle(title);
    return this;
  }

  /**
   * Set the document subject
   * @param subject - Document subject
   */
  setSubject(subject: string): this {
    this._native.setSubject(subject);
    return this;
  }

  /**
   * Set the document comment/description
   * @param comment - Document comment
   */
  setComment(comment: string): this {
    this._native.setComment(comment);
    return this;
  }

  /**
   * Set the company name
   * @param company - Company name
   */
  setCompany(company: string): this {
    this._native.setCompany(company);
    return this;
  }

  /**
   * Set the manager name
   * @param manager - Manager name
   */
  setManager(manager: string): this {
    this._native.setManager(manager);
    return this;
  }

  /**
   * Set the document status
   * @param status - Document status (e.g., 'Draft', 'Final')
   */
  setStatus(status: string): this {
    this._native.setStatus(status);
    return this;
  }

  /**
   * Set the document category
   * @param category - Document category
   */
  setCategory(category: string): this {
    this._native.setCategory(category);
    return this;
  }

  /**
   * Set document keywords for searching
   * @param keywords - Comma-separated keywords
   */
  setKeywords(keywords: string): this {
    this._native.setKeywords(keywords);
    return this;
  }

  /**
   * Set the hyperlink base URL
   * @param base - Base URL for relative hyperlinks
   */
  setHyperlinkBase(base: string): this {
    this._native.setHyperlinkBase(base);
    return this;
  }
}

/**
 * Button class for Excel form control buttons
 *
 * @example
 * const button = new Button();
 * button.setCaption('Click Me')
 *       .setMacro('MyMacro')
 *       .setWidth(100)
 *       .setHeight(30);
 * worksheet.insertButton(0, 0, button);
 */
export class Button {
  /** Internal native binding */
  _native: RuscButton;

  /** Create a new button */
  constructor() {
    this._native = RuscButton.new();
  }

  /**
   * Set the button caption/text
   * @param caption - Text displayed on the button
   */
  setCaption(caption: string): this {
    this._native.setCaption(caption);
    return this;
  }

  /**
   * Set the VBA macro to run when button is clicked
   * @param macroName - Name of the VBA macro
   */
  setMacro(macroName: string): this {
    this._native.setMacro(macroName);
    return this;
  }

  /**
   * Set the button width in pixels
   * @param width - Width in pixels
   */
  setWidth(width: number): this {
    this._native.setWidth(width);
    return this;
  }

  /**
   * Set the button height in pixels
   * @param height - Height in pixels
   */
  setHeight(height: number): this {
    this._native.setHeight(height);
    return this;
  }

  /**
   * Set alt text for accessibility
   * @param altText - Accessibility description
   */
  setAltText(altText: string): this {
    this._native.setAltText(altText);
    return this;
  }

  /**
   * Set how the button moves and sizes with cells
   * @param option - Movement option: 'moveAndSize', 'moveOnly', 'none', or 'moveAndSizeAfter'
   */
  setObjectMovement(option: 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'): this {
    this._native.setObjectMovement(option);
    return this;
  }
}

/**
 * Shape class for Excel shapes (textboxes)
 *
 * @example
 * const shape = Shape.textbox();
 * shape.setText('Important Note')
 *      .setWidth(200)
 *      .setHeight(100);
 * worksheet.insertShape(0, 0, shape);
 */
export class Shape {
  /** Internal native binding */
  _native: RuscShape;

  /** Shapes are created through factories such as {@link Shape.textbox}. */
  private constructor(native: RuscShape) {
    this._native = native;
  }

  /**
   * Create a new textbox shape
   * @returns Shape instance configured as a textbox
   */
  static textbox(): Shape {
    return new Shape(RuscShape.textbox());
  }

  /**
   * Set the text content of the textbox
   * @param text - Text content
   */
  setText(text: string): this {
    this._native.setText(text);
    return this;
  }

  /**
   * Set the shape width in pixels
   * @param width - Width in pixels
   */
  setWidth(width: number): this {
    this._native.setWidth(width);
    return this;
  }

  /**
   * Set the shape height in pixels
   * @param height - Height in pixels
   */
  setHeight(height: number): this {
    this._native.setHeight(height);
    return this;
  }

  /**
   * Set alt text for accessibility
   * @param altText - Accessibility description
   */
  setAltText(altText: string): this {
    this._native.setAltText(altText);
    return this;
  }

  /**
   * Set how the shape moves and sizes with cells
   * @param option - Movement option: 'moveAndSize', 'moveOnly', 'none', or 'moveAndSizeAfter'
   */
  setObjectMovement(option: 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'): this {
    this._native.setObjectMovement(option);
    return this;
  }
}

/**
 * Workbook class for Excel file creation
 *
 * @example
 * const workbook = new Workbook();
 * const worksheet = workbook.addWorksheet('Sheet1');
 * worksheet.write(0, 0, 'Hello World');
 * workbook.save('output.xlsx');
 */
export class Workbook {
  /** Internal native binding */
  _native: RuscWorkbook;

  /** Create a new workbook */
  constructor() {
    this._native = new RuscWorkbook();
  }

  /**
   * Add a new worksheet to the workbook
   * @param name - Optional worksheet name
   * @returns Worksheet instance
   */
  addWorksheet(name?: string): Worksheet {
    const nativeWorksheet = this._native.addWorksheet(name);
    return new Worksheet(nativeWorksheet);
  }

  /**
   * Add a new worksheet with constant memory mode (minimal memory footprint)
   * Sequential row writing only, limited merge support, no tables.
   * @param name - Optional worksheet name
   * @returns Worksheet instance in constant memory mode
   */
  addWorksheetWithConstantMemory(name?: string): Worksheet {
    const nativeWorksheet = this._native.addWorksheetWithConstantMemory(name);
    return new Worksheet(nativeWorksheet);
  }

  /**
   * Add a new worksheet with low memory mode (balanced optimization)
   * Good for large datasets with repeated strings.
   * @param name - Optional worksheet name
   * @returns Worksheet instance in low memory mode
   */
  addWorksheetWithLowMemory(name?: string): Worksheet {
    const nativeWorksheet = this._native.addWorksheetWithLowMemory(name);
    return new Worksheet(nativeWorksheet);
  }

  /**
   * Set custom temporary directory for constant/low memory modes
   * Must be called BEFORE adding memory mode worksheets
   * @param dir - Directory path for temporary files
   */
  setTempdir(dir: string): void {
    return this._native.setTempdir(bindingPath(dir));
  }

  /**
   * Set the default workbook format and row/column dimensions.
   * Must be called before adding worksheets.
   * @param format - Default cell format
   * @param rowHeight - Default row height in pixels
   * @param columnWidth - Supported default column width in pixels: 56, 64, 72, 80, 96, 104, or 120
   */
  setDefaultFormat(format: Format, rowHeight: number, columnWidth: number): void {
    return this._native.setDefaultFormat(format._native, rowHeight, columnWidth);
  }

  /**
   * Use the Excel 2023 Office theme with the Aptos default font.
   * Must be called before adding worksheets.
   */
  useExcel2023Theme(): void {
    return this._native.useExcel2023Theme();
  }

  /**
   * Set the zip compression level of the xlsx container.
   * Lower levels are faster, higher levels produce smaller files.
   * @param level - Compression level from 1 to 9 (default: 6)
   * @throws Error if the level is outside 1-9
   * @example
   * workbook.setZipCompressionLevel(1); // fastest
   */
  setZipCompressionLevel(level: number): void {
    return this._native.setZipCompressionLevel(level);
  }

  /**
   * Allow the xlsx file, and the files inside it, to exceed the 4 GB limit of a
   * standard zip file by using ZIP64. Without it, saving such a workbook fails.
   * Worksheets reach the limit at roughly 150 million cells. Excel reads ZIP64
   * files; some other readers don't.
   * @param enable - True to use ZIP64
   * @example
   * workbook.useZipLargeFile(true);
   */
  useZipLargeFile(enable: boolean): void {
    return this._native.useZipLargeFile(enable);
  }

  /**
   * Set when Excel recalculates the workbook's formulas
   * @param mode - `automatic` (default), `manual`, or `partial` (all but data tables)
   * @example
   * workbook.setCalculationMode('manual');
   */
  setCalculationMode(mode: CalculationMode): void {
    return this._native.setCalculationMode(mode);
  }

  /**
   * Set the width of the worksheet tab bar as a percentage of the space it
   * shares with the horizontal scroll bar
   * @param ratio - Percentage from 0 to 100 (default: 60)
   * @throws Error if the ratio is outside 0-100
   * @example
   * workbook.setTabRatio(80);
   */
  setTabRatio(ratio: number): void {
    return this._native.setTabRatio(ratio);
  }

  /**
   * Load a custom theme from a .thmx, .xlsx, or theme XML file.
   * Must be called before adding worksheets.
   */
  useCustomTheme(path: string): void {
    return this._native.useCustomTheme(bindingPath(path));
  }

  /**
   * Save the workbook to a file
   * @param filename - Path where the file should be saved
   */
  save(filename: string): void {
    return this._native.save(bindingPath(filename));
  }

  /**
   * Save the workbook to a Buffer
   *
   * In browsers the result is a `Buffer` from the `buffer` package, a
   * `Uint8Array` subclass, so it can be passed to `Blob` directly.
   * @returns Buffer containing the Excel file data
   * @example
   * // In a browser
   * const blob = new Blob([workbook.saveToBuffer()], {
   *   type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
   * });
   */
  saveToBuffer(): Buffer {
    return this._native.saveToBuffer();
  }

  /**
   * Save the workbook to a Writable stream
   *
   * Saves the workbook to a temporary file, then pipes that file into the
   * writable without buffering the entire XLSX file in JavaScript memory.
   * The temporary file is removed after the pipeline succeeds or fails.
   * The promise rejects on native save, read, writable, pipeline, or cleanup errors.
   * Node.js only: in browsers the promise rejects; use {@link saveToBuffer} there.
   *
   * @param writable - A Node.js Writable stream (e.g., fs.createWriteStream, HTTP response)
   * @returns A promise that settles after streaming and temporary-file cleanup complete
   * @example
   * const fs = require('fs');
   * const stream = fs.createWriteStream('output.xlsx');
   * await workbook.saveToStream(stream);
   *
   * @example
   * // HTTP response
   * res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
   * await workbook.saveToStream(res);
   */
  async saveToStream(writable: Writable): Promise<void> {
    const node = nodeProcess();
    if (node === undefined) {
      throw new Error('saveToStream() needs Node.js streams; use saveToBuffer() instead');
    }
    const fs = node.getBuiltinModule('node:fs');
    const os = node.getBuiltinModule('node:os');
    const { join } = node.getBuiltinModule('node:path');
    const { randomBytes } = node.getBuiltinModule('node:crypto');
    const { pipeline } = node.getBuiltinModule('node:stream/promises');
    const tmpFile = join(os.tmpdir(), `rusc-xlsx-${randomBytes(8).toString('hex')}.xlsx`);
    let operationError: unknown;

    try {
      this._native.save(tmpFile);
      await pipeline(fs.createReadStream(tmpFile), writable);
    } catch (error) {
      operationError = error;
    }

    try {
      await fs.promises.unlink(tmpFile);
    } catch (cleanupError) {
      if (!isErrnoException(cleanupError) || cleanupError.code !== 'ENOENT') {
        if (operationError) {
          throw new AggregateError(
            [operationError, cleanupError],
            'Failed to save workbook to stream and remove the temporary file',
            { cause: cleanupError },
          );
        }
        throw cleanupError;
      }
    }

    if (operationError) {
      throw operationError;
    }
  }

  /**
   * Define a named range or formula
   * @param name - Name for the range (e.g., 'Sales' or 'Sheet1!Sales' for local names)
   * @param formula - Formula or range (e.g., '=Sheet1!$A$1:$A$10' or '=SUM(A1:A10)')
   * @example
   * workbook.defineName('Sales', '=Sheet1!$A$1:$A$10');
   * workbook.defineName('Sheet1!LocalName', '=Sheet1!$B$1:$B$10');
   */
  defineName(name: string, formula: string): void {
    return this._native.defineName(name, formula);
  }

  /**
   * Set the workbook to be opened in read-only recommended mode
   * Users will be prompted to open as read-only but can choose to edit
   * @example
   * workbook.readOnlyRecommended();
   */
  readOnlyRecommended(): void {
    return this._native.readOnlyRecommended();
  }

  /**
   * Add a VBA macro file to the workbook
   * @param path - Path to the vbaProject.bin file
   * @example
   * workbook.addVbaProject('./vbaProject.bin');
   */
  addVbaProject(path: string): void {
    return this._native.addVbaProject(bindingPath(path));
  }

  /**
   * Set document properties (metadata) for the workbook
   * @param properties - DocProperties instance with metadata
   * @example
   * const props = new DocProperties();
   * props.setAuthor('Jane Doe').setTitle('Report');
   * workbook.setProperties(props);
   */
  setProperties(properties: DocProperties): void {
    return this._native.setProperties(properties._native);
  }
}

// Native bindings for advanced use
export {
  RuscFormat,
  RuscNote,
  RuscFormula,
  RuscDataValidation,
  RuscChart,
  RuscWorkbook,
  RuscWorksheet,
  RuscProtectionOptions,
  RuscRichText,
  RuscImage,
  RuscTable,
  RuscTableColumn,
  RuscSparkline,
  RuscDocProperties,
  RuscButton,
  RuscShape,
};
export type { RuscFormulaData };

// Rust allocator diagnostics; documented on the Rust functions in src/lib.rs
export { rustMemoryResetPeak, rustMemoryStats } from './index.js';
