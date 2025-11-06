// TypeScript definitions for wrapper.js
// These provide the fluent API with method chaining

/** Color value - can be a named color or hex code */
export type Color =
  | 'black' | 'blue' | 'brown' | 'cyan' | 'gray' | 'green'
  | 'lime' | 'magenta' | 'navy' | 'orange' | 'pink' | 'purple'
  | 'red' | 'silver' | 'white' | 'yellow'
  | `#${string}`; // Hex colors like #FF0000

/** Horizontal alignment options */
export type HorizontalAlign = 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerAcross' | 'distributed';

/** Vertical alignment options */
export type VerticalAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/** Border style options */
export type BorderStyle =
  | 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed'
  | 'dashDot' | 'dashDotDot' | 'slantDashDot'
  | 'mediumDashed' | 'mediumDashDot' | 'mediumDashDotDot';

/** Underline style options */
export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/** Fill pattern options */
export type FillPattern =
  | 'solid' | 'mediumGray' | 'darkGray' | 'lightGray'
  | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp'
  | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical'
  | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis'
  | 'gray125' | 'gray0625';

/** Data validation rule type */
export type ValidationRule = 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';

/** Data validation error style */
export type ValidationErrorStyle = 'stop' | 'warning' | 'information';

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
export declare class Format {
  /** Internal native binding */
  _native: any;

  /** Create a new format */
  constructor();

  // Font styling methods (return this for chaining)
  /** Set bold text */
  setBold(): this;
  /** Set italic text */
  setItalic(): this;
  /** Set font size in points */
  setFontSize(size: number): this;
  /** Set font name (e.g., 'Arial', 'Calibri', 'Times New Roman') */
  setFontName(fontName: string): this;
  /** Set font color using named color or hex code */
  setFontColor(color: Color): this;
  /** Set text underline style */
  setUnderline(underline: UnderlineStyle): this;
  /** Set strikethrough text */
  setStrikethrough(): this;

  // Fill and background methods
  /** Set background color using named color or hex code */
  setBackgroundColor(color: Color): this;
  /** Set foreground color (for patterns) using named color or hex code */
  setForegroundColor(color: Color): this;
  /** Set fill pattern */
  setPattern(pattern: FillPattern): this;

  // Alignment and text formatting methods
  /** Set horizontal alignment */
  setAlign(align: HorizontalAlign): this;
  /** Set vertical alignment */
  setVerticalAlign(align: VerticalAlign): this;
  /** Enable text wrapping in the cell */
  setTextWrap(): this;
  /** Set text indentation level (0-15) */
  setIndent(indent: number): this;
  /** Set text rotation in degrees (-90 to 90, or 270 for vertical text) */
  setRotation(rotation: number): this;
  /** Shrink text to fit cell */
  setShrink(): this;

  // Border methods
  /** Set border style for all borders */
  setBorder(border: BorderStyle): this;
  /** Set top border style */
  setBorderTop(border: BorderStyle): this;
  /** Set bottom border style */
  setBorderBottom(border: BorderStyle): this;
  /** Set left border style */
  setBorderLeft(border: BorderStyle): this;
  /** Set right border style */
  setBorderRight(border: BorderStyle): this;
  /** Set color for all borders */
  setBorderColor(color: Color): this;
  /** Set top border color */
  setBorderTopColor(color: Color): this;
  /** Set bottom border color */
  setBorderBottomColor(color: Color): this;
  /** Set left border color */
  setBorderLeftColor(color: Color): this;
  /** Set right border color */
  setBorderRightColor(color: Color): this;

  // Number format
  /**
   * Set number format using Excel format strings
   * @example
   * format.setNumFormat('#,##0.00');      // Thousands separator
   * format.setNumFormat('$#,##0.00');     // Currency
   * format.setNumFormat('0.00%');         // Percentage
   * format.setNumFormat('yyyy-mm-dd');    // Date
   */
  setNumFormat(format: string): this;

  // Cell protection
  /** Unlock cell (for use with worksheet protection) */
  setUnlocked(): this;
  /** Lock cell (cell will be protected when worksheet is protected) */
  setLocked(): this;
  /** Hide formulas in cell (for use with worksheet protection) */
  setHidden(): this;

  /**
   * Set font script (superscript or subscript)
   * @param script - 'superscript' for X² or 'subscript' for H₂O
   * @example
   * format.setFontScript('superscript'); // X²
   * format.setFontScript('subscript');   // H₂O
   */
  setFontScript(script: 'superscript' | 'subscript'): this;

  /**
   * Set diagonal border style
   * @param border - Border style
   * @example
   * format.setBorderDiagonal('thin');
   */
  setBorderDiagonal(border: BorderStyle): this;

  /**
   * Set diagonal border color
   * @param color - Color for diagonal border
   */
  setBorderDiagonalColor(color: Color): this;

  /**
   * Set diagonal border type
   * @param type - 'up' for / border, 'down' for \ border, 'both' for X
   * @example
   * format.setBorderDiagonalType('up');   // /
   * format.setBorderDiagonalType('down'); // \
   * format.setBorderDiagonalType('both'); // X
   */
  setBorderDiagonalType(type: 'up' | 'down' | 'both'): this;

  /**
   * Set reading direction for RTL languages
   * @param direction - 0=context, 1=left-to-right, 2=right-to-left
   * @example
   * format.setReadingDirection(2); // Right-to-left (Arabic, Hebrew)
   */
  setReadingDirection(direction: 0 | 1 | 2): this;

  /**
   * Force string interpretation with leading apostrophe
   * Useful for numbers that should be treated as text
   * @example
   * format.setQuotePrefix(); // '00123 will display as 00123
   */
  setQuotePrefix(): this;
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
export declare class Note {
  /** Internal native binding */
  _native: any;

  /**
   * Create a new note with text
   * @param text - The text content of the note
   */
  constructor(text: string);

  /**
   * Set the author name for the note (max 52 characters)
   * @param name - Author name
   */
  setAuthor(name: string): this;

  /**
   * Set the width of the note in pixels
   * @param width - Width in pixels
   */
  setWidth(width: number): this;

  /**
   * Set the height of the note in pixels
   * @param height - Height in pixels
   */
  setHeight(height: number): this;

  /**
   * Set the background color of the note
   * @param color - Color name or hex code
   */
  setBackgroundColor(color: Color): this;

  /**
   * Make the note visible by default (normally hidden until hover)
   * @param visible - True to show by default, false to show on hover
   */
  setVisible(visible: boolean): this;

  /**
   * Set alt text for accessibility
   * @param altText - Accessibility description
   */
  setAltText(altText: string): this;
}

/**
 * Formula class for Excel formulas with fluent API
 *
 * @example
 * const formula = new Formula('=SUM(A1:A10)');
 * const formulaWithResult = new Formula('=2+2').setResult('4');
 */
export declare class Formula {
  /** Internal native binding */
  _native: any;

  /**
   * Create a new formula
   * @param formula - Excel formula string (e.g., '=SUM(A1:A10)')
   */
  constructor(formula: string);

  /**
   * Set the result value for applications that don't auto-calculate formulas
   * @param result - The calculated result as a string
   */
  setResult(result: string): this;
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
export declare class DataValidation {
  /** Internal native binding */
  _native: any;

  /** Create a new data validation */
  constructor();

  /**
   * Allow whole number validation with a rule
   * @param rule - Validation rule type
   * @param value - Comparison value
   */
  allowWholeNumber(rule: ValidationRule, value: number): this;

  /**
   * Allow whole number validation between min and max
   * @param min - Minimum value (inclusive)
   * @param max - Maximum value (inclusive)
   */
  allowWholeNumberBetween(min: number, max: number): this;

  /**
   * Allow decimal number validation with a rule
   * @param rule - Validation rule type
   * @param value - Comparison value
   */
  allowDecimalNumber(rule: ValidationRule, value: number): this;

  /**
   * Allow decimal number validation between min and max
   * @param min - Minimum value (inclusive)
   * @param max - Maximum value (inclusive)
   */
  allowDecimalNumberBetween(min: number, max: number): this;

  /**
   * Allow list validation with an array of strings (creates dropdown)
   * @param list - Array of allowed values (max 255 chars total)
   */
  allowListStrings(list: string[]): this;

  /**
   * Allow list validation from a cell range formula
   * @param formula - Cell range formula (e.g., '=$A$1:$A$10')
   */
  allowListFormula(formula: string): this;

  /**
   * Allow custom formula validation
   * @param formula - Excel formula that returns true/false
   */
  allowCustom(formula: string): this;

  /**
   * Allow any value (mainly used for displaying input messages)
   */
  allowAnyValue(): this;

  /**
   * Set whether to ignore blank cells
   * @param enable - True to allow blank cells (default: true)
   */
  ignoreBlank(enable: boolean): this;

  /**
   * Set whether to show dropdown arrow for list validations
   * @param enable - True to show dropdown (default: true)
   */
  showDropdown(enable: boolean): this;

  /**
   * Set whether to show input message when cell is selected
   * @param enable - True to show message (default: true)
   */
  showInputMessage(enable: boolean): this;

  /**
   * Set whether to show error message on invalid input
   * @param enable - True to show message (default: true)
   */
  showErrorMessage(enable: boolean): this;

  /**
   * Set the input message title (shown when cell is selected)
   * @param title - Title text (max 32 characters)
   */
  setInputTitle(title: string): this;

  /**
   * Set the input message text (shown when cell is selected)
   * @param message - Message text (max 255 characters)
   */
  setInputMessage(message: string): this;

  /**
   * Set the error message title (shown on invalid input)
   * @param title - Title text (max 32 characters)
   */
  setErrorTitle(title: string): this;

  /**
   * Set the error message text (shown on invalid input)
   * @param message - Message text (max 255 characters)
   */
  setErrorMessage(message: string): this;

  /**
   * Set the error dialog style
   * @param style - Error style: 'stop' (red X), 'warning' (yellow !), or 'information' (blue i)
   */
  setErrorStyle(style: ValidationErrorStyle): this;
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
export declare class ProtectionOptions {
  /** Internal native binding */
  _native: any;

  /** Create a new ProtectionOptions with default settings */
  constructor();

  /**
   * Allow or prevent selecting locked cells (default: true)
   * @param allow - true to allow, false to prevent
   */
  setSelectLockedCells(allow: boolean): this;

  /**
   * Allow or prevent selecting unlocked cells (default: true)
   * @param allow - true to allow, false to prevent
   */
  setSelectUnlockedCells(allow: boolean): this;

  /**
   * Allow or prevent formatting cells (default: false)
   * @param allow - true to allow, false to prevent
   */
  setFormatCells(allow: boolean): this;

  /**
   * Allow or prevent formatting columns (default: false)
   * @param allow - true to allow, false to prevent
   */
  setFormatColumns(allow: boolean): this;

  /**
   * Allow or prevent formatting rows (default: false)
   * @param allow - true to allow, false to prevent
   */
  setFormatRows(allow: boolean): this;

  /**
   * Allow or prevent inserting columns (default: false)
   * @param allow - true to allow, false to prevent
   */
  setInsertColumns(allow: boolean): this;

  /**
   * Allow or prevent inserting rows (default: false)
   * @param allow - true to allow, false to prevent
   */
  setInsertRows(allow: boolean): this;

  /**
   * Allow or prevent inserting hyperlinks (default: false)
   * @param allow - true to allow, false to prevent
   */
  setInsertLinks(allow: boolean): this;

  /**
   * Allow or prevent deleting columns (default: false)
   * @param allow - true to allow, false to prevent
   */
  setDeleteColumns(allow: boolean): this;

  /**
   * Allow or prevent deleting rows (default: false)
   * @param allow - true to allow, false to prevent
   */
  setDeleteRows(allow: boolean): this;

  /**
   * Allow or prevent sorting data (default: false)
   * @param allow - true to allow, false to prevent
   */
  setSort(allow: boolean): this;

  /**
   * Allow or prevent using autofilters (default: false)
   * @param allow - true to allow, false to prevent
   */
  setUseAutofilter(allow: boolean): this;

  /**
   * Allow or prevent using pivot tables (default: false)
   * @param allow - true to allow, false to prevent
   */
  setUsePivotTables(allow: boolean): this;

  /**
   * Allow or prevent editing scenarios (default: false)
   * @param allow - true to allow, false to prevent
   */
  setEditScenarios(allow: boolean): this;

  /**
   * Allow or prevent editing objects like images (default: false)
   * @param allow - true to allow, false to prevent
   */
  setEditObjects(allow: boolean): this;

  /**
   * Allow or prevent editing contents in chartsheets (default: true)
   * @param allow - true to allow, false to prevent
   */
  setContents(allow: boolean): this;
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
export declare class Image {
  /** Internal native binding */
  _native: any;

  /**
   * Create a new Image from a file path
   * @param path - Path to the image file (PNG, JPG, GIF, or BMP)
   * @throws Error if file cannot be read or format is unsupported
   */
  constructor(path: string);

  /**
   * Create an Image from a buffer
   * @param buffer - Buffer containing image data
   * @returns Image instance
   * @example
   * const buffer = fs.readFileSync('logo.png');
   * const image = Image.fromBuffer(buffer);
   */
  static fromBuffer(buffer: Buffer): Image;

  /**
   * Set the width of the image in pixels
   * @param width - Width in pixels
   * @returns this for method chaining
   */
  setWidth(width: number): this;

  /**
   * Set the height of the image in pixels
   * @param height - Height in pixels
   * @returns this for method chaining
   */
  setHeight(height: number): this;

  /**
   * Scale the width of the image as a percentage
   * @param scale - Scale percentage (e.g., 1.5 for 150%)
   * @returns this for method chaining
   */
  setScaleWidth(scale: number): this;

  /**
   * Scale the height of the image as a percentage
   * @param scale - Scale percentage (e.g., 0.5 for 50%)
   * @returns this for method chaining
   */
  setScaleHeight(scale: number): this;

  /**
   * Scale the image to a specific size
   * @param width - Target width
   * @param height - Target height
   * @param keepAspectRatio - Whether to maintain aspect ratio (default: true)
   * @returns this for method chaining
   * @example
   * image.setScaleToSize(300, 200, true);
   */
  setScaleToSize(width: number, height: number, keepAspectRatio?: boolean): this;

  /**
   * Set alt text for the image (for accessibility)
   * @param altText - Descriptive text for the image
   * @returns this for method chaining
   */
  setAltText(altText: string): this;

  /**
   * Mark the image as decorative (no alt text required)
   * @param enable - true to mark as decorative
   * @returns this for method chaining
   */
  setDecorative(enable: boolean): this;

  /**
   * Set how the image moves and sizes with cells
   * @param option - 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'
   * @returns this for method chaining
   * @example
   * image.setObjectMovement('moveOnly'); // Image moves but doesn't resize with cells
   */
  setObjectMovement(option: 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'): this;

  /**
   * Add a URL link to the image
   * @param url - URL to link to
   * @returns this for method chaining
   * @example
   * image.setUrl('https://example.com');
   */
  setUrl(url: string): this;
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
export declare class RichText {
  /** Internal native binding */
  _native: any;

  /** Create a new empty rich text string */
  constructor();

  /**
   * Add a text segment with formatting
   * @param text - The text content
   * @param format - Optional format to apply to this segment
   * @returns this for method chaining
   * @example
   * richText.addSegment('Hello ', normalFormat)
   *         .addSegment('World', boldFormat);
   */
  addSegment(text: string, format?: Format): this;
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
export declare class Worksheet {
  /** Internal native binding */
  _native: any;

  /** @internal */
  constructor(nativeWorksheet: any);

  /**
   * Write a value to a cell (auto-detects type)
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Value to write (string, number, or boolean)
   * @param format - Optional format to apply
   */
  write(row: number, col: number, value: string | number | boolean, format?: Format): void;

  /**
   * Write a string to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - String value to write
   */
  writeString(row: number, col: number, value: string): void;

  /**
   * Write a number to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Number value to write
   */
  writeNumber(row: number, col: number, value: number): void;

  /**
   * Write a boolean to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Boolean value to write
   */
  writeBoolean(row: number, col: number, value: boolean): void;

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
  writeBlank(row: number, col: number, format: Format): void;

  /**
   * Write a datetime from a Unix timestamp (seconds since 1970-01-01 UTC)
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param timestamp - Unix timestamp in seconds
   * @param format - Optional format (should include date/time number format)
   * @example
   * const dateFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm');
   * worksheet.writeDatetimeFromTimestamp(0, 0, 1706184600, dateFormat);
   */
  writeDatetimeFromTimestamp(row: number, col: number, timestamp: number, format?: Format): void;

  /**
   * Write a datetime from year, month, day components
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param year - Year (1900-9999)
   * @param month - Month (1-12)
   * @param day - Day (1-31)
   * @param format - Optional format (should include date number format)
   * @example
   * const dateFormat = new Format().setNumFormat('yyyy-mm-dd');
   * worksheet.writeDatetimeFromYmd(0, 0, 2023, 1, 25, dateFormat);
   */
  writeDatetimeFromYmd(row: number, col: number, year: number, month: number, day: number, format?: Format): void;

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
   * @param format - Optional format (should include date/time number format)
   * @example
   * const datetimeFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm:ss');
   * worksheet.writeDatetimeFromYmdHms(0, 0, 2023, 1, 25, 12, 30, 45.5, datetimeFormat);
   */
  writeDatetimeFromYmdHms(row: number, col: number, year: number, month: number, day: number, hour: number, min: number, sec: number, format?: Format): void;

  /**
   * Write a datetime from an ISO 8601 string
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param datetimeStr - ISO 8601 date/time string (e.g., "2023-01-25T12:30:00" or "2023-01-25")
   * @param format - Optional format (should include date/time number format)
   * @example
   * const datetimeFormat = new Format().setNumFormat('yyyy-mm-dd hh:mm');
   * worksheet.writeDatetimeFromString(0, 0, '2023-01-25T12:30:00', datetimeFormat);
   */
  writeDatetimeFromString(row: number, col: number, datetimeStr: string, format?: Format): void;

  /**
   * Write a datetime (convenience method that auto-detects input type)
   * Accepts: Date object, Unix timestamp (number), ISO string, or object with date components
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param datetime - Date/time value in various formats
   * @param format - Optional format (should include date/time number format)
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
  writeDatetime(row: number, col: number, datetime: Date | number | string | { year: number, month: number, day: number, hour?: number, min?: number, sec?: number }, format?: Format): void;

  /**
   * Write a value with formatting to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param value - Value to write (string, number, or boolean)
   * @param format - Format to apply
   */
  writeWithFormat(row: number, col: number, value: string | number | boolean, format: Format): void;

  /**
   * Set the width of a column
   * @param col - Column number (0-indexed)
   * @param width - Width in Excel units (approximately character width)
   */
  setColumnWidth(col: number, width: number): void;

  /**
   * Set the height of a row
   * @param row - Row number (0-indexed)
   * @param height - Height in points
   */
  setRowHeight(row: number, height: number): void;

  /**
   * Set the worksheet name
   * @param name - Worksheet name
   */
  setName(name: string): void;

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
    format?: Format
  ): void;

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
    format: Format
  ): void;

  /**
   * Insert a note (post-it style comment) to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param note - Note object to insert
   */
  insertNote(row: number, col: number, note: Note): void;

  /**
   * Set the default author name for all notes in the worksheet
   * @param name - Author name (max 52 characters)
   */
  setDefaultNoteAuthor(name: string): void;

  /**
   * Write a formula to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param formula - Formula object
   * @param format - Optional format to apply
   */
  writeFormula(row: number, col: number, formula: Formula, format?: Format): void;

  /**
   * Write a dynamic array formula to a cell (Excel 365+)
   * Dynamic arrays automatically spill into adjacent cells
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param formula - Formula object with dynamic array function (FILTER, SORT, UNIQUE, etc.)
   * @param format - Optional format to apply
   */
  writeDynamicFormula(row: number, col: number, formula: Formula, format?: Format): void;

  /**
   * Write an array formula to a range of cells
   * @param firstRow - First row of range (0-indexed)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   * @param formula - Formula object
   * @param format - Optional format to apply
   */
  writeArrayFormula(
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    formula: Formula,
    format?: Format
  ): void;

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
    validation: DataValidation
  ): void;

  /**
   * Write a URL/hyperlink to a cell
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param url - URL string (http://, https://, file://, or internal:Sheet1!A1)
   * @param format - Optional format to apply
   */
  writeUrl(row: number, col: number, url: string, format?: Format): void;

  /**
   * Write a URL/hyperlink to a cell with custom display text
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param url - URL string
   * @param text - Display text for the link
   * @param format - Optional format to apply
   */
  writeUrlWithText(row: number, col: number, url: string, text: string, format?: Format): void;

  /**
   * Add autofilter dropdown buttons to a range
   * @param firstRow - First row of range (0-indexed, typically header row)
   * @param firstCol - First column of range (0-indexed)
   * @param lastRow - Last row of range (0-indexed)
   * @param lastCol - Last column of range (0-indexed)
   */
  autofilter(firstRow: number, firstCol: number, lastRow: number, lastCol: number): void;

  /**
   * Freeze panes at the specified row and column
   * @param row - Row to freeze above (0 = no horizontal freeze)
   * @param col - Column to freeze left of (0 = no vertical freeze)
   * @example
   * worksheet.setFreezePanes(1, 0); // Freeze top row
   * worksheet.setFreezePanes(0, 1); // Freeze first column
   * worksheet.setFreezePanes(1, 1); // Freeze both top row and first column
   */
  setFreezePanes(row: number, col: number): void;

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
    maxColor: Color
  ): void;

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
    maxColor: Color
  ): void;

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
    barColor: Color
  ): void;

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
    rule: 'greaterThan' | 'lessThan' | 'equal' | 'notEqual' | 'greaterThanOrEqual' | 'lessThanOrEqual',
    value: number,
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    rule: 'yesterday' | 'today' | 'tomorrow' | 'last7Days' | 'lastWeek' | 'thisWeek' | 'nextWeek' | 'lastMonth' | 'thisMonth' | 'nextMonth',
    format: Format
  ): void;

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
    format: Format
  ): void;

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
    iconType: 'threeArrows' | 'threeArrowsGray' | 'threeFlags' | 'threeTrafficLights' | 'threeTrafficLightsWithRim' |
              'threeSigns' | 'threeSymbols' | 'threeSymbolsCircled' | 'threeStars' | 'threeTriangles' |
              'fourArrows' | 'fourArrowsGray' | 'fourRedToBlack' | 'fourHistograms' | 'fourTrafficLights' |
              'fiveArrows' | 'fiveArrowsGray' | 'fiveHistograms' | 'fiveQuadrants' | 'fiveBoxes'
  ): void;

  /**
   * Set page margins for printing (in inches)
   * @param left - Left margin
   * @param right - Right margin
   * @param top - Top margin
   * @param bottom - Bottom margin
   * @param header - Header margin
   * @param footer - Footer margin
   */
  setMargins(left: number, right: number, top: number, bottom: number, header: number, footer: number): void;

  /**
   * Set page orientation to portrait
   */
  setPortrait(): void;

  /**
   * Set page orientation to landscape
   */
  setLandscape(): void;

  /**
   * Set paper size for printing
   * @param paperSize - Paper size code (1=Letter, 5=Legal, 9=A4, 11=A5, etc.)
   */
  setPaperSize(paperSize: number): void;

  /**
   * Set header text for printing
   * @param header - Header text (use &L, &C, &R for left/center/right, &P for page number, &D for date)
   * @example
   * worksheet.setHeader('&CReport Title'); // Centered header
   * worksheet.setHeader('&LCompany&CTitle&RPage &P'); // Left, center, and right
   */
  setHeader(header: string): void;

  /**
   * Set footer text for printing
   * @param footer - Footer text (use &L, &C, &R for left/center/right, &P for page number, &D for date)
   * @example
   * worksheet.setFooter('&CPage &P of &N'); // Centered page number
   * worksheet.setFooter('&L&D&RPage &P'); // Date left, page right
   */
  setFooter(footer: string): void;

  /**
   * Set print area (only this range will be printed)
   * @param firstRow - First row of print area (0-indexed)
   * @param firstCol - First column of print area (0-indexed)
   * @param lastRow - Last row of print area (0-indexed)
   * @param lastCol - Last column of print area (0-indexed)
   */
  setPrintArea(firstRow: number, firstCol: number, lastRow: number, lastCol: number): void;

  /**
   * Show or hide gridlines when printing
   * @param enable - True to show gridlines, false to hide
   */
  setPrintGridlines(enable: boolean): void;

  /**
   * Center worksheet horizontally on the page when printing
   * @param enable - True to center horizontally
   */
  setPrintCenterHorizontally(enable: boolean): void;

  /**
   * Center worksheet vertically on the page when printing
   * @param enable - True to center vertically
   */
  setPrintCenterVertically(enable: boolean): void;

  /**
   * Fit worksheet to N pages wide by M pages tall when printing
   * @param width - Number of pages wide (0 = no limit)
   * @param height - Number of pages tall (0 = no limit)
   * @example
   * worksheet.setPrintFitToPages(1, 0); // Fit to 1 page wide, unlimited height
   */
  setPrintFitToPages(width: number, height: number): void;

  /**
   * Set print scale as a percentage
   * @param scale - Scale percentage (10-400)
   * @example
   * worksheet.setPrintScale(75); // Print at 75% scale
   */
  setPrintScale(scale: number): void;

  /**
   * Set the first page number for printing
   * @param pageNumber - Starting page number
   */
  setPrintFirstPageNumber(pageNumber: number): void;

  /**
   * Enable/disable black and white printing
   * @param enable - True to print in black and white
   */
  setPrintBlackAndWhite(enable: boolean): void;

  /**
   * Enable/disable draft quality printing
   * @param enable - True for draft quality
   */
  setPrintDraft(enable: boolean): void;

  /**
   * Enable/disable printing of row and column headings
   * @param enable - True to print headings
   */
  setPrintHeadings(enable: boolean): void;

  /**
   * Repeat rows at the top of each printed page
   * @param firstRow - First row to repeat (0-indexed)
   * @param lastRow - Last row to repeat (0-indexed)
   * @example
   * worksheet.setRepeatRows(0, 0); // Repeat first row on each page
   * worksheet.setRepeatRows(0, 2); // Repeat first 3 rows on each page
   */
  setRepeatRows(firstRow: number, lastRow: number): void;

  /**
   * Repeat columns at the left of each printed page
   * @param firstCol - First column to repeat (0-indexed)
   * @param lastCol - Last column to repeat (0-indexed)
   * @example
   * worksheet.setRepeatColumns(0, 0); // Repeat first column on each page
   * worksheet.setRepeatColumns(0, 1); // Repeat first 2 columns on each page
   */
  setRepeatColumns(firstCol: number, lastCol: number): void;

  /**
   * Scale header and footer with document
   * @param enable - True to scale with document
   */
  setHeaderFooterScaleWithDoc(enable: boolean): void;

  /**
   * Align header and footer with page margins
   * @param enable - True to align with page margins
   */
  setHeaderFooterAlignWithPage(enable: boolean): void;

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
  insertChart(row: number, col: number, chart: Chart): void;

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
  insertChartWithOffset(row: number, col: number, chart: Chart, xOffset: number, yOffset: number): void;

  /**
   * Insert an image at the specified cell position
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @param image - Image instance
   * @example
   * const image = new Image('logo.png');
   * worksheet.insertImage(0, 0, image);
   */
  insertImage(row: number, col: number, image: Image): void;

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
  insertImageWithOffset(row: number, col: number, image: Image, xOffset: number, yOffset: number): void;

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
  insertImageFitToCell(row: number, col: number, image: Image, keepAspectRatio?: boolean): void;

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
  addTable(firstRow: number, firstCol: number, lastRow: number, lastCol: number, table: Table): void;

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
  addSparkline(row: number, col: number, sparkline: Sparkline): void;

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
  addSparklineGroup(firstRow: number, firstCol: number, lastRow: number, lastCol: number, sparkline: Sparkline): void;

  /**
   * Set the worksheet tab color
   * @param color - Color name or hex code
   * @example
   * worksheet.setTabColor('red');
   * worksheet.setTabColor('#FF0000');
   */
  setTabColor(color: Color): void;

  /**
   * Hide or show the worksheet
   * @param enable - True to hide worksheet, false to show
   * @example
   * worksheet.setHidden(true); // Hide worksheet
   */
  setHidden(enable: boolean): void;

  /**
   * Set the worksheet zoom level
   * @param zoom - Zoom percentage (10-400)
   * @example
   * worksheet.setZoom(150); // 150% zoom
   */
  setZoom(zoom: number): void;

  /**
   * Set the default cell selection when opening the worksheet
   * @param firstRow - First row of selection (0-indexed)
   * @param firstCol - First column of selection (0-indexed)
   * @param lastRow - Last row of selection (0-indexed)
   * @param lastCol - Last column of selection (0-indexed)
   * @example
   * worksheet.setSelection(0, 0, 5, 5); // Select A1:F6
   */
  setSelection(firstRow: number, firstCol: number, lastRow: number, lastCol: number): void;

  /**
   * Set which cell appears in the top-left corner when opening the worksheet
   * @param row - Row number (0-indexed)
   * @param col - Column number (0-indexed)
   * @example
   * worksheet.setTopLeftCell(10, 0); // Start view at row 11
   */
  setTopLeftCell(row: number, col: number): void;

  /**
   * Protect the worksheet from editing
   * Cells must be unlocked via format.setUnlocked() to remain editable
   * @example
   * worksheet.protect(); // Protect all cells
   */
  protect(): void;

  /**
   * Protect the worksheet with a password
   * @param password - Password to protect worksheet
   * @example
   * worksheet.protectWithPassword('secret123');
   */
  protectWithPassword(password: string): void;

  /**
   * Protect the worksheet with granular protection options
   * @param options - ProtectionOptions instance with desired settings
   * @example
   * const options = new ProtectionOptions()
   *   .setInsertRows(true)
   *   .setInsertColumns(true);
   * worksheet.protectWithOptions(options);
   */
  protectWithOptions(options: ProtectionOptions): void;

  /**
   * Protect the worksheet with a password and granular protection options
   * @param password - Password to protect worksheet
   * @param options - ProtectionOptions instance with desired settings
   * @example
   * const options = new ProtectionOptions()
   *   .setInsertRows(true);
   * worksheet.protectWithPasswordAndOptions('secret123', options);
   */
  protectWithPasswordAndOptions(password: string, options: ProtectionOptions): void;

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
  writeRichString(row: number, col: number, richText: RichText): void;

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
  writeRichStringWithFormat(row: number, col: number, richText: RichText, cellFormat: Format): void;

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
  insertButton(row: number, col: number, button: Button): void;

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
  insertButtonWithOffset(row: number, col: number, button: Button, xOffset: number, yOffset: number): void;

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
  insertShape(row: number, col: number, shape: Shape): void;

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
  insertShapeWithOffset(row: number, col: number, shape: Shape, xOffset: number, yOffset: number): void;
}

/** Chart type options */
export type ChartType =
  | 'area' | 'areaStacked' | 'area_stacked' | 'areaPercentStacked' | 'area_percent_stacked'
  | 'bar' | 'barStacked' | 'bar_stacked' | 'barPercentStacked' | 'bar_percent_stacked'
  | 'column' | 'columnStacked' | 'column_stacked' | 'columnPercentStacked' | 'column_percent_stacked'
  | 'doughnut'
  | 'line' | 'lineStacked' | 'line_stacked' | 'linePercentStacked' | 'line_percent_stacked'
  | 'pie'
  | 'radar' | 'radarWithMarkers' | 'radar_with_markers' | 'radarFilled' | 'radar_filled'
  | 'scatter' | 'scatterStraight' | 'scatter_straight'
  | 'scatterStraightWithMarkers' | 'scatter_straight_with_markers'
  | 'scatterSmooth' | 'scatter_smooth'
  | 'scatterSmoothWithMarkers' | 'scatter_smooth_with_markers'
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
export declare class Chart {
  /** Internal native binding */
  _native: any;

  /**
   * Create a new chart
   * @param chartType - Type of chart to create
   * @example
   * const chart = new Chart('column');
   * const pieChart = new Chart('pie');
   * const lineChart = new Chart('line');
   */
  constructor(chartType: ChartType);

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
  addSeries(options: ChartSeriesOptions): void;

  /**
   * Set the chart title
   * @param title - Chart title text
   * @example
   * chart.setTitle('Q1 Sales Results');
   */
  setTitle(title: string): this;

  /**
   * Set the X-axis label
   * @param name - X-axis label text
   * @example
   * chart.setXAxisName('Month');
   */
  setXAxisName(name: string): this;

  /**
   * Set the Y-axis label
   * @param name - Y-axis label text
   * @example
   * chart.setYAxisName('Revenue ($)');
   */
  setYAxisName(name: string): this;

  /**
   * Set the chart style (1-48)
   * @param style - Excel chart style number (1-48)
   * @example
   * chart.setStyle(11); // Apply Excel style 11
   */
  setStyle(style: number): this;

  /**
   * Hide the chart legend
   * @example
   * chart.setLegendHidden();
   */
  setLegendHidden(): this;
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
export declare class Table {
  /** Internal native binding */
  _native: any;

  /** Create a new table */
  constructor();

  /**
   * Set whether the table has a header row
   * @param show - Show header row (default: true)
   */
  setHeaderRow(show: boolean): this;

  /**
   * Set whether the table has a total row
   * @param show - Show total row (default: false)
   */
  setTotalRow(show: boolean): this;

  /**
   * Set whether the table has banded rows
   * @param show - Show banded rows (default: true)
   */
  setBandedRows(show: boolean): this;

  /**
   * Set whether the table has banded columns
   * @param show - Show banded columns (default: false)
   */
  setBandedColumns(show: boolean): this;

  /**
   * Set whether the first column is formatted
   * @param show - Format first column (default: false)
   */
  setFirstColumn(show: boolean): this;

  /**
   * Set whether the last column is formatted
   * @param show - Format last column (default: false)
   */
  setLastColumn(show: boolean): this;

  /**
   * Set the name of the table
   * @param name - Table name (must be unique)
   */
  setName(name: string): this;

  /**
   * Set the style of the table
   * @param styleId - Table style ID (1-21: light, 22-49: medium, 50-60: dark)
   * @example
   * table.setStyle(9); // Medium style 9 (default)
   */
  setStyle(styleId: number): this;

  /**
   * Set columns for the table
   * @param columns - Array of TableColumn objects
   */
  setColumns(columns: TableColumn[]): this;
}

/**
 * TableColumn class for configuring table columns
 *
 * @example
 * const col = new TableColumn();
 * col.setHeader('Sales')
 *    .setTotalFunction('sum');
 */
export declare class TableColumn {
  /** Internal native binding */
  _native: any;

  /** Create a new table column */
  constructor();

  /**
   * Set the header name for the column
   * @param header - Column header text
   */
  setHeader(header: string): this;

  /**
   * Set the total function for the column
   * @param func - Function: 'average', 'count', 'countNumbers', 'max', 'min', 'stdDev', 'sum', 'var'
   */
  setTotalFunction(func: 'average' | 'count' | 'countNumbers' | 'max' | 'min' | 'stdDev' | 'sum' | 'var'): this;

  /**
   * Set a custom total label for the column
   * @param label - Total row label text
   */
  setTotalLabel(label: string): this;

  /**
   * Set a custom formula for the column
   * @param formula - Excel formula (without '=')
   */
  setFormula(formula: string): this;
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
export declare class Sparkline {
  /** Internal native binding */
  _native: any;

  /** Create a new sparkline */
  constructor();

  /**
   * Set the data range for the sparkline
   * @param range - Range string (e.g., 'Sheet1!A1:E1' or 'A1:E1')
   */
  setRange(range: string): this;

  /**
   * Set the type of sparkline
   * @param type - Sparkline type: 'line', 'column', or 'winLose'
   */
  setType(type: 'line' | 'column' | 'winLose'): this;

  /**
   * Show markers for the highest point
   * @param enable - Enable high point marker
   */
  showHighPoint(enable: boolean): this;

  /**
   * Show markers for the lowest point
   * @param enable - Enable low point marker
   */
  showLowPoint(enable: boolean): this;

  /**
   * Show markers for the first point
   * @param enable - Enable first point marker
   */
  showFirstPoint(enable: boolean): this;

  /**
   * Show markers for the last point
   * @param enable - Enable last point marker
   */
  showLastPoint(enable: boolean): this;

  /**
   * Show markers for negative points
   * @param enable - Enable negative point markers
   */
  showNegativePoints(enable: boolean): this;

  /**
   * Show markers for all points
   * @param enable - Enable all point markers
   */
  showMarkers(enable: boolean): this;

  /**
   * Show the horizontal axis
   * @param enable - Enable axis display
   */
  showAxis(enable: boolean): this;

  /**
   * Show data from hidden rows/columns
   * @param enable - Enable hidden data display
   */
  showHiddenData(enable: boolean): this;

  /**
   * Display sparkline right to left
   * @param enable - Enable right-to-left direction
   */
  setRightToLeft(enable: boolean): this;

  /**
   * Set the color of the sparkline
   * @param color - Hex color (e.g., '#FF0000')
   */
  setSparklineColor(color: string): this;

  /**
   * Set the color of the high point marker
   * @param color - Hex color (e.g., '#00FF00')
   */
  setHighPointColor(color: string): this;

  /**
   * Set the color of the low point marker
   * @param color - Hex color (e.g., '#FF0000')
   */
  setLowPointColor(color: string): this;

  /**
   * Set the color of the first point marker
   * @param color - Hex color (e.g., '#0000FF')
   */
  setFirstPointColor(color: string): this;

  /**
   * Set the color of the last point marker
   * @param color - Hex color (e.g., '#0000FF')
   */
  setLastPointColor(color: string): this;

  /**
   * Set the color of negative point markers
   * @param color - Hex color (e.g., '#FFA500')
   */
  setNegativePointsColor(color: string): this;

  /**
   * Set the color of all point markers
   * @param color - Hex color (e.g., '#000000')
   */
  setMarkersColor(color: string): this;

  /**
   * Set the line weight/width
   * @param weight - Line weight (default: 0.75)
   */
  setLineWeight(weight: number): this;

  /**
   * Set the custom maximum value for vertical axis
   * @param max - Maximum value
   */
  setCustomMax(max: number): this;

  /**
   * Set the custom minimum value for vertical axis
   * @param min - Minimum value
   */
  setCustomMin(min: number): this;

  /**
   * Set maximum value based on group maximum
   * @param enable - Use group maximum
   */
  setGroupMax(enable: boolean): this;

  /**
   * Set minimum value based on group minimum
   * @param enable - Use group minimum
   */
  setGroupMin(enable: boolean): this;

  /**
   * Set a predefined style
   * @param style - Style number (1-36)
   */
  setStyle(style: number): this;

  /**
   * Set how empty cells are displayed
   * @param option - Display option: 'gaps', 'zero', or 'connected'
   */
  showEmptyCellsAs(option: 'gaps' | 'zero' | 'connected'): this;
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
export declare class DocProperties {
  /** Internal native binding */
  _native: any;

  /** Create a new DocProperties instance */
  constructor();

  /**
   * Set the document author
   * @param author - Author name
   */
  setAuthor(author: string): this;

  /**
   * Set the document title
   * @param title - Document title
   */
  setTitle(title: string): this;

  /**
   * Set the document subject
   * @param subject - Document subject
   */
  setSubject(subject: string): this;

  /**
   * Set the document comment/description
   * @param comment - Document comment
   */
  setComment(comment: string): this;

  /**
   * Set the company name
   * @param company - Company name
   */
  setCompany(company: string): this;

  /**
   * Set the manager name
   * @param manager - Manager name
   */
  setManager(manager: string): this;

  /**
   * Set the document status
   * @param status - Document status (e.g., 'Draft', 'Final')
   */
  setStatus(status: string): this;

  /**
   * Set the document category
   * @param category - Document category
   */
  setCategory(category: string): this;

  /**
   * Set document keywords for searching
   * @param keywords - Comma-separated keywords
   */
  setKeywords(keywords: string): this;

  /**
   * Set the hyperlink base URL
   * @param base - Base URL for relative hyperlinks
   */
  setHyperlinkBase(base: string): this;
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
export declare class Button {
  /** Internal native binding */
  _native: any;

  /** Create a new button */
  constructor();

  /**
   * Set the button caption/text
   * @param caption - Text displayed on the button
   */
  setCaption(caption: string): this;

  /**
   * Set the VBA macro to run when button is clicked
   * @param macroName - Name of the VBA macro
   */
  setMacro(macroName: string): this;

  /**
   * Set the button width in pixels
   * @param width - Width in pixels
   */
  setWidth(width: number): this;

  /**
   * Set the button height in pixels
   * @param height - Height in pixels
   */
  setHeight(height: number): this;

  /**
   * Set alt text for accessibility
   * @param altText - Accessibility description
   */
  setAltText(altText: string): this;

  /**
   * Set how the button moves and sizes with cells
   * @param option - Movement option: 'moveAndSize', 'moveOnly', 'none', or 'moveAndSizeAfter'
   */
  setObjectMovement(option: 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'): this;
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
export declare class Shape {
  /** Internal native binding */
  _native: any;

  /**
   * Create a new textbox shape
   * @returns Shape instance configured as a textbox
   */
  static textbox(): Shape;

  /**
   * Set the text content of the textbox
   * @param text - Text content
   */
  setText(text: string): this;

  /**
   * Set the shape width in pixels
   * @param width - Width in pixels
   */
  setWidth(width: number): this;

  /**
   * Set the shape height in pixels
   * @param height - Height in pixels
   */
  setHeight(height: number): this;

  /**
   * Set alt text for accessibility
   * @param altText - Accessibility description
   */
  setAltText(altText: string): this;

  /**
   * Set how the shape moves and sizes with cells
   * @param option - Movement option: 'moveAndSize', 'moveOnly', 'none', or 'moveAndSizeAfter'
   */
  setObjectMovement(option: 'moveAndSize' | 'moveOnly' | 'none' | 'moveAndSizeAfter'): this;
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
export declare class Workbook {
  /** Internal native binding */
  _native: any;

  /** Create a new workbook */
  constructor();

  /**
   * Add a new worksheet to the workbook
   * @param name - Optional worksheet name
   * @returns Worksheet instance
   */
  addWorksheet(name?: string): Worksheet;

  /**
   * Add a new worksheet with constant memory mode (minimal memory footprint)
   * Sequential row writing only, limited merge support, no tables.
   * @param name - Optional worksheet name
   * @returns Worksheet instance in constant memory mode
   */
  addWorksheetWithConstantMemory(name?: string): Worksheet;

  /**
   * Add a new worksheet with low memory mode (balanced optimization)
   * Good for large datasets with repeated strings.
   * @param name - Optional worksheet name
   * @returns Worksheet instance in low memory mode
   */
  addWorksheetWithLowMemory(name?: string): Worksheet;

  /**
   * Set custom temporary directory for constant/low memory modes
   * Must be called BEFORE adding memory mode worksheets
   * @param dir - Directory path for temporary files
   */
  setTempdir(dir: string): void;

  /**
   * Save the workbook to a file
   * @param filename - Path where the file should be saved
   */
  save(filename: string): void;

  /**
   * Save the workbook to a Buffer
   * @returns Buffer containing the Excel file data
   */
  saveToBuffer(): Buffer;

  /**
   * Define a named range or formula
   * @param name - Name for the range (e.g., 'Sales' or 'Sheet1!Sales' for local names)
   * @param formula - Formula or range (e.g., '=Sheet1!$A$1:$A$10' or '=SUM(A1:A10)')
   * @example
   * workbook.defineName('Sales', '=Sheet1!$A$1:$A$10');
   * workbook.defineName('Sheet1!LocalName', '=Sheet1!$B$1:$B$10');
   */
  defineName(name: string, formula: string): void;

  /**
   * Set the workbook to be opened in read-only recommended mode
   * Users will be prompted to open as read-only but can choose to edit
   * @example
   * workbook.readOnlyRecommended();
   */
  readOnlyRecommended(): void;

  /**
   * Add a VBA macro file to the workbook
   * @param path - Path to the vbaProject.bin file
   * @example
   * workbook.addVbaProject('./vbaProject.bin');
   */
  addVbaProject(path: string): void;

  /**
   * Set document properties (metadata) for the workbook
   * @param properties - DocProperties instance with metadata
   * @example
   * const props = new DocProperties();
   * props.setAuthor('Jane Doe').setTitle('Report');
   * workbook.setProperties(props);
   */
  setProperties(properties: DocProperties): void;
}

// Re-export native bindings for advanced use
export { RuscFormat, RuscNote, RuscFormula, RuscDataValidation, RuscChart, RuscWorkbook, RuscWorksheet, RuscProtectionOptions, RuscRichText, RuscImage, RuscTable, RuscTableColumn, RuscSparkline, RuscDocProperties, RuscButton, RuscShape } from './index';
