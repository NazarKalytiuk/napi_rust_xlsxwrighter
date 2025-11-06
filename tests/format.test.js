/**
 * Tests for Format functionality
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Format', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  test('should create a new format', () => {
    const format = new Format();
    expect(format).toBeInstanceOf(Format);
    expect(format._native).toBeDefined();
  });

  test('should set bold', () => {
    const format = new Format();
    const result = format.setBold();
    expect(result).toBe(format); // Check chaining
  });

  test('should set italic', () => {
    const format = new Format();
    const result = format.setItalic();
    expect(result).toBe(format);
  });

  test('should set font size', () => {
    const format = new Format();
    const result = format.setFontSize(14);
    expect(result).toBe(format);
  });

  test('should set font color with named color', () => {
    const format = new Format();
    expect(() => {
      format.setFontColor('red');
    }).not.toThrow();
  });

  test('should set font color with hex color', () => {
    const format = new Format();
    expect(() => {
      format.setFontColor('#FF0000');
    }).not.toThrow();
  });

  test('should set background color', () => {
    const format = new Format();
    expect(() => {
      format.setBackgroundColor('yellow');
    }).not.toThrow();
  });

  test('should set horizontal alignment', () => {
    const format = new Format();
    expect(() => {
      format.setAlign('center');
    }).not.toThrow();
  });

  test('should set vertical alignment', () => {
    const format = new Format();
    expect(() => {
      format.setVerticalAlign('center');
    }).not.toThrow();
  });

  test('should set border', () => {
    const format = new Format();
    expect(() => {
      format.setBorder('thin');
    }).not.toThrow();
  });

  test('should set number format', () => {
    const format = new Format();
    expect(() => {
      format.setNumFormat('#,##0.00');
    }).not.toThrow();
  });

  test('should chain multiple format operations', () => {
    const format = new Format()
      .setBold()
      .setItalic()
      .setFontSize(16)
      .setFontColor('blue')
      .setBackgroundColor('yellow')
      .setAlign('center')
      .setBorder('thin');

    expect(format).toBeInstanceOf(Format);
  });

  test('should write data with format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    const format = new Format()
      .setBold()
      .setFontColor('red');

    expect(() => {
      worksheet.write(0, 0, 'Formatted Text', format);
      worksheet.write(1, 0, 42, format);
      worksheet.write(2, 0, true, format);
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-formatted.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });

  test('should handle invalid color', () => {
    const format = new Format();
    expect(() => {
      format.setFontColor('invalidcolor123');
    }).toThrow();
  });

  test('should handle invalid alignment', () => {
    const format = new Format();
    expect(() => {
      format.setAlign('invalidalign');
    }).toThrow();
  });

  test('should apply number formatting', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    const currencyFormat = new Format().setNumFormat('$#,##0.00');
    const percentFormat = new Format().setNumFormat('0.00%');

    expect(() => {
      worksheet.write(0, 0, 1234.56, currencyFormat);
      worksheet.write(1, 0, 0.85, percentFormat);
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-number-format.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });
});
