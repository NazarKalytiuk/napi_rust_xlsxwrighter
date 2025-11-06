/**
 * Tests for Worksheet functionality
 */

const fs = require('fs');
const path = require('path');
const { Workbook } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Worksheet', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  test('should write string values', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.write(0, 0, 'Hello');
      worksheet.writeString(1, 0, 'World');
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-strings.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });

  test('should write numeric values', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.write(0, 0, 42);
      worksheet.write(1, 0, 3.14159);
      worksheet.writeNumber(2, 0, -99.99);
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-numbers.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });

  test('should write boolean values', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.write(0, 0, true);
      worksheet.write(1, 0, false);
      worksheet.writeBoolean(2, 0, true);
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-booleans.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });

  test('should set column width', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.setColumnWidth(0, 20);
      worksheet.setColumnWidth(1, 15.5);
    }).not.toThrow();
  });

  test('should set row height', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.setRowHeight(0, 30);
      worksheet.setRowHeight(1, 25.5);
    }).not.toThrow();
  });

  test('should set worksheet name', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.setName('CustomName');
    }).not.toThrow();
  });

  test('should merge cells', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => {
      worksheet.mergeRange(0, 0, 0, 2, 'Merged Cell');
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-merge.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });

  test('should handle large datasets', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    // Write 100 rows x 10 columns
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 10; col++) {
        worksheet.write(row, col, `R${row}C${col}`);
      }
    }

    const filename = path.join(testOutputDir, 'test-large.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
    expect(fs.statSync(filename).size).toBeGreaterThan(1000);
  });
});
