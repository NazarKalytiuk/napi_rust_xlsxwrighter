/**
 * Tests for Worksheet functionality
 */

import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { Format, Workbook } from '../wrapper.ts';
import {
  cleanupTestOutput,
  createTestOutputDir,
  readCellStyle,
  readXlsxPart,
} from './test-helpers.mjs';

describe('Worksheet', () => {
  let testOutputDir;

  beforeAll(() => {
    testOutputDir = createTestOutputDir();
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
      worksheet.write(1, 0, Math.PI);
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

function bold() {
  return new Format().setBold();
}

describe('Worksheet formatted writes in XLSX output', () => {
  test('writeRow applies the format to every value', () => {
    const workbook = new Workbook();
    workbook.addWorksheet().writeRow(0, 0, ['text', 1, true], bold());

    const xlsx = workbook.saveToBuffer();
    for (const cell of ['A1', 'B1', 'C1']) {
      expect(readCellStyle(xlsx, cell).font).toContain('<b/>');
    }
  });

  test('writeRows applies the format to every value of every row', () => {
    const workbook = new Workbook();
    workbook.addWorksheet().writeRows(
      0,
      0,
      [
        ['a', 1],
        [false, 'b'],
      ],
      bold(),
    );

    const xlsx = workbook.saveToBuffer();
    for (const cell of ['A1', 'B1', 'A2', 'B2']) {
      expect(readCellStyle(xlsx, cell).font).toContain('<b/>');
    }
  });

  test.each([
    ['mergeRange', (worksheet, format) => worksheet.mergeRange(1, 1, 2, 3, 'Merged', format)],
    [
      'mergeRangeWithFormat',
      (worksheet, format) => worksheet.mergeRangeWithFormat(1, 1, 2, 3, 'Merged', format),
    ],
  ])('%s merges the range and formats all of its cells', (_method, merge) => {
    const workbook = new Workbook();
    merge(workbook.addWorksheet(), bold());

    const xlsx = workbook.saveToBuffer();
    expect(readXlsxPart(xlsx, 'xl/worksheets/sheet1.xml')).toContain(
      '<mergeCells count="1"><mergeCell ref="B2:D3"/></mergeCells>',
    );
    for (const cell of ['B2', 'C2', 'D2', 'B3', 'C3', 'D3']) {
      expect(readCellStyle(xlsx, cell).font).toContain('<b/>');
    }
  });

  test('mergeRange without a format leaves the cells unstyled', () => {
    const workbook = new Workbook();
    workbook.addWorksheet().mergeRange(0, 0, 0, 1, 'Merged');

    const xlsx = workbook.saveToBuffer();
    expect(readXlsxPart(xlsx, 'xl/worksheets/sheet1.xml')).toContain('<mergeCell ref="A1:B1"/>');
    expect(readCellStyle(xlsx, 'A1').font).not.toContain('<b/>');
  });
});

describe('Worksheet.write with missing and unsupported values', () => {
  test.each([null, undefined])('writes no cell for %s without a format', (value) => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    worksheet.write(0, 0, 'kept');
    worksheet.write(0, 1, value);
    worksheet.write(1, 0, value);

    const sheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain(
      '<sheetData><row r="1" spans="1:1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
    );
  });

  test.each([
    ['write', (worksheet, value, format) => worksheet.write(0, 0, value, format)],
    [
      'writeWithFormat',
      (worksheet, value, format) => worksheet.writeWithFormat(0, 0, value, format),
    ],
  ])('%s writes a formatted blank cell for null', (_method, write) => {
    const workbook = new Workbook();
    write(workbook.addWorksheet(), null, bold());

    const xlsx = workbook.saveToBuffer();
    expect(readXlsxPart(xlsx, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1" s="1"/>');
    expect(readCellStyle(xlsx, 'A1').font).toContain('<b/>');
  });

  test.each([
    [{}, 'Unsupported cell value type: object. Use a string, number, boolean or null'],
    [['a'], 'Unsupported cell value type: array. Use a string, number, boolean or null'],
    [10n, 'Unsupported cell value type: bigint. Use a string, number, boolean or null'],
    [Symbol('x'), 'Unsupported cell value type: symbol. Use a string, number, boolean or null'],
    [() => 1, 'Unsupported cell value type: function. Use a string, number, boolean or null'],
    [new Date(0), 'Unsupported cell value: Date. Use writeDatetime() for dates'],
  ])('rejects %o with or without a format', (value, message) => {
    const worksheet = new Workbook().addWorksheet();
    expect(() => worksheet.write(0, 0, value)).toThrow(message);
    expect(() => worksheet.write(0, 0, value, bold())).toThrow(message);
    expect(() => worksheet.writeWithFormat(0, 0, value, bold())).toThrow(message);
  });
});

describe('Worksheet.writeRow and writeRows with missing and unsupported values', () => {
  test('leave cells for null, undefined and holes empty without a format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    // oxlint-disable-next-line no-sparse-arrays -- the hole is the value under test
    worksheet.writeRow(0, 0, ['a', null, undefined, , 'e']);
    worksheet.writeRows(1, 0, [[null, 'b']]);

    const sheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
    const cells = [...sheet.matchAll(/<c r="([A-Z]+\d+)"/g)].map(([, cell]) => cell);
    expect(cells).toEqual(['A1', 'E1', 'B2']);
  });

  test('write formatted blank cells for null, undefined and holes with a format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    // oxlint-disable-next-line no-sparse-arrays -- the hole is the value under test
    worksheet.writeRow(0, 0, ['a', null, undefined, , 'e'], bold());
    worksheet.writeRows(1, 0, [[null, 'b']], bold());

    const xlsx = workbook.saveToBuffer();
    const sheet = readXlsxPart(xlsx, 'xl/worksheets/sheet1.xml');
    for (const cell of ['B1', 'C1', 'D1', 'A2']) {
      expect(sheet).toContain(`<c r="${cell}" s="1"/>`);
      expect(readCellStyle(xlsx, cell).font).toContain('<b/>');
    }
  });

  test.each([
    ['writeRow', (worksheet, value) => worksheet.writeRow(3, 1, ['ok', value])],
    ['writeRows', (worksheet, value) => worksheet.writeRows(2, 1, [['ok'], ['ok', value]])],
  ])('%s rejects unsupported values with their cell position', (_method, write) => {
    const worksheet = new Workbook().addWorksheet();
    expect(() => write(worksheet, { amount: 1 })).toThrow(
      'Invalid value at row 3, col 2: Unsupported cell value type: object. Use a string, number, boolean or null',
    );
    expect(() => write(worksheet, new Date(0))).toThrow(
      'Invalid value at row 3, col 2: Unsupported cell value: Date. Use writeDatetime() for dates',
    );
  });

  test.each([null, 'row', { 0: 'a' }])('writeRows rejects the row %o', (row) => {
    const worksheet = new Workbook().addWorksheet();
    expect(() => worksheet.writeRows(0, 0, [['a'], row])).toThrow(
      'Invalid row at index 1: expected an array of cell values',
    );
  });

  test.each([
    ['writeRow', (worksheet) => worksheet.writeRow(0, 0, 'a,b')],
    ['writeRows', (worksheet) => worksheet.writeRows(0, 0, 'a,b')],
  ])('%s rejects values that are not an array', (_method, write) => {
    expect(() => write(new Workbook().addWorksheet())).toThrow('Expected an array');
  });
});
