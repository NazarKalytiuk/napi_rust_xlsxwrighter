/**
 * Tests for constant memory and low memory modes
 */

const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const { Workbook } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Memory Modes', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Constant Memory Mode', () => {
    test('should create worksheet with constant memory mode', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory('ConstantMem');

      expect(worksheet).toBeDefined();
      expect(worksheet._native).toBeDefined();
    });

    test('should create worksheet with constant memory mode and name', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory('CustomName');

      expect(worksheet).toBeDefined();
    });

    test('should write data sequentially in constant memory mode', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory();

      // Write data sequentially row by row
      expect(() => {
        for (let row = 0; row < 1000; row++) {
          worksheet.write(row, 0, `Row ${row}`);
          worksheet.write(row, 1, row * 100);
          worksheet.write(row, 2, row % 2 === 0);
        }
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'test-constant-memory.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should write mixed rows and preserve null column positions', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory();

      worksheet.writeRow(0, 0, ['Alice', 30, true, null, 'ready']);
      worksheet.writeRows(1, 0, [
        ['Bob', 25, false, null, 'done'],
      ]);

      const archive = new AdmZip(workbook.saveToBuffer());
      const xml = archive.readAsText('xl/worksheets/sheet1.xml');
      const cell = address => xml.match(
        new RegExp(`<c r="${address}"[^>]*>[\\s\\S]*?</c>`),
      )?.[0];
      const numericCellValue = address => {
        const value = cell(address)?.match(/<v>([^<]+)<\/v>/)?.[1];
        return value === undefined ? undefined : Number(value);
      };

      expect(cell('A1')).toContain('<t>Alice</t>');
      expect(numericCellValue('B1')).toBe(30);
      expect(numericCellValue('C1')).toBe(1);
      expect(cell('D1')).toBeUndefined();
      expect(cell('E2')).toContain('<t>done</t>');
    });

    test('should handle large dataset in constant memory mode', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory('Large');

      // Write 10,000 rows
      expect(() => {
        for (let row = 0; row < 10000; row++) {
          worksheet.write(row, 0, row);
          worksheet.write(row, 1, `Data ${row}`);
        }
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'test-constant-memory-large.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
      expect(fs.statSync(filename).size).toBeGreaterThan(10000);
    });

    test('should save constant memory worksheet to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory();

      for (let row = 0; row < 100; row++) {
        worksheet.write(row, 0, `Row ${row}`);
      }

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });

  describe('Low Memory Mode', () => {
    test('should create worksheet with low memory mode', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithLowMemory('LowMem');

      expect(worksheet).toBeDefined();
      expect(worksheet._native).toBeDefined();
    });

    test('should create worksheet with low memory mode and name', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithLowMemory('CustomName');

      expect(worksheet).toBeDefined();
    });

    test('should write data sequentially in low memory mode', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithLowMemory();

      // Write data sequentially row by row
      expect(() => {
        for (let row = 0; row < 1000; row++) {
          worksheet.write(row, 0, `Row ${row}`);
          worksheet.write(row, 1, row * 100);
          worksheet.write(row, 2, row % 2 === 0);
        }
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'test-low-memory.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should efficiently handle repeated strings in low memory mode', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithLowMemory('Repeated');

      const categories = ['A', 'B', 'C', 'D', 'E'];
      const statuses = ['Active', 'Inactive', 'Pending'];

      // Write data with lots of repeated strings
      expect(() => {
        for (let row = 0; row < 5000; row++) {
          worksheet.write(row, 0, row);
          worksheet.write(row, 1, categories[row % categories.length]);
          worksheet.write(row, 2, statuses[row % statuses.length]);
        }
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'test-low-memory-repeated.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);

      // File size should be reasonable due to shared string table
      const fileSize = fs.statSync(filename).size;
      expect(fileSize).toBeGreaterThan(5000);
      expect(fileSize).toBeLessThan(100000); // Should be well under 100KB for 5000 rows
    });

    test('should save low memory worksheet to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithLowMemory();

      for (let row = 0; row < 100; row++) {
        worksheet.write(row, 0, `Row ${row}`);
      }

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });

  describe('setTempdir', () => {
    test('should set custom temp directory', () => {
      const workbook = new Workbook();
      const tempDir = testOutputDir;

      // Must be called before adding worksheets
      expect(() => {
        workbook.setTempdir(tempDir);
      }).not.toThrow();

      // Add worksheet after setting tempdir
      const worksheet = workbook.addWorksheetWithConstantMemory();
      worksheet.write(0, 0, 'Test');

      const filename = path.join(testOutputDir, 'test-tempdir.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle invalid temp directory path', () => {
      const workbook = new Workbook();

      expect(() => {
        workbook.setTempdir('/invalid/nonexistent/path/that/does/not/exist');
      }).toThrow();
    });
  });

  describe('Mixed Modes', () => {
    test('should support mixing normal and memory mode worksheets', () => {
      const workbook = new Workbook();

      const normalSheet = workbook.addWorksheet('Normal');
      const constantSheet = workbook.addWorksheetWithConstantMemory('Constant');
      const lowSheet = workbook.addWorksheetWithLowMemory('Low');

      // Write to each sheet
      normalSheet.write(0, 0, 'Normal mode');
      constantSheet.write(0, 0, 'Constant memory mode');
      lowSheet.write(0, 0, 'Low memory mode');

      const filename = path.join(testOutputDir, 'test-mixed-modes.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Performance Characteristics', () => {
    test('constant memory mode should handle very large datasets', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithConstantMemory('VeryLarge');

      const startTime = Date.now();
      const rowCount = 20000;

      for (let row = 0; row < rowCount; row++) {
        worksheet.write(row, 0, row);
        worksheet.write(row, 1, `Data-${row}`);
        worksheet.write(row, 2, row * 3.14);
      }

      const writeTime = Date.now() - startTime;

      const filename = path.join(testOutputDir, 'test-very-large-constant.xlsx');
      workbook.save(filename);

      expect(fs.existsSync(filename)).toBe(true);
      console.log(`      Wrote ${rowCount} rows in constant memory mode: ${writeTime}ms`);
    });

    test('low memory mode should handle large datasets with repeated data', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheetWithLowMemory('VeryLarge');

      const categories = ['Cat1', 'Cat2', 'Cat3', 'Cat4', 'Cat5'];
      const startTime = Date.now();
      const rowCount = 20000;

      for (let row = 0; row < rowCount; row++) {
        worksheet.write(row, 0, row);
        worksheet.write(row, 1, categories[row % categories.length]);
        worksheet.write(row, 2, row * 3.14);
      }

      const writeTime = Date.now() - startTime;

      const filename = path.join(testOutputDir, 'test-very-large-low.xlsx');
      workbook.save(filename);

      expect(fs.existsSync(filename)).toBe(true);
      console.log(`      Wrote ${rowCount} rows in low memory mode: ${writeTime}ms`);
    });
  });
});
