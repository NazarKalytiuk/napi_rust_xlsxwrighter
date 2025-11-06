/**
 * Tests for Hyperlinks, Autofilter, and Freeze Panes
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Hyperlinks, Autofilter, and Freeze Panes', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Hyperlinks', () => {
    test('should write external URL', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('URLs');

      worksheet.writeUrl(0, 0, 'https://www.example.com');

      const outputPath = path.join(testOutputDir, 'url-basic.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write URL with custom text', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('URLs with Text');

      worksheet.writeUrlWithText(0, 0, 'https://www.rust-lang.org', 'Learn Rust');

      const outputPath = path.join(testOutputDir, 'url-text.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write URL with format', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Formatted URLs');

      const format = new Format()
        .setBold()
        .setFontColor('red');

      worksheet.writeUrl(0, 0, 'https://www.example.com', format);

      const outputPath = path.join(testOutputDir, 'url-formatted.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write URL with text and format', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Formatted URLs with Text');

      const format = new Format()
        .setItalic()
        .setFontColor('blue')
        .setUnderline('single');

      worksheet.writeUrlWithText(0, 0, 'https://github.com', 'Visit GitHub', format);

      const outputPath = path.join(testOutputDir, 'url-text-formatted.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write multiple different URLs', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Multiple URLs');

      // External URLs
      worksheet.writeUrlWithText(0, 0, 'https://www.google.com', 'Google');
      worksheet.writeUrlWithText(1, 0, 'https://www.github.com', 'GitHub');
      worksheet.writeUrlWithText(2, 0, 'https://www.stackoverflow.com', 'Stack Overflow');

      // Email link
      worksheet.writeUrlWithText(4, 0, 'mailto:test@example.com', 'Send Email');

      const outputPath = path.join(testOutputDir, 'url-multiple.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write internal link to another sheet', () => {
      const workbook = new Workbook();
      const worksheet1 = workbook.addWorksheet('Sheet1');
      const worksheet2 = workbook.addWorksheet('Sheet2');

      worksheet2.write(5, 5, 'Target Cell');

      // Link from Sheet1 to Sheet2
      worksheet1.writeUrlWithText(0, 0, 'internal:Sheet2!F6', 'Go to Sheet2');

      const outputPath = path.join(testOutputDir, 'url-internal.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should write https URL', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('HTTPS');

      worksheet.writeUrl(0, 0, 'https://secure.example.com');

      const outputPath = path.join(testOutputDir, 'url-https.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Autofilter', () => {
    test('should add autofilter to range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Autofilter');

      // Write headers
      worksheet.write(0, 0, 'Name');
      worksheet.write(0, 1, 'Age');
      worksheet.write(0, 2, 'City');

      // Write data
      worksheet.write(1, 0, 'John');
      worksheet.write(1, 1, 30);
      worksheet.write(1, 2, 'New York');

      worksheet.write(2, 0, 'Jane');
      worksheet.write(2, 1, 25);
      worksheet.write(2, 2, 'Los Angeles');

      // Add autofilter
      worksheet.autofilter(0, 0, 2, 2);

      const outputPath = path.join(testOutputDir, 'autofilter-basic.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should add autofilter to large dataset', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Large Data');

      const headerFormat = new Format().setBold().setBackgroundColor('#D3D3D3');

      // Headers
      worksheet.write(0, 0, 'ID', headerFormat);
      worksheet.write(0, 1, 'Name', headerFormat);
      worksheet.write(0, 2, 'Department', headerFormat);
      worksheet.write(0, 3, 'Salary', headerFormat);

      // Generate data
      for (let i = 0; i < 100; i++) {
        worksheet.write(i + 1, 0, i + 1);
        worksheet.write(i + 1, 1, `Employee ${i + 1}`);
        worksheet.write(i + 1, 2, ['Sales', 'Engineering', 'Marketing'][i % 3]);
        worksheet.write(i + 1, 3, 50000 + (i * 1000));
      }

      // Add autofilter to entire range
      worksheet.autofilter(0, 0, 100, 3);

      const outputPath = path.join(testOutputDir, 'autofilter-large.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should add autofilter with frozen header', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Filter + Freeze');

      const headerFormat = new Format().setBold();

      // Headers
      worksheet.write(0, 0, 'Product', headerFormat);
      worksheet.write(0, 1, 'Price', headerFormat);
      worksheet.write(0, 2, 'Stock', headerFormat);

      // Data
      for (let i = 0; i < 20; i++) {
        worksheet.write(i + 1, 0, `Product ${i + 1}`);
        worksheet.write(i + 1, 1, 10 + i);
        worksheet.write(i + 1, 2, 100 + (i * 10));
      }

      // Freeze header row
      worksheet.setFreezePanes(1, 0);

      // Add autofilter
      worksheet.autofilter(0, 0, 20, 2);

      const outputPath = path.join(testOutputDir, 'autofilter-freeze.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Freeze Panes', () => {
    test('should freeze top row', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Freeze Row');

      const headerFormat = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');

      // Header
      worksheet.write(0, 0, 'Name', headerFormat);
      worksheet.write(0, 1, 'Value', headerFormat);

      // Data
      for (let i = 1; i <= 50; i++) {
        worksheet.write(i, 0, `Item ${i}`);
        worksheet.write(i, 1, i * 10);
      }

      // Freeze first row
      worksheet.setFreezePanes(1, 0);

      const outputPath = path.join(testOutputDir, 'freeze-row.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should freeze first column', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Freeze Column');

      const labelFormat = new Format().setBold();

      // Labels in first column
      for (let i = 0; i < 20; i++) {
        worksheet.write(i, 0, `Row ${i + 1}`, labelFormat);
      }

      // Data in other columns
      for (let col = 1; col <= 10; col++) {
        for (let row = 0; row < 20; row++) {
          worksheet.write(row, col, `C${col}R${row + 1}`);
        }
      }

      // Freeze first column
      worksheet.setFreezePanes(0, 1);

      const outputPath = path.join(testOutputDir, 'freeze-column.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should freeze both row and column', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Freeze Both');

      const headerFormat = new Format().setBold().setBackgroundColor('#D3D3D3');

      // Top-left corner
      worksheet.write(0, 0, 'ID', headerFormat);

      // Column headers
      for (let col = 1; col <= 10; col++) {
        worksheet.write(0, col, `Col ${col}`, headerFormat);
      }

      // Row headers
      for (let row = 1; row <= 30; row++) {
        worksheet.write(row, 0, `Row ${row}`, headerFormat);
      }

      // Data
      for (let row = 1; row <= 30; row++) {
        for (let col = 1; col <= 10; col++) {
          worksheet.write(row, col, row * col);
        }
      }

      // Freeze first row and first column
      worksheet.setFreezePanes(1, 1);

      const outputPath = path.join(testOutputDir, 'freeze-both.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should freeze multiple rows', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Freeze Multiple Rows');

      const titleFormat = new Format().setBold().setFontSize(16);
      const headerFormat = new Format().setBold().setBackgroundColor('#D3D3D3');

      // Title row
      worksheet.write(0, 0, 'Sales Report 2024', titleFormat);

      // Sub-header
      worksheet.write(1, 0, 'Q1 Results', new Format().setItalic());

      // Column headers
      worksheet.write(2, 0, 'Product', headerFormat);
      worksheet.write(2, 1, 'Sales', headerFormat);
      worksheet.write(2, 2, 'Revenue', headerFormat);

      // Data
      for (let i = 0; i < 50; i++) {
        worksheet.write(i + 3, 0, `Product ${i + 1}`);
        worksheet.write(i + 3, 1, 100 + i);
        worksheet.write(i + 3, 2, (100 + i) * 50);
      }

      // Freeze first 3 rows (title, sub-header, headers)
      worksheet.setFreezePanes(3, 0);

      const outputPath = path.join(testOutputDir, 'freeze-multiple-rows.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should not freeze when row and col are 0', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('No Freeze');

      worksheet.write(0, 0, 'No freeze panes');
      worksheet.setFreezePanes(0, 0);

      const outputPath = path.join(testOutputDir, 'freeze-none.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Combined Features', () => {
    test('should combine autofilter, freeze panes, and URLs', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Combined');

      const headerFormat = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');

      // Headers
      worksheet.write(0, 0, 'Name', headerFormat);
      worksheet.write(0, 1, 'Email', headerFormat);
      worksheet.write(0, 2, 'Website', headerFormat);
      worksheet.write(0, 3, 'Status', headerFormat);

      // Data
      const people = [
        ['John Doe', 'john@example.com', 'https://john.example.com', 'Active'],
        ['Jane Smith', 'jane@example.com', 'https://jane.example.com', 'Active'],
        ['Bob Johnson', 'bob@example.com', 'https://bob.example.com', 'Inactive'],
      ];

      people.forEach((person, i) => {
        const row = i + 1;
        worksheet.write(row, 0, person[0]);
        worksheet.writeUrlWithText(row, 1, `mailto:${person[1]}`, person[1]);
        worksheet.writeUrl(row, 2, person[2]);
        worksheet.write(row, 3, person[3]);
      });

      // Freeze header row
      worksheet.setFreezePanes(1, 0);

      // Add autofilter
      worksheet.autofilter(0, 0, people.length, 3);

      const outputPath = path.join(testOutputDir, 'combined-features.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Save to Buffer', () => {
    test('should save hyperlinks to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('URLs Buffer');

      worksheet.writeUrl(0, 0, 'https://example.com');

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });

    test('should save autofilter to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Autofilter Buffer');

      worksheet.write(0, 0, 'Header');
      worksheet.write(1, 0, 'Data');
      worksheet.autofilter(0, 0, 1, 0);

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });

    test('should save freeze panes to buffer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Freeze Buffer');

      worksheet.write(0, 0, 'Header');
      worksheet.setFreezePanes(1, 0);

      const buffer = workbook.saveToBuffer();
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });
});
