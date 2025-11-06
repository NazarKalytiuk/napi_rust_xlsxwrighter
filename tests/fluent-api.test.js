/**
 * Tests for Fluent API (Method Chaining) vs Individual Method Calls
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format, Note } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Fluent API and Method Chaining', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Format - Individual Method Calls', () => {
    test('should support calling methods individually', () => {
      const format = new Format();

      // Call each method separately
      const result1 = format.setBold();
      expect(result1).toBe(format); // Returns this

      const result2 = format.setItalic();
      expect(result2).toBe(format); // Returns this

      const result3 = format.setFontColor('red');
      expect(result3).toBe(format); // Returns this

      const result4 = format.setAlign('center');
      expect(result4).toBe(format); // Returns this

      // All methods should return the same instance
      expect(result1).toBe(result2);
      expect(result2).toBe(result3);
      expect(result3).toBe(result4);
    });

    test('should create workbook with individual method calls', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Individual');

      const format = new Format();
      format.setBold();
      format.setItalic();
      format.setFontSize(14);
      format.setFontColor('#4472C4');
      format.setBackgroundColor('yellow');
      format.setAlign('center');
      format.setBorder('thin');

      worksheet.write(0, 0, 'Individual Calls', format);

      const outputPath = path.join(testOutputDir, 'individual-calls.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Format - Fluent API (Method Chaining)', () => {
    test('should support method chaining', () => {
      const format = new Format()
        .setBold()
        .setItalic()
        .setFontColor('red')
        .setAlign('center');

      expect(format).toBeInstanceOf(Format);
    });

    test('should support long chains', () => {
      const format = new Format()
        .setBold()
        .setItalic()
        .setFontSize(16)
        .setFontName('Arial')
        .setFontColor('#FF0000')
        .setUnderline('double')
        .setStrikethrough()
        .setBackgroundColor('#FFFF00')
        .setForegroundColor('#0000FF')
        .setPattern('mediumGray')
        .setAlign('center')
        .setVerticalAlign('center')
        .setTextWrap()
        .setIndent(2)
        .setRotation(45)
        .setShrink()
        .setBorder('medium')
        .setBorderTop('thick')
        .setBorderBottom('thin')
        .setBorderLeft('dashed')
        .setBorderRight('dotted')
        .setBorderColor('#000000')
        .setBorderTopColor('#FF0000')
        .setBorderBottomColor('#00FF00')
        .setBorderLeftColor('#0000FF')
        .setBorderRightColor('#FFFF00')
        .setNumFormat('#,##0.00')
        .setUnlocked()
        .setHidden();

      expect(format).toBeInstanceOf(Format);
    });

    test('should create workbook with method chaining', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Chained');

      const format = new Format()
        .setBold()
        .setItalic()
        .setFontSize(14)
        .setFontColor('#4472C4')
        .setBackgroundColor('yellow')
        .setAlign('center')
        .setBorder('thin');

      worksheet.write(0, 0, 'Method Chaining', format);

      const outputPath = path.join(testOutputDir, 'method-chaining.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Format - Mixed Approach', () => {
    test('should support mixing individual calls and chaining', () => {
      const format = new Format();

      // Start with individual calls
      format.setBold();
      format.setItalic();

      // Continue with chaining
      format
        .setFontSize(14)
        .setFontColor('red')
        .setAlign('center');

      // Back to individual call
      format.setBorder('thin');

      expect(format).toBeInstanceOf(Format);
    });

    test('should support creating base format and extending it', () => {
      // Create base format with chaining
      const baseFormat = new Format()
        .setFontSize(12)
        .setFontColor('#000000');

      // Clone and extend (simulated by creating new format)
      const headerFormat = new Format()
        .setFontSize(12)
        .setFontColor('#000000')
        .setBold()
        .setBackgroundColor('#D3D3D3');

      expect(baseFormat).toBeInstanceOf(Format);
      expect(headerFormat).toBeInstanceOf(Format);
    });
  });

  describe('Note - Individual Method Calls', () => {
    test('should support calling methods individually', () => {
      const note = new Note('Test note');

      const result1 = note.setAuthor('Author');
      expect(result1).toBe(note); // Returns this

      const result2 = note.setWidth(200);
      expect(result2).toBe(note); // Returns this

      const result3 = note.setHeight(100);
      expect(result3).toBe(note); // Returns this

      const result4 = note.setBackgroundColor('#FFFFE1');
      expect(result4).toBe(note); // Returns this

      // All methods should return the same instance
      expect(result1).toBe(result2);
      expect(result2).toBe(result3);
      expect(result3).toBe(result4);
    });

    test('should create note with individual method calls', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Notes Individual');

      const note = new Note('Individual note');
      note.setAuthor('John Doe');
      note.setWidth(200);
      note.setHeight(100);
      note.setBackgroundColor('#FFFFE1');
      note.setVisible(true);

      worksheet.write(0, 0, 'Cell with note');
      worksheet.insertNote(0, 0, note);

      const outputPath = path.join(testOutputDir, 'notes-individual.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Note - Fluent API (Method Chaining)', () => {
    test('should support method chaining', () => {
      const note = new Note('Chained note')
        .setAuthor('Jane Smith')
        .setWidth(200)
        .setHeight(100);

      expect(note).toBeInstanceOf(Note);
    });

    test('should support full chain', () => {
      const note = new Note('Fully chained note')
        .setAuthor('Full Name')
        .setWidth(250)
        .setHeight(120)
        .setBackgroundColor('#FFB6C1')
        .setVisible(true)
        .setAltText('Accessibility text');

      expect(note).toBeInstanceOf(Note);
    });

    test('should create note with method chaining', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Notes Chained');

      const note = new Note('Chained note')
        .setAuthor('Jane Smith')
        .setWidth(250)
        .setHeight(100)
        .setBackgroundColor('#FFE4B5')
        .setVisible(true)
        .setAltText('Important note');

      worksheet.write(0, 0, 'Cell with chained note');
      worksheet.insertNote(0, 0, note);

      const outputPath = path.join(testOutputDir, 'notes-chained.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });

  describe('Comparison Tests', () => {
    test('individual calls and chaining should produce same result', () => {
      const workbook1 = new Workbook();
      const worksheet1 = workbook1.addWorksheet('Individual');

      // Individual calls
      const format1 = new Format();
      format1.setBold();
      format1.setItalic();
      format1.setFontSize(14);
      format1.setFontColor('red');
      format1.setBackgroundColor('yellow');
      format1.setAlign('center');

      worksheet1.write(0, 0, 'Test', format1);
      const buffer1 = workbook1.saveToBuffer();

      const workbook2 = new Workbook();
      const worksheet2 = workbook2.addWorksheet('Individual');

      // Method chaining
      const format2 = new Format()
        .setBold()
        .setItalic()
        .setFontSize(14)
        .setFontColor('red')
        .setBackgroundColor('yellow')
        .setAlign('center');

      worksheet2.write(0, 0, 'Test', format2);
      const buffer2 = workbook2.saveToBuffer();

      // Both should produce buffers of similar size
      expect(buffer1.length).toBeGreaterThan(0);
      expect(buffer2.length).toBeGreaterThan(0);
      expect(Math.abs(buffer1.length - buffer2.length)).toBeLessThan(10);
    });

    test('note individual calls and chaining should produce same result', () => {
      const workbook1 = new Workbook();
      const worksheet1 = workbook1.addWorksheet('Sheet');

      // Individual calls
      const note1 = new Note('Test');
      note1.setAuthor('Author');
      note1.setWidth(200);
      note1.setHeight(100);

      worksheet1.write(0, 0, 'Test');
      worksheet1.insertNote(0, 0, note1);
      const buffer1 = workbook1.saveToBuffer();

      const workbook2 = new Workbook();
      const worksheet2 = workbook2.addWorksheet('Sheet');

      // Method chaining
      const note2 = new Note('Test')
        .setAuthor('Author')
        .setWidth(200)
        .setHeight(100);

      worksheet2.write(0, 0, 'Test');
      worksheet2.insertNote(0, 0, note2);
      const buffer2 = workbook2.saveToBuffer();

      // Both should produce buffers of similar size
      expect(buffer1.length).toBeGreaterThan(0);
      expect(buffer2.length).toBeGreaterThan(0);
      expect(Math.abs(buffer1.length - buffer2.length)).toBeLessThan(10);
    });
  });

  describe('Real-world Patterns', () => {
    test('should support builder pattern', () => {
      // Helper function that builds a format
      function createHeaderFormat() {
        return new Format()
          .setBold()
          .setFontSize(14)
          .setBackgroundColor('#D3D3D3')
          .setAlign('center');
      }

      function createDataFormat() {
        return new Format()
          .setNumFormat('#,##0.00');
      }

      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Builder Pattern');

      worksheet.write(0, 0, 'Header', createHeaderFormat());
      worksheet.write(1, 0, 1234.56, createDataFormat());

      const outputPath = path.join(testOutputDir, 'builder-pattern.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });

    test('should support configuration object pattern', () => {
      function configureFormat(config) {
        const format = new Format();

        if (config.bold) format.setBold();
        if (config.italic) format.setItalic();
        if (config.fontSize) format.setFontSize(config.fontSize);
        if (config.color) format.setFontColor(config.color);
        if (config.align) format.setAlign(config.align);

        return format;
      }

      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Config Pattern');

      const format = configureFormat({
        bold: true,
        italic: true,
        fontSize: 14,
        color: 'red',
        align: 'center'
      });

      worksheet.write(0, 0, 'Configured', format);

      const outputPath = path.join(testOutputDir, 'config-pattern.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
    });
  });
});
