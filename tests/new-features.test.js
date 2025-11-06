const { Workbook, Format } = require('../wrapper.js');
const fs = require('fs');
const path = require('path');

describe('New Features v0.8.0', () => {
  let workbook;
  let worksheet;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('Test');
  });

  describe('writeBlank', () => {
    it('should write blank cell with format', () => {
      const format = new Format().setBackgroundColor('yellow');
      expect(() => {
        worksheet.writeBlank(0, 0, format);
      }).not.toThrow();
    });

    it('should create valid Excel file with blank cells', () => {
      const format = new Format().setBackgroundColor('gray').setBorder('thin');
      worksheet.writeBlank(0, 0, format);
      worksheet.writeBlank(0, 1, format);
      worksheet.writeBlank(0, 2, format);

      const filename = path.join(__dirname, 'temp-blank-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Worksheet Tab Color', () => {
    it('should set tab color with named color', () => {
      expect(() => worksheet.setTabColor('red')).not.toThrow();
    });

    it('should set tab color with hex code', () => {
      expect(() => worksheet.setTabColor('#FF0000')).not.toThrow();
    });

    it('should create file with colored tab', () => {
      worksheet.setTabColor('blue');
      worksheet.write(0, 0, 'Blue Tab');

      const filename = path.join(__dirname, 'temp-tab-color-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Hide/Show Worksheet', () => {
    it('should hide worksheet', () => {
      expect(() => worksheet.setHidden(true)).not.toThrow();
    });

    it('should show worksheet', () => {
      expect(() => worksheet.setHidden(false)).not.toThrow();
    });

    it('should create file with hidden sheet', () => {
      const sheet1 = workbook.addWorksheet('Visible');
      const sheet2 = workbook.addWorksheet('Hidden');

      sheet1.write(0, 0, 'Visible Sheet');
      sheet2.write(0, 0, 'Hidden Sheet');
      sheet2.setHidden(true);

      const filename = path.join(__dirname, 'temp-hidden-sheet-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Worksheet Zoom', () => {
    it('should set zoom level', () => {
      expect(() => worksheet.setZoom(150)).not.toThrow();
    });

    it('should accept minimum zoom (10)', () => {
      expect(() => worksheet.setZoom(10)).not.toThrow();
    });

    it('should accept maximum zoom (400)', () => {
      expect(() => worksheet.setZoom(400)).not.toThrow();
    });

    it('should throw error for zoom < 10', () => {
      expect(() => worksheet.setZoom(5)).toThrow();
    });

    it('should throw error for zoom > 400', () => {
      expect(() => worksheet.setZoom(500)).toThrow();
    });
  });

  describe('Worksheet Selection', () => {
    it('should set default selection', () => {
      expect(() => worksheet.setSelection(0, 0, 5, 5)).not.toThrow();
    });

    it('should create file with selection', () => {
      worksheet.write(0, 0, 'A1');
      worksheet.write(5, 5, 'F6');
      worksheet.setSelection(0, 0, 5, 5); // Select A1:F6

      const filename = path.join(__dirname, 'temp-selection-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Worksheet Top Left Cell', () => {
    it('should set top left cell', () => {
      expect(() => worksheet.setTopLeftCell(10, 0)).not.toThrow();
    });

    it('should create file with custom top left', () => {
      for (let i = 0; i < 20; i++) {
        worksheet.write(i, 0, `Row ${i + 1}`);
      }
      worksheet.setTopLeftCell(10, 0); // Start view at row 11

      const filename = path.join(__dirname, 'temp-top-left-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Worksheet Protection', () => {
    it('should protect worksheet', () => {
      expect(() => worksheet.protect()).not.toThrow();
    });

    it('should protect worksheet with password', () => {
      expect(() => worksheet.protectWithPassword('secret123')).not.toThrow();
    });

    it('should create protected worksheet', () => {
      worksheet.write(0, 0, 'Protected');
      worksheet.protect();

      const filename = path.join(__dirname, 'temp-protected-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });

    it('should create password protected worksheet', () => {
      worksheet.write(0, 0, 'Password Protected');
      worksheet.protectWithPassword('test123');

      const filename = path.join(__dirname, 'temp-password-protected-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });

    it('should allow unlocked cells to be editable when protected', () => {
      const lockedFormat = new Format().setLocked();
      const unlockedFormat = new Format().setUnlocked();

      worksheet.writeWithFormat(0, 0, 'Locked', lockedFormat);
      worksheet.writeWithFormat(0, 1, 'Unlocked', unlockedFormat);
      worksheet.protect();

      const filename = path.join(__dirname, 'temp-mixed-protection-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Advanced Format - Font Script', () => {
    it('should set superscript', () => {
      const format = new Format();
      expect(() => format.setFontScript('superscript')).not.toThrow();
    });

    it('should set subscript', () => {
      const format = new Format();
      expect(() => format.setFontScript('subscript')).not.toThrow();
    });

    it('should throw error for invalid script', () => {
      const format = new Format();
      expect(() => format.setFontScript('invalid')).toThrow();
    });

    it('should create file with superscript and subscript', () => {
      const superFormat = new Format().setFontScript('superscript');
      const subFormat = new Format().setFontScript('subscript');

      worksheet.writeWithFormat(0, 0, 'X2', superFormat); // X²
      worksheet.writeWithFormat(1, 0, 'H2O', subFormat);  // H₂O

      const filename = path.join(__dirname, 'temp-script-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Advanced Format - Diagonal Borders', () => {
    it('should set diagonal border', () => {
      const format = new Format();
      expect(() => format.setBorderDiagonal('thin')).not.toThrow();
    });

    it('should set diagonal border color', () => {
      const format = new Format();
      expect(() => format.setBorderDiagonalColor('red')).not.toThrow();
    });

    it('should set diagonal border type - up', () => {
      const format = new Format();
      expect(() => format.setBorderDiagonalType('up')).not.toThrow();
    });

    it('should set diagonal border type - down', () => {
      const format = new Format();
      expect(() => format.setBorderDiagonalType('down')).not.toThrow();
    });

    it('should set diagonal border type - both', () => {
      const format = new Format();
      expect(() => format.setBorderDiagonalType('both')).not.toThrow();
    });

    it('should throw error for invalid border type', () => {
      const format = new Format();
      expect(() => format.setBorderDiagonalType('invalid')).toThrow();
    });

    it('should create file with diagonal borders', () => {
      const upFormat = new Format()
        .setBorderDiagonal('medium')
        .setBorderDiagonalColor('blue')
        .setBorderDiagonalType('up');

      const downFormat = new Format()
        .setBorderDiagonal('medium')
        .setBorderDiagonalColor('red')
        .setBorderDiagonalType('down');

      const bothFormat = new Format()
        .setBorderDiagonal('thick')
        .setBorderDiagonalColor('green')
        .setBorderDiagonalType('both');

      worksheet.writeWithFormat(0, 0, 'Up', upFormat);
      worksheet.writeWithFormat(0, 1, 'Down', downFormat);
      worksheet.writeWithFormat(0, 2, 'Both', bothFormat);

      const filename = path.join(__dirname, 'temp-diagonal-border-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Advanced Format - Reading Direction', () => {
    it('should set reading direction to context', () => {
      const format = new Format();
      expect(() => format.setReadingDirection(0)).not.toThrow();
    });

    it('should set reading direction to LTR', () => {
      const format = new Format();
      expect(() => format.setReadingDirection(1)).not.toThrow();
    });

    it('should set reading direction to RTL', () => {
      const format = new Format();
      expect(() => format.setReadingDirection(2)).not.toThrow();
    });

    it('should throw error for invalid direction', () => {
      const format = new Format();
      expect(() => format.setReadingDirection(3)).toThrow();
    });

    it('should create file with RTL text', () => {
      const rtlFormat = new Format().setReadingDirection(2);
      worksheet.writeWithFormat(0, 0, 'مرحبا', rtlFormat); // Arabic "Hello"

      const filename = path.join(__dirname, 'temp-rtl-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Advanced Format - Quote Prefix', () => {
    it('should set quote prefix', () => {
      const format = new Format();
      expect(() => format.setQuotePrefix()).not.toThrow();
    });

    it('should create file with quote prefix', () => {
      const quoteFormat = new Format().setQuotePrefix();
      worksheet.writeWithFormat(0, 0, '00123', quoteFormat); // Display as 00123, not 123

      const filename = path.join(__dirname, 'temp-quote-prefix-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Advanced Format - setLocked', () => {
    it('should lock cell', () => {
      const format = new Format();
      expect(() => format.setLocked()).not.toThrow();
    });

    it('should support method chaining', () => {
      const format = new Format();
      expect(() => {
        format.setBold().setLocked().setBackgroundColor('yellow');
      }).not.toThrow();
    });
  });

  describe('Complex Scenarios', () => {
    it('should combine multiple new features', () => {
      // Create a worksheet with all new features
      worksheet.setTabColor('purple');
      worksheet.setZoom(125);

      // Blank cells with formatting
      const headerFormat = new Format()
        .setBold()
        .setBackgroundColor('blue')
        .setFontColor('white');

      for (let col = 0; col < 5; col++) {
        worksheet.writeBlank(0, col, headerFormat);
      }

      // Diagonal borders
      const diagFormat = new Format()
        .setBorderDiagonal('medium')
        .setBorderDiagonalType('both');

      worksheet.writeWithFormat(2, 2, 'X', diagFormat);

      // Superscript
      const superFormat = new Format().setFontScript('superscript');
      worksheet.writeWithFormat(4, 0, 'E=mc2', superFormat);

      // Protection
      worksheet.protect();

      const filename = path.join(__dirname, 'temp-complex-features-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });

    it('should create styled protected worksheet', () => {
      worksheet.setTabColor('green');
      worksheet.setZoom(100);

      const headerFormat = new Format()
        .setBold()
        .setBackgroundColor('gray')
        .setFontColor('white')
        .setLocked();

      const dataFormat = new Format().setUnlocked();

      worksheet.writeWithFormat(0, 0, 'Name', headerFormat);
      worksheet.writeWithFormat(0, 1, 'Value', headerFormat);

      worksheet.writeWithFormat(1, 0, '', dataFormat);
      worksheet.writeWithFormat(1, 1, '', dataFormat);

      worksheet.protectWithPassword('test');

      const filename = path.join(__dirname, 'temp-styled-protected-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });
});
