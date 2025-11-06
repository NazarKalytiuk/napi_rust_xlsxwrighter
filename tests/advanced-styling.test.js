/**
 * Tests for Advanced Styling Features
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format } = require('../wrapper');

describe('Advanced Styling', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => {
    if (fs.existsSync(testOutputDir)) {
      fs.rmSync(testOutputDir, { recursive: true, force: true });
    }
  });

  describe('Font Styling', () => {
    test('should set font name', () => {
      const format = new Format();
      const result = format.setFontName('Arial');
      expect(result).toBe(format); // Check chaining
    });

    test('should set underline style - single', () => {
      const format = new Format();
      expect(() => {
        format.setUnderline('single');
      }).not.toThrow();
    });

    test('should set underline style - double', () => {
      const format = new Format();
      expect(() => {
        format.setUnderline('double');
      }).not.toThrow();
    });

    test('should set underline style - singleAccounting', () => {
      const format = new Format();
      expect(() => {
        format.setUnderline('singleAccounting');
      }).not.toThrow();
    });

    test('should set underline style - doubleAccounting', () => {
      const format = new Format();
      expect(() => {
        format.setUnderline('doubleAccounting');
      }).not.toThrow();
    });

    test('should reject invalid underline style', () => {
      const format = new Format();
      expect(() => {
        format.setUnderline('invalid');
      }).toThrow();
    });

    test('should set strikethrough', () => {
      const format = new Format();
      const result = format.setStrikethrough();
      expect(result).toBe(format);
    });
  });

  describe('Text Formatting', () => {
    test('should set text wrap', () => {
      const format = new Format();
      const result = format.setTextWrap();
      expect(result).toBe(format);
    });

    test('should set indent', () => {
      const format = new Format();
      expect(() => {
        format.setIndent(3);
      }).not.toThrow();
    });

    test('should reject invalid indent (too high)', () => {
      const format = new Format();
      expect(() => {
        format.setIndent(20);
      }).toThrow();
    });

    test('should set rotation - positive degrees', () => {
      const format = new Format();
      expect(() => {
        format.setRotation(45);
      }).not.toThrow();
    });

    test('should set rotation - negative degrees', () => {
      const format = new Format();
      expect(() => {
        format.setRotation(-30);
      }).not.toThrow();
    });

    test('should set rotation - vertical text (270)', () => {
      const format = new Format();
      expect(() => {
        format.setRotation(270);
      }).not.toThrow();
    });

    test('should reject invalid rotation', () => {
      const format = new Format();
      expect(() => {
        format.setRotation(100);
      }).toThrow();
    });

    test('should set shrink', () => {
      const format = new Format();
      const result = format.setShrink();
      expect(result).toBe(format);
    });
  });

  describe('Individual Borders', () => {
    test('should set top border', () => {
      const format = new Format();
      expect(() => {
        format.setBorderTop('thick');
      }).not.toThrow();
    });

    test('should set bottom border', () => {
      const format = new Format();
      expect(() => {
        format.setBorderBottom('medium');
      }).not.toThrow();
    });

    test('should set left border', () => {
      const format = new Format();
      expect(() => {
        format.setBorderLeft('dashed');
      }).not.toThrow();
    });

    test('should set right border', () => {
      const format = new Format();
      expect(() => {
        format.setBorderRight('dotted');
      }).not.toThrow();
    });

    test('should set border color for all borders', () => {
      const format = new Format();
      expect(() => {
        format.setBorderColor('red');
      }).not.toThrow();
    });

    test('should set top border color', () => {
      const format = new Format();
      expect(() => {
        format.setBorderTopColor('#FF0000');
      }).not.toThrow();
    });

    test('should set bottom border color', () => {
      const format = new Format();
      expect(() => {
        format.setBorderBottomColor('blue');
      }).not.toThrow();
    });

    test('should set left border color', () => {
      const format = new Format();
      expect(() => {
        format.setBorderLeftColor('green');
      }).not.toThrow();
    });

    test('should set right border color', () => {
      const format = new Format();
      expect(() => {
        format.setBorderRightColor('orange');
      }).not.toThrow();
    });

    test('should support extended border styles', () => {
      const format = new Format();
      const styles = ['dashDot', 'dashDotDot', 'slantDashDot', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot'];

      styles.forEach(style => {
        expect(() => {
          format.setBorderTop(style);
        }).not.toThrow();
      });
    });
  });

  describe('Fill Patterns', () => {
    test('should set foreground color', () => {
      const format = new Format();
      expect(() => {
        format.setForegroundColor('#4472C4');
      }).not.toThrow();
    });

    test('should set pattern - solid', () => {
      const format = new Format();
      expect(() => {
        format.setPattern('solid');
      }).not.toThrow();
    });

    test('should set pattern - mediumGray', () => {
      const format = new Format();
      expect(() => {
        format.setPattern('mediumGray');
      }).not.toThrow();
    });

    test('should set pattern - darkGray', () => {
      const format = new Format();
      expect(() => {
        format.setPattern('darkGray');
      }).not.toThrow();
    });

    test('should reject invalid pattern', () => {
      const format = new Format();
      expect(() => {
        format.setPattern('invalid');
      }).toThrow();
    });
  });

  describe('Cell Protection', () => {
    test('should set unlocked', () => {
      const format = new Format();
      const result = format.setUnlocked();
      expect(result).toBe(format);
    });

    test('should set hidden', () => {
      const format = new Format();
      const result = format.setHidden();
      expect(result).toBe(format);
    });
  });

  describe('Method Chaining', () => {
    test('should support chaining all methods', () => {
      const format = new Format();
      const result = format
        .setBold()
        .setItalic()
        .setFontName('Arial')
        .setFontSize(14)
        .setUnderline('double')
        .setStrikethrough()
        .setFontColor('red')
        .setBackgroundColor('yellow')
        .setForegroundColor('blue')
        .setPattern('mediumGray')
        .setAlign('center')
        .setVerticalAlign('center')
        .setTextWrap()
        .setIndent(2)
        .setRotation(45)
        .setShrink()
        .setBorder('thin')
        .setBorderTop('thick')
        .setBorderBottom('medium')
        .setBorderLeft('dashed')
        .setBorderRight('dotted')
        .setBorderColor('#FF0000')
        .setBorderTopColor('red')
        .setBorderBottomColor('blue')
        .setBorderLeftColor('green')
        .setBorderRightColor('purple')
        .setNumFormat('#,##0.00')
        .setUnlocked()
        .setHidden();

      expect(result).toBe(format);
    });
  });

  describe('Integration Tests', () => {
    test('should create file with all styling features', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Styling');

      // Font styling
      const fontFormat = new Format()
        .setFontName('Times New Roman')
        .setFontSize(14)
        .setUnderline('double')
        .setStrikethrough()
        .setFontColor('purple');

      worksheet.write(0, 0, 'Font Styling', fontFormat);

      // Text formatting
      const textFormat = new Format()
        .setTextWrap()
        .setIndent(2)
        .setRotation(45);

      worksheet.setRowHeight(1, 50);
      worksheet.write(1, 0, 'Text Format', textFormat);

      // Individual borders
      const borderFormat = new Format()
        .setBorderTop('thick')
        .setBorderTopColor('red')
        .setBorderBottom('medium')
        .setBorderBottomColor('blue')
        .setBorderLeft('dashed')
        .setBorderLeftColor('green')
        .setBorderRight('dotted')
        .setBorderRightColor('orange');

      worksheet.write(2, 0, 'Borders', borderFormat);

      // Patterns
      const patternFormat = new Format()
        .setPattern('mediumGray')
        .setBackgroundColor('#D9E1F2')
        .setForegroundColor('#4472C4');

      worksheet.write(3, 0, 'Pattern', patternFormat);

      const outputPath = path.join(testOutputDir, 'advanced-styling-test.xlsx');
      workbook.save(outputPath);

      expect(fs.existsSync(outputPath)).toBe(true);
      const stats = fs.statSync(outputPath);
      expect(stats.size).toBeGreaterThan(0);
    });
  });
});
