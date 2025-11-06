/**
 * Tests for Conditional Formatting and Page Setup functionality
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Conditional Formatting', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('2-Color Scale', () => {
    test('should add 2-color scale conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('2ColorScale');

      // Write test data
      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      // Apply 2-color scale
      expect(() => {
        worksheet.addConditionalFormat2ColorScale(0, 0, 9, 0, 'red', 'green');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-2color-scale.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should accept hex colors for 2-color scale', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('2ColorHex');

      for (let i = 0; i < 5; i++) {
        worksheet.writeNumber(i, 0, i * 20);
      }

      expect(() => {
        worksheet.addConditionalFormat2ColorScale(0, 0, 4, 0, '#FF0000', '#00FF00');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-2color-hex.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle single cell range for 2-color scale', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('2ColorSingle');

      worksheet.writeNumber(0, 0, 50);

      expect(() => {
        worksheet.addConditionalFormat2ColorScale(0, 0, 0, 0, 'blue', 'yellow');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-2color-single.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('3-Color Scale', () => {
    test('should add 3-color scale conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('3ColorScale');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      expect(() => {
        worksheet.addConditionalFormat3ColorScale(0, 0, 9, 0, 'red', 'yellow', 'green');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-3color-scale.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should accept hex colors for 3-color scale', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('3ColorHex');

      for (let i = 0; i < 5; i++) {
        worksheet.writeNumber(i, 0, i * 20);
      }

      expect(() => {
        worksheet.addConditionalFormat3ColorScale(0, 0, 4, 0, '#0000FF', '#FFFF00', '#FF0000');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-3color-hex.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle rectangular range for 3-color scale', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('3ColorRect');

      for (let i = 0; i < 5; i++) {
        for (let j = 0; j < 3; j++) {
          worksheet.writeNumber(i, j, i * 3 + j);
        }
      }

      expect(() => {
        worksheet.addConditionalFormat3ColorScale(0, 0, 4, 2, 'blue', 'white', 'red');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-3color-rect.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Data Bar', () => {
    test('should add data bar conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('DataBar');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 5);
      }

      expect(() => {
        worksheet.addConditionalFormatDataBar(0, 0, 9, 0, 'blue');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-data-bar.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle multiple data bar columns', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('DataBarMulti');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 5);
        worksheet.writeNumber(i, 1, (i + 1) * 3);
      }

      expect(() => {
        worksheet.addConditionalFormatDataBar(0, 0, 9, 0, 'green');
        worksheet.addConditionalFormatDataBar(0, 1, 9, 1, 'red');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-data-bar-multi.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should accept hex colors for data bar', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('DataBarHex');

      for (let i = 0; i < 5; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 10);
      }

      expect(() => {
        worksheet.addConditionalFormatDataBar(0, 0, 4, 0, '#FF6600');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-data-bar-hex.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Cell Rules', () => {
    test('should add greaterThan cell rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellGreaterThan');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format()
        .setBold()
        .setBackgroundColor('yellow');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'greaterThan', 50, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-greater-than.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add lessThan cell rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellLessThan');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format()
        .setBackgroundColor('red')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'lessThan', 30, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-less-than.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add equal cell rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellEqual');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format()
        .setBold()
        .setBackgroundColor('green')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'equal', 50, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-equal.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add notEqual cell rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellNotEqual');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format()
        .setBackgroundColor('silver');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'notEqual', 50, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-not-equal.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add greaterThanOrEqual cell rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellGTE');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format()
        .setBold()
        .setFontColor('blue');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'greaterThanOrEqual', 50, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-gte.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add lessThanOrEqual cell rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellLTE');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format()
        .setItalic()
        .setFontColor('purple');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'lessThanOrEqual', 50, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-lte.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle multiple cell rules on same range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CellMultiRules');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, Math.random() * 100);
      }

      const lowFormat = new Format()
        .setBackgroundColor('red')
        .setFontColor('white');

      const highFormat = new Format()
        .setBackgroundColor('green')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'lessThan', 30, lowFormat);
        worksheet.addConditionalFormatCell(0, 0, 9, 0, 'greaterThan', 70, highFormat);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-cell-multi-rules.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Text Rules', () => {
    test('should add contains text rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('TextContains');

      const testData = ['Completed', 'In Progress', 'Completed', 'Pending', 'Completed'];
      for (let i = 0; i < testData.length; i++) {
        worksheet.writeString(i, 0, testData[i]);
      }

      const format = new Format()
        .setBold()
        .setBackgroundColor('green')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatText(0, 0, 4, 0, 'contains', 'Completed', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-text-contains.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add notContains text rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('TextNotContains');

      const testData = ['Active', 'Inactive', 'Active', 'Suspended', 'Active'];
      for (let i = 0; i < testData.length; i++) {
        worksheet.writeString(i, 0, testData[i]);
      }

      const format = new Format()
        .setBackgroundColor('yellow');

      expect(() => {
        worksheet.addConditionalFormatText(0, 0, 4, 0, 'notContains', 'Active', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-text-not-contains.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add beginsWith text rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('TextBeginsWith');

      const testData = ['Error: Connection failed', 'Warning: Low memory', 'Error: Timeout', 'Info: Success', 'Error: Not found'];
      for (let i = 0; i < testData.length; i++) {
        worksheet.writeString(i, 0, testData[i]);
      }

      const format = new Format()
        .setBold()
        .setBackgroundColor('red')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatText(0, 0, 4, 0, 'beginsWith', 'Error', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-text-begins-with.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add endsWith text rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('TextEndsWith');

      const testData = ['report.pdf', 'data.csv', 'summary.pdf', 'config.json', 'analysis.pdf'];
      for (let i = 0; i < testData.length; i++) {
        worksheet.writeString(i, 0, testData[i]);
      }

      const format = new Format()
        .setBold()
        .setBackgroundColor('blue')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatText(0, 0, 4, 0, 'endsWith', '.pdf', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-text-ends-with.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle multiple text rules on same range', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('TextMultiRules');

      const testData = ['Success', 'Error: Failed', 'Warning: Low', 'Success', 'Error: Timeout'];
      for (let i = 0; i < testData.length; i++) {
        worksheet.writeString(i, 0, testData[i]);
      }

      const successFormat = new Format()
        .setBackgroundColor('green')
        .setFontColor('white');

      const errorFormat = new Format()
        .setBackgroundColor('red')
        .setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatText(0, 0, 4, 0, 'contains', 'Success', successFormat);
        worksheet.addConditionalFormatText(0, 0, 4, 0, 'beginsWith', 'Error', errorFormat);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-text-multi-rules.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Error Handling', () => {
    test('should throw error for invalid color in 2-color scale', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('InvalidColor');

      worksheet.writeNumber(0, 0, 100);

      expect(() => {
        worksheet.addConditionalFormat2ColorScale(0, 0, 0, 0, 'invalid_color', 'green');
      }).toThrow();
    });

    test('should throw error for invalid text rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('InvalidRule');

      worksheet.writeString(0, 0, 'Test');

      const format = new Format();

      expect(() => {
        worksheet.addConditionalFormatText(0, 0, 0, 0, 'invalid_rule', 'Test', format);
      }).toThrow();
    });
  });
});

describe('Page Setup', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Page Orientation', () => {
    test('should set portrait orientation', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Portrait');

      worksheet.writeString(0, 0, 'Portrait Test');

      expect(() => {
        worksheet.setPortrait();
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-portrait.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set landscape orientation', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Landscape');

      worksheet.writeString(0, 0, 'Landscape Test');

      expect(() => {
        worksheet.setLandscape();
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-landscape.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Margins', () => {
    test('should set page margins', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Margins');

      worksheet.writeString(0, 0, 'Margins Test');

      expect(() => {
        worksheet.setMargins(1.0, 1.0, 1.5, 1.5, 0.75, 0.75);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-margins.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set narrow margins', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('NarrowMargins');

      worksheet.writeString(0, 0, 'Narrow Margins Test');

      expect(() => {
        worksheet.setMargins(0.25, 0.25, 0.25, 0.25, 0.25, 0.25);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-narrow-margins.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set wide margins', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('WideMargins');

      worksheet.writeString(0, 0, 'Wide Margins Test');

      expect(() => {
        worksheet.setMargins(2.0, 2.0, 2.0, 2.0, 1.0, 1.0);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-wide-margins.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Paper Size', () => {
    test('should set A4 paper size', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('A4');

      worksheet.writeString(0, 0, 'A4 Paper Test');

      expect(() => {
        worksheet.setPaperSize(9); // A4
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-a4.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set Letter paper size', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Letter');

      worksheet.writeString(0, 0, 'Letter Paper Test');

      expect(() => {
        worksheet.setPaperSize(1); // Letter
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-letter.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set Legal paper size', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Legal');

      worksheet.writeString(0, 0, 'Legal Paper Test');

      expect(() => {
        worksheet.setPaperSize(5); // Legal
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-legal.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Headers and Footers', () => {
    test('should set header', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Header');

      worksheet.writeString(0, 0, 'Header Test');

      expect(() => {
        worksheet.setHeader('&CTest Header');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-header.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set footer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Footer');

      worksheet.writeString(0, 0, 'Footer Test');

      expect(() => {
        worksheet.setFooter('&CPage &P of &N');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-footer.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set both header and footer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('HeaderFooter');

      worksheet.writeString(0, 0, 'Header and Footer Test');

      expect(() => {
        worksheet.setHeader('&LLeft Header&CCenter Header&RRight Header');
        worksheet.setFooter('&LLeft Footer&C&D&RPage &P');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-header-footer.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should handle special codes in header and footer', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('SpecialCodes');

      worksheet.writeString(0, 0, 'Special Codes Test');

      expect(() => {
        worksheet.setHeader('&LDate: &D&CTime: &T&RPage &P of &N');
        worksheet.setFooter('&L&F&C&A&R&Z');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-special-codes.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print Area', () => {
    test('should set print area', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('PrintArea');

      // Write more data than print area
      for (let i = 0; i < 20; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setPrintArea(0, 0, 9, 4); // Only print first 10 rows, 5 columns
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-print-area.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set small print area', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('SmallPrintArea');

      for (let i = 0; i < 10; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setPrintArea(0, 0, 4, 4); // 5x5 area
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-small-print-area.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print Options', () => {
    test('should enable print gridlines', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('PrintGridlines');

      worksheet.writeString(0, 0, 'Gridlines Test');

      expect(() => {
        worksheet.setPrintGridlines(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-print-gridlines.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should disable print gridlines', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('NoPrintGridlines');

      worksheet.writeString(0, 0, 'No Gridlines Test');

      expect(() => {
        worksheet.setPrintGridlines(false);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-no-print-gridlines.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should center horizontally', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CenterH');

      worksheet.writeString(0, 0, 'Center Horizontally Test');

      expect(() => {
        worksheet.setPrintCenterHorizontally(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-center-h.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should center vertically', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CenterV');

      worksheet.writeString(0, 0, 'Center Vertically Test');

      expect(() => {
        worksheet.setPrintCenterVertically(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-center-v.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should center both horizontally and vertically', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('CenterBoth');

      worksheet.writeString(0, 0, 'Center Both Test');

      expect(() => {
        worksheet.setPrintCenterHorizontally(true);
        worksheet.setPrintCenterVertically(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-center-both.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Combined Page Setup', () => {
    test('should apply multiple page setup options', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Combined');

      for (let i = 0; i < 10; i++) {
        for (let j = 0; j < 5; j++) {
          worksheet.writeNumber(i, j, i * 5 + j);
        }
      }

      expect(() => {
        worksheet.setLandscape();
        worksheet.setMargins(1.0, 1.0, 1.5, 1.5, 0.75, 0.75);
        worksheet.setPaperSize(9); // A4
        worksheet.setHeader('&CComprehensive Test');
        worksheet.setFooter('&LDate: &D&RPage &P');
        worksheet.setPrintArea(0, 0, 9, 4);
        worksheet.setPrintGridlines(true);
        worksheet.setPrintCenterHorizontally(true);
        worksheet.setPrintCenterVertically(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-combined.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });
});
