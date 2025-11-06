/**
 * Tests for 100% Coverage Features (v0.7.0)
 * - 8 new conditional formatting types
 * - 10 new page setup methods
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format } = require('../wrapper');

describe('100% Coverage - Conditional Formatting (Advanced)', () => {
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

  describe('Average Conditional Formatting', () => {
    test('should add above average conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Above Average');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, 50 + (i - 5) * 10);
      }

      const format = new Format().setBackgroundColor('yellow');

      expect(() => {
        worksheet.addConditionalFormatAverage(0, 0, 9, 0, 'aboveAverage', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-average-above.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add below average conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Below Average');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, 50 + (i - 5) * 10);
      }

      const format = new Format().setBackgroundColor('red').setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatAverage(0, 0, 9, 0, 'belowAverage', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-average-below.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add equal or above average conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Equal or Above');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format().setBold();

      expect(() => {
        worksheet.addConditionalFormatAverage(0, 0, 9, 0, 'equalOrAboveAverage', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-average-eq-above.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add equal or below average conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Equal or Below');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i * 10);
      }

      const format = new Format().setItalic();

      expect(() => {
        worksheet.addConditionalFormatAverage(0, 0, 9, 0, 'equalOrBelowAverage', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-average-eq-below.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Top/Bottom N Conditional Formatting', () => {
    test('should add top N values conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Top N');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, Math.random() * 100);
      }

      const format = new Format().setBold().setBackgroundColor('green').setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatTop(0, 0, 9, 0, 'top', 3, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-top-n.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add bottom N values conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Bottom N');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, Math.random() * 100);
      }

      const format = new Format().setBold().setBackgroundColor('red').setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatTop(0, 0, 9, 0, 'bottom', 3, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-bottom-n.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add top N percent conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Top Percent');

      for (let i = 0; i < 20; i++) {
        worksheet.writeNumber(i, 0, i * 5);
      }

      const format = new Format().setBackgroundColor('yellow');

      expect(() => {
        worksheet.addConditionalFormatTop(0, 0, 19, 0, 'topPercent', 25, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-top-percent.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add bottom N percent conditional formatting', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Bottom Percent');

      for (let i = 0; i < 20; i++) {
        worksheet.writeNumber(i, 0, i * 5);
      }

      const format = new Format().setBackgroundColor('orange');

      expect(() => {
        worksheet.addConditionalFormatTop(0, 0, 19, 0, 'bottomPercent', 25, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-bottom-percent.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Duplicate/Unique Values Conditional Formatting', () => {
    test('should highlight duplicate values', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Duplicates');

      const values = [1, 2, 3, 2, 4, 5, 3, 6, 1, 7, 8, 2];
      for (let i = 0; i < values.length; i++) {
        worksheet.writeNumber(i, 0, values[i]);
      }

      const format = new Format().setBackgroundColor('red').setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatDuplicate(0, 0, 11, 0, false, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-duplicates.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should highlight unique values', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Unique');

      const values = [1, 2, 3, 2, 4, 5, 3, 6, 1, 7, 8, 2];
      for (let i = 0; i < values.length; i++) {
        worksheet.writeNumber(i, 0, values[i]);
      }

      const format = new Format().setBackgroundColor('green').setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatDuplicate(0, 0, 11, 0, true, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-unique.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Blank/Non-blank Cells Conditional Formatting', () => {
    test('should highlight blank cells', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Blanks');

      for (let i = 0; i < 10; i++) {
        if (i % 3 !== 0) {
          worksheet.writeString(i, 0, 'Data');
        }
      }

      const format = new Format().setBackgroundColor('silver');

      expect(() => {
        worksheet.addConditionalFormatBlank(0, 0, 9, 0, false, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-blanks.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should highlight non-blank cells', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Non-Blanks');

      for (let i = 0; i < 10; i++) {
        if (i % 3 !== 0) {
          worksheet.writeString(i, 0, 'Data');
        }
      }

      const format = new Format().setBackgroundColor('cyan');

      expect(() => {
        worksheet.addConditionalFormatBlank(0, 0, 9, 0, true, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-non-blanks.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Error Cells Conditional Formatting', () => {
    test('should highlight error cells', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Errors');

      worksheet.writeNumber(0, 0, 10);
      worksheet.writeNumber(1, 0, 0);
      worksheet.writeString(2, 0, '=A1/A2'); // Division by zero error

      const format = new Format().setBackgroundColor('red').setFontColor('white');

      expect(() => {
        worksheet.addConditionalFormatError(0, 0, 9, 0, false, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-errors.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should highlight non-error cells', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Non-Errors');

      worksheet.writeNumber(0, 0, 10);
      worksheet.writeNumber(1, 0, 0);
      worksheet.writeString(2, 0, '=A1/A2'); // Division by zero error

      const format = new Format().setBackgroundColor('green');

      expect(() => {
        worksheet.addConditionalFormatError(0, 0, 9, 0, true, format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-non-errors.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Date-based Conditional Formatting', () => {
    test('should add date conditional formatting - today', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Date Today');

      const today = new Date();
      worksheet.writeNumber(0, 0, today.getTime() / 86400000 + 25569); // Excel date serial

      const format = new Format().setBackgroundColor('yellow');

      expect(() => {
        worksheet.addConditionalFormatDate(0, 0, 9, 0, 'today', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-date-today.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add date conditional formatting - this week', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Date This Week');

      const today = new Date();
      worksheet.writeNumber(0, 0, today.getTime() / 86400000 + 25569);

      const format = new Format().setBackgroundColor('green');

      expect(() => {
        worksheet.addConditionalFormatDate(0, 0, 9, 0, 'thisWeek', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-date-week.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Formula-based Conditional Formatting', () => {
    test('should add formula-based conditional formatting - even rows', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Formula Even');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, i + 1);
      }

      const format = new Format().setBackgroundColor('cyan');

      expect(() => {
        worksheet.addConditionalFormatFormula(0, 0, 9, 0, '=MOD(A1,2)=0', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-formula-even.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add formula-based conditional formatting - custom condition', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Formula Custom');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, Math.random() * 100);
      }

      const format = new Format().setBold().setFontColor('red');

      expect(() => {
        worksheet.addConditionalFormatFormula(0, 0, 9, 0, '=A1>50', format);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-formula-custom.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Icon Set Conditional Formatting', () => {
    test('should add icon set - three arrows', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Icons 3 Arrows');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 10);
      }

      expect(() => {
        worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeArrows');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-icons-3arrows.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add icon set - three traffic lights', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Icons Traffic');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 10);
      }

      expect(() => {
        worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeTrafficLights');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-icons-traffic.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add icon set - five arrows', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Icons 5 Arrows');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 10);
      }

      expect(() => {
        worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fiveArrows');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-icons-5arrows.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should add icon set - four histograms', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Icons Histograms');

      for (let i = 0; i < 10; i++) {
        worksheet.writeNumber(i, 0, (i + 1) * 10);
      }

      expect(() => {
        worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fourHistograms');
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'conditional-icons-hist.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });
});

describe('100% Coverage - Page Setup (Advanced)', () => {
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

  describe('Print Fit to Pages', () => {
    test('should fit to 1 page wide', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Fit 1 Wide');

      for (let i = 0; i < 50; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setPrintFitToPages(1, 0);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-fit-1-wide.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should fit to 1 page tall', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Fit 1 Tall');

      for (let i = 0; i < 50; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setPrintFitToPages(0, 1);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-fit-1-tall.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should fit to 2x2 pages', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Fit 2x2');

      for (let i = 0; i < 50; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setPrintFitToPages(2, 2);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-fit-2x2.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print Scale', () => {
    test('should set print scale to 75%', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Scale 75');

      worksheet.writeString(0, 0, 'Scaled to 75%');

      expect(() => {
        worksheet.setPrintScale(75);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-scale-75.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should set print scale to 150%', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Scale 150');

      worksheet.writeString(0, 0, 'Scaled to 150%');

      expect(() => {
        worksheet.setPrintScale(150);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-scale-150.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print First Page Number', () => {
    test('should set first page number to 5', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('First Page 5');

      worksheet.writeString(0, 0, 'First page is 5');

      expect(() => {
        worksheet.setPrintFirstPageNumber(5);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-first-page-5.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print Black and White', () => {
    test('should enable black and white printing', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('BW Print');

      const format = new Format().setBackgroundColor('red');
      worksheet.write(0, 0, 'Color', format);

      expect(() => {
        worksheet.setPrintBlackAndWhite(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-bw-print.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print Draft', () => {
    test('should enable draft printing', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Draft');

      worksheet.writeString(0, 0, 'Draft quality');

      expect(() => {
        worksheet.setPrintDraft(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-draft.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Print Headings', () => {
    test('should enable printing row and column headings', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('With Headings');

      worksheet.writeString(0, 0, 'Data');

      expect(() => {
        worksheet.setPrintHeadings(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-headings.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Repeat Rows and Columns', () => {
    test('should repeat first row on each page', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Repeat Row');

      worksheet.writeString(0, 0, 'Header');
      for (let i = 1; i < 50; i++) {
        worksheet.writeNumber(i, 0, i);
      }

      expect(() => {
        worksheet.setRepeatRows(0, 0);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-repeat-row.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should repeat first column on each page', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Repeat Col');

      worksheet.writeString(0, 0, 'Header');
      for (let i = 0; i < 20; i++) {
        for (let j = 1; j < 20; j++) {
          worksheet.writeNumber(i, j, i * j);
        }
      }

      expect(() => {
        worksheet.setRepeatColumns(0, 0);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-repeat-col.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should repeat first 2 rows on each page', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Repeat 2 Rows');

      worksheet.writeString(0, 0, 'Header 1');
      worksheet.writeString(1, 0, 'Header 2');
      for (let i = 2; i < 50; i++) {
        worksheet.writeNumber(i, 0, i);
      }

      expect(() => {
        worksheet.setRepeatRows(0, 1);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-repeat-2-rows.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should repeat both rows and columns', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('Repeat Both');

      worksheet.writeString(0, 0, 'Header');
      for (let i = 0; i < 50; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setRepeatRows(0, 0);
        worksheet.setRepeatColumns(0, 0);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-repeat-both.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Header/Footer Options', () => {
    test('should scale header and footer with document', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('HF Scale');

      worksheet.setHeader('&CScaled Header');
      worksheet.writeString(0, 0, 'Data');

      expect(() => {
        worksheet.setHeaderFooterScaleWithDoc(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-hf-scale.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });

    test('should align header and footer with page margins', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('HF Align');

      worksheet.setHeader('&CAligned Header');
      worksheet.writeString(0, 0, 'Data');

      expect(() => {
        worksheet.setHeaderFooterAlignWithPage(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-hf-align.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });

  describe('Combined Advanced Page Setup', () => {
    test('should apply all advanced page setup options', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet('All Options');

      worksheet.writeString(0, 0, 'Header Row');
      for (let i = 1; i < 50; i++) {
        for (let j = 0; j < 10; j++) {
          worksheet.writeNumber(i, j, i * 10 + j);
        }
      }

      expect(() => {
        worksheet.setPrintFitToPages(1, 0);
        worksheet.setPrintFirstPageNumber(10);
        worksheet.setPrintBlackAndWhite(false);
        worksheet.setPrintDraft(false);
        worksheet.setPrintHeadings(true);
        worksheet.setRepeatRows(0, 0);
        worksheet.setRepeatColumns(0, 0);
        worksheet.setHeader('&CComprehensive Test');
        worksheet.setFooter('&RPage &P');
        worksheet.setHeaderFooterScaleWithDoc(true);
        worksheet.setHeaderFooterAlignWithPage(true);
      }).not.toThrow();

      const filename = path.join(testOutputDir, 'page-all-options.xlsx');
      workbook.save(filename);
      expect(fs.existsSync(filename)).toBe(true);
    });
  });
});
