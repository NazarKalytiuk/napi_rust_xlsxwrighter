/**
 * Tests to achieve 100% coverage - Edge cases and enum variations
 */

import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import path from 'node:path';
import { Workbook, Format, Chart, Table, Sparkline } from '../wrapper.ts';
import { cleanupTestOutput, createTestOutputDir, readXlsxPart } from './test-helpers.mjs';

describe('Complete Coverage - Edge Cases and Enum Variations', () => {
  let testOutputDir;

  beforeAll(() => {
    testOutputDir = createTestOutputDir();
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  describe('Format - All BorderStyle Values', () => {
    test('should accept dashDot border style on individual borders', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format();

      expect(() => {
        format.setBorderTop('dashDot');
      }).not.toThrow();

      worksheet.write(0, 0, 'Test', format);
      workbook.save(path.join(testOutputDir, 'border_dashdot.xlsx'));
    });

    test('should accept dashDotDot border style on individual borders', () => {
      const format = new Format().setBorderBottom('dashDotDot');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept slantDashDot border style on individual borders', () => {
      const format = new Format().setBorderLeft('slantDashDot');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept mediumDashed border style on individual borders', () => {
      const format = new Format().setBorderRight('mediumDashed');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept mediumDashDot border style on individual borders', () => {
      const format = new Format().setBorderTop('mediumDashDot');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept mediumDashDotDot border style on individual borders', () => {
      const format = new Format().setBorderBottom('mediumDashDotDot');
      expect(format).toBeInstanceOf(Format);
    });

    test('should test all border styles in one file', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();

      const styles = [
        'dashDot',
        'dashDotDot',
        'slantDashDot',
        'mediumDashed',
        'mediumDashDot',
        'mediumDashDotDot',
      ];
      styles.forEach((style, index) => {
        const format = new Format().setBorderTop(style).setBorderBottom(style);
        worksheet.write(index, 0, style, format);
      });

      const file = path.join(testOutputDir, 'all_border_styles.xlsx');
      workbook.save(file);
      const stylesXml = readXlsxPart(file, 'xl/styles.xml');
      for (const style of styles) {
        expect(stylesXml).toContain(`style="${style}"`);
      }
    });
  });

  describe('Format - All FillPattern Values', () => {
    test('should accept mediumGray pattern', () => {
      const format = new Format().setPattern('mediumGray');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkGray pattern', () => {
      const format = new Format().setPattern('darkGray');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightGray pattern', () => {
      const format = new Format().setPattern('lightGray');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkHorizontal pattern', () => {
      const format = new Format().setPattern('darkHorizontal');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkVertical pattern', () => {
      const format = new Format().setPattern('darkVertical');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkDown pattern', () => {
      const format = new Format().setPattern('darkDown');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkUp pattern', () => {
      const format = new Format().setPattern('darkUp');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkGrid pattern', () => {
      const format = new Format().setPattern('darkGrid');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept darkTrellis pattern', () => {
      const format = new Format().setPattern('darkTrellis');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightHorizontal pattern', () => {
      const format = new Format().setPattern('lightHorizontal');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightVertical pattern', () => {
      const format = new Format().setPattern('lightVertical');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightDown pattern', () => {
      const format = new Format().setPattern('lightDown');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightUp pattern', () => {
      const format = new Format().setPattern('lightUp');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightGrid pattern', () => {
      const format = new Format().setPattern('lightGrid');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept lightTrellis pattern', () => {
      const format = new Format().setPattern('lightTrellis');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept gray125 pattern', () => {
      const format = new Format().setPattern('gray125');
      expect(format).toBeInstanceOf(Format);
    });

    test('should accept gray0625 pattern', () => {
      const format = new Format().setPattern('gray0625');
      expect(format).toBeInstanceOf(Format);
    });

    test('should test all fill patterns in one file', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();

      const patterns = [
        'mediumGray',
        'darkGray',
        'lightGray',
        'darkHorizontal',
        'darkVertical',
        'darkDown',
        'darkUp',
        'darkGrid',
        'darkTrellis',
        'lightHorizontal',
        'lightVertical',
        'lightDown',
        'lightUp',
        'lightGrid',
        'lightTrellis',
        'gray125',
        'gray0625',
      ];

      patterns.forEach((pattern, index) => {
        const format = new Format()
          .setPattern(pattern)
          .setForegroundColor('#FF0000')
          .setBackgroundColor('#FFFF00');
        worksheet.write(index, 0, pattern, format);
      });

      const file = path.join(testOutputDir, 'all_fill_patterns.xlsx');
      workbook.save(file);
      const stylesXml = readXlsxPart(file, 'xl/styles.xml');
      // Match the colored fill: every workbook already has a bare gray125 default fill.
      for (const pattern of patterns) {
        expect(stylesXml).toContain(
          `<patternFill patternType="${pattern}"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFF00"/></patternFill>`,
        );
      }
    });
  });

  describe('Chart - All Type Variations', () => {
    test('should accept area_stacked (underscore variant)', () => {
      const chart = new Chart('area_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept areaPercentStacked', () => {
      const chart = new Chart('areaPercentStacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept area_percent_stacked (underscore variant)', () => {
      const chart = new Chart('area_percent_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept bar_stacked (underscore variant)', () => {
      const chart = new Chart('bar_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept barPercentStacked', () => {
      const chart = new Chart('barPercentStacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept bar_percent_stacked (underscore variant)', () => {
      const chart = new Chart('bar_percent_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept column_stacked (underscore variant)', () => {
      const chart = new Chart('column_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept columnPercentStacked', () => {
      const chart = new Chart('columnPercentStacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept column_percent_stacked (underscore variant)', () => {
      const chart = new Chart('column_percent_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept line_stacked (underscore variant)', () => {
      const chart = new Chart('line_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept linePercentStacked', () => {
      const chart = new Chart('linePercentStacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept line_percent_stacked (underscore variant)', () => {
      const chart = new Chart('line_percent_stacked');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept radarWithMarkers', () => {
      const chart = new Chart('radarWithMarkers');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept radar_with_markers (underscore variant)', () => {
      const chart = new Chart('radar_with_markers');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept radarFilled', () => {
      const chart = new Chart('radarFilled');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept radar_filled (underscore variant)', () => {
      const chart = new Chart('radar_filled');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatterStraight', () => {
      const chart = new Chart('scatterStraight');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatter_straight (underscore variant)', () => {
      const chart = new Chart('scatter_straight');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatterStraightWithMarkers', () => {
      const chart = new Chart('scatterStraightWithMarkers');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatter_straight_with_markers (underscore variant)', () => {
      const chart = new Chart('scatter_straight_with_markers');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatterSmooth', () => {
      const chart = new Chart('scatterSmooth');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatter_smooth (underscore variant)', () => {
      const chart = new Chart('scatter_smooth');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatterSmoothWithMarkers', () => {
      const chart = new Chart('scatterSmoothWithMarkers');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept scatter_smooth_with_markers (underscore variant)', () => {
      const chart = new Chart('scatter_smooth_with_markers');
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept stock chart type', () => {
      const chart = new Chart('stock');
      expect(chart).toBeInstanceOf(Chart);
    });
  });

  describe('Chart - Style Variations', () => {
    test('should accept chart style 2', () => {
      const chart = new Chart('column').setStyle(2);
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept chart style 10', () => {
      const chart = new Chart('column').setStyle(10);
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept chart style 20', () => {
      const chart = new Chart('column').setStyle(20);
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept chart style 30', () => {
      const chart = new Chart('column').setStyle(30);
      expect(chart).toBeInstanceOf(Chart);
    });

    test('should accept chart style 40', () => {
      const chart = new Chart('column').setStyle(40);
      expect(chart).toBeInstanceOf(Chart);
    });
  });

  describe('Conditional Formatting - All Icon Set Types', () => {
    test.each([
      ['threeArrowsGray', '3ArrowsGray'],
      ['threeFlags', '3Flags'],
      ['threeTrafficLightsWithRim', '3TrafficLights2'],
      ['threeSigns', '3Signs'],
      ['threeSymbols', '3Symbols2'],
      ['threeSymbolsCircled', '3Symbols'],
      ['threeStars', '3Stars'],
      ['threeTriangles', '3Triangles'],
      ['fourArrowsGray', '4ArrowsGray'],
      ['fourRedToBlack', '4RedToBlack'],
      ['fourTrafficLights', '4TrafficLights'],
      ['fiveArrowsGray', '5ArrowsGray'],
      ['fiveQuadrants', '5Quarters'],
      ['fiveBoxes', '5Boxes'],
    ])('should add %s icon set', (iconType, excelIconSet) => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, iconType);

      const file = path.join(testOutputDir, `icons_${iconType}.xlsx`);
      workbook.save(file);
      expect(readXlsxPart(file, 'xl/worksheets/sheet1.xml')).toContain(`iconSet="${excelIconSet}"`);
    });
  });

  describe('Conditional Formatting - All Date Rules', () => {
    test.each([
      'yesterday',
      'tomorrow',
      'last7Days',
      'lastWeek',
      'nextWeek',
      'lastMonth',
      'thisMonth',
      'nextMonth',
    ])('should add %s date rule', (timePeriod) => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('yellow');
      worksheet.addConditionalFormatDate(0, 0, 9, 0, timePeriod, format);

      const file = path.join(testOutputDir, `date_${timePeriod}.xlsx`);
      workbook.save(file);
      expect(readXlsxPart(file, 'xl/worksheets/sheet1.xml')).toContain(
        `timePeriod="${timePeriod}"`,
      );
    });
  });

  describe('Table - Style Variations', () => {
    test('should accept table style 1 (light)', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const table = new Table().setStyle(1);

      worksheet.write(0, 0, 'Header');
      worksheet.write(1, 0, 'Data');
      worksheet.addTable(0, 0, 1, 0, table);

      const file = path.join(testOutputDir, 'table_style_1.xlsx');
      workbook.save(file);
      expect(readXlsxPart(file, 'xl/tables/table1.xml')).toContain(
        '<tableStyleInfo name="TableStyleLight1"',
      );
    });

    test('should accept table style 5 (light)', () => {
      const table = new Table().setStyle(5);
      expect(table).toBeInstanceOf(Table);
    });

    test('should accept table style 25 (medium)', () => {
      const table = new Table().setStyle(25);
      expect(table).toBeInstanceOf(Table);
    });

    test('should accept table style 35 (medium)', () => {
      const table = new Table().setStyle(35);
      expect(table).toBeInstanceOf(Table);
    });

    test('should accept table style 50 (dark)', () => {
      const table = new Table().setStyle(50);
      expect(table).toBeInstanceOf(Table);
    });

    test('should accept table style 55 (dark)', () => {
      const table = new Table().setStyle(55);
      expect(table).toBeInstanceOf(Table);
    });

    test('should accept table style 60 (dark)', () => {
      const table = new Table().setStyle(60);
      expect(table).toBeInstanceOf(Table);
    });
  });

  describe('Sparkline - Style Variations', () => {
    test('should accept sparkline style 2', () => {
      const sparkline = new Sparkline().setStyle(2);
      expect(sparkline).toBeInstanceOf(Sparkline);
    });

    test('should accept sparkline style 5', () => {
      const sparkline = new Sparkline().setStyle(5);
      expect(sparkline).toBeInstanceOf(Sparkline);
    });

    test('should accept sparkline style 15', () => {
      const sparkline = new Sparkline().setStyle(15);
      expect(sparkline).toBeInstanceOf(Sparkline);
    });

    test('should accept sparkline style 25', () => {
      const sparkline = new Sparkline().setStyle(25);
      expect(sparkline).toBeInstanceOf(Sparkline);
    });

    test('should accept sparkline style 30', () => {
      const sparkline = new Sparkline().setStyle(30);
      expect(sparkline).toBeInstanceOf(Sparkline);
    });

    test('should accept sparkline style 36', () => {
      const sparkline = new Sparkline().setStyle(36);
      expect(sparkline).toBeInstanceOf(Sparkline);
    });
  });

  describe('Page Setup - Paper Size Variations', () => {
    test('should accept paper size 11 (A5)', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();

      expect(() => {
        worksheet.setPaperSize(11);
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'paper_a5.xlsx'));
    });

    test('should accept paper size 3 (Tabloid)', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.setPaperSize(3);

      const file = path.join(testOutputDir, 'paper_tabloid.xlsx');
      workbook.save(file);
      expect(readXlsxPart(file, 'xl/worksheets/sheet1.xml')).toContain('<pageSetup paperSize="3"');
    });

    test('should accept paper size 8 (A3)', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.setPaperSize(8);

      const file = path.join(testOutputDir, 'paper_a3.xlsx');
      workbook.save(file);
      expect(readXlsxPart(file, 'xl/worksheets/sheet1.xml')).toContain('<pageSetup paperSize="8"');
    });
  });
});
