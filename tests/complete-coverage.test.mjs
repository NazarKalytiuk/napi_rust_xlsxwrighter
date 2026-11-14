/**
 * Tests to achieve 100% coverage - Edge cases and enum variations
 */

import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import path from 'node:path';
import { Workbook, Format, Chart, Table, Sparkline } from '../wrapper.ts';
import { cleanupTestOutput, createTestOutputDir, readXlsxPart } from './test-helpers.mjs';

/** The `<tableStyleInfo>` element of a one-column table using `table`. */
function tableStyleInfo(table) {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  worksheet.write(0, 0, 'Header');
  worksheet.write(1, 0, 'Data');
  worksheet.addTable(0, 0, 1, 0, table);
  const xml = readXlsxPart(workbook.saveToBuffer(), 'xl/tables/table1.xml');
  return /<tableStyleInfo [^>]*\/>/.exec(xml)[0];
}

/** The chart XML from `<c:plotArea>` up to the first series' data references. */
function plotAreaStart(chart) {
  const workbook = new Workbook();
  chart.addSeries({ categories: 'Data!$A$1:$A$3', values: 'Data!$B$1:$B$3' });
  workbook.addWorksheet('Data').insertChart(4, 0, chart);
  const xml = readXlsxPart(workbook.saveToBuffer(), 'xl/charts/chart1.xml');
  const plotArea = xml.slice(xml.indexOf('<c:plotArea>')).replace('<c:layout/>', '');
  return plotArea.slice(0, plotArea.search(/<c:(?:cat|xVal|val)>/));
}

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
    // Each type is the plot element plus its grouping, bar direction or
    // style; line-less and marker-less series override the type's default.
    test.each([
      ['area', '<c:areaChart><c:grouping val="standard"/>', ''],
      ['areaStacked', '<c:areaChart><c:grouping val="stacked"/>', ''],
      ['area_stacked', '<c:areaChart><c:grouping val="stacked"/>', ''],
      ['areaPercentStacked', '<c:areaChart><c:grouping val="percentStacked"/>', ''],
      ['area_percent_stacked', '<c:areaChart><c:grouping val="percentStacked"/>', ''],
      ['bar', '<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>', ''],
      ['barStacked', '<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>', ''],
      ['bar_stacked', '<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>', ''],
      [
        'barPercentStacked',
        '<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/>',
        '',
      ],
      [
        'bar_percent_stacked',
        '<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/>',
        '',
      ],
      ['column', '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>', ''],
      ['columnStacked', '<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>', ''],
      ['column_stacked', '<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>', ''],
      [
        'columnPercentStacked',
        '<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>',
        '',
      ],
      [
        'column_percent_stacked',
        '<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>',
        '',
      ],
      ['doughnut', '<c:doughnutChart><c:varyColors val="1"/>', ''],
      [
        'line',
        '<c:lineChart><c:grouping val="standard"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      [
        'lineStacked',
        '<c:lineChart><c:grouping val="stacked"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      [
        'line_stacked',
        '<c:lineChart><c:grouping val="stacked"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      [
        'linePercentStacked',
        '<c:lineChart><c:grouping val="percentStacked"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      [
        'line_percent_stacked',
        '<c:lineChart><c:grouping val="percentStacked"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      ['pie', '<c:pieChart><c:varyColors val="1"/>', ''],
      [
        'radar',
        '<c:radarChart><c:radarStyle val="marker"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      ['radarWithMarkers', '<c:radarChart><c:radarStyle val="marker"/>', ''],
      ['radar_with_markers', '<c:radarChart><c:radarStyle val="marker"/>', ''],
      ['radarFilled', '<c:radarChart><c:radarStyle val="filled"/>', ''],
      ['radar_filled', '<c:radarChart><c:radarStyle val="filled"/>', ''],
      [
        'scatter',
        '<c:scatterChart><c:scatterStyle val="lineMarker"/>',
        '<c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr>',
      ],
      [
        'scatterStraight',
        '<c:scatterChart><c:scatterStyle val="lineMarker"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      [
        'scatter_straight',
        '<c:scatterChart><c:scatterStyle val="lineMarker"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      ['scatterStraightWithMarkers', '<c:scatterChart><c:scatterStyle val="lineMarker"/>', ''],
      ['scatter_straight_with_markers', '<c:scatterChart><c:scatterStyle val="lineMarker"/>', ''],
      [
        'scatterSmooth',
        '<c:scatterChart><c:scatterStyle val="smoothMarker"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      [
        'scatter_smooth',
        '<c:scatterChart><c:scatterStyle val="smoothMarker"/>',
        '<c:marker><c:symbol val="none"/></c:marker>',
      ],
      ['scatterSmoothWithMarkers', '<c:scatterChart><c:scatterStyle val="smoothMarker"/>', ''],
      ['scatter_smooth_with_markers', '<c:scatterChart><c:scatterStyle val="smoothMarker"/>', ''],
      ['stock', '<c:stockChart>', ''],
    ])('should write a %s chart', (chartType, plot, series) => {
      expect(plotAreaStart(new Chart(chartType))).toBe(
        `<c:plotArea>${plot}<c:ser><c:idx val="0"/><c:order val="0"/>${series}`,
      );
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
      ['threeArrows', '3Arrows'],
      ['threeArrowsGray', '3ArrowsGray'],
      ['threeFlags', '3Flags'],
      ['threeTrafficLightsWithRim', '3TrafficLights2'],
      ['threeSigns', '3Signs'],
      ['threeSymbols', '3Symbols2'],
      ['threeSymbolsCircled', '3Symbols'],
      ['threeStars', '3Stars'],
      ['threeTriangles', '3Triangles'],
      ['fourArrows', '4Arrows'],
      ['fourArrowsGray', '4ArrowsGray'],
      ['fourRedToBlack', '4RedToBlack'],
      ['fourHistograms', '4Rating'],
      ['fourTrafficLights', '4TrafficLights'],
      ['fiveArrows', '5Arrows'],
      ['fiveArrowsGray', '5ArrowsGray'],
      ['fiveHistograms', '5Rating'],
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

    test('should leave threeTrafficLights, the Excel default icon set, implicit', () => {
      const workbook = new Workbook();
      workbook.addWorksheet().addConditionalFormatIconSet(0, 0, 9, 0, 'threeTrafficLights');

      const sheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
      expect(sheet).toContain('<cfRule type="iconSet" priority="1"><iconSet><cfvo ');
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
    // Style IDs number Excel's gallery: 0 is "None", then Light 1-21,
    // Medium 1-28 and Dark 1-11.
    test('writes table style 0 without a style name', () => {
      expect(tableStyleInfo(new Table().setStyle(0))).not.toContain('name=');
    });

    test.each([
      [1, 'TableStyleLight1'],
      [2, 'TableStyleLight2'],
      [3, 'TableStyleLight3'],
      [4, 'TableStyleLight4'],
      [5, 'TableStyleLight5'],
      [6, 'TableStyleLight6'],
      [7, 'TableStyleLight7'],
      [8, 'TableStyleLight8'],
      [9, 'TableStyleLight9'],
      [10, 'TableStyleLight10'],
      [11, 'TableStyleLight11'],
      [12, 'TableStyleLight12'],
      [13, 'TableStyleLight13'],
      [14, 'TableStyleLight14'],
      [15, 'TableStyleLight15'],
      [16, 'TableStyleLight16'],
      [17, 'TableStyleLight17'],
      [18, 'TableStyleLight18'],
      [19, 'TableStyleLight19'],
      [20, 'TableStyleLight20'],
      [21, 'TableStyleLight21'],
      [22, 'TableStyleMedium1'],
      [23, 'TableStyleMedium2'],
      [24, 'TableStyleMedium3'],
      [25, 'TableStyleMedium4'],
      [26, 'TableStyleMedium5'],
      [27, 'TableStyleMedium6'],
      [28, 'TableStyleMedium7'],
      [29, 'TableStyleMedium8'],
      [30, 'TableStyleMedium9'],
      [31, 'TableStyleMedium10'],
      [32, 'TableStyleMedium11'],
      [33, 'TableStyleMedium12'],
      [34, 'TableStyleMedium13'],
      [35, 'TableStyleMedium14'],
      [36, 'TableStyleMedium15'],
      [37, 'TableStyleMedium16'],
      [38, 'TableStyleMedium17'],
      [39, 'TableStyleMedium18'],
      [40, 'TableStyleMedium19'],
      [41, 'TableStyleMedium20'],
      [42, 'TableStyleMedium21'],
      [43, 'TableStyleMedium22'],
      [44, 'TableStyleMedium23'],
      [45, 'TableStyleMedium24'],
      [46, 'TableStyleMedium25'],
      [47, 'TableStyleMedium26'],
      [48, 'TableStyleMedium27'],
      [49, 'TableStyleMedium28'],
      [50, 'TableStyleDark1'],
      [51, 'TableStyleDark2'],
      [52, 'TableStyleDark3'],
      [53, 'TableStyleDark4'],
      [54, 'TableStyleDark5'],
      [55, 'TableStyleDark6'],
      [56, 'TableStyleDark7'],
      [57, 'TableStyleDark8'],
      [58, 'TableStyleDark9'],
      [59, 'TableStyleDark10'],
      [60, 'TableStyleDark11'],
    ])('writes table style %i as %s', (styleId, styleName) => {
      expect(tableStyleInfo(new Table().setStyle(styleId))).toContain(
        `<tableStyleInfo name="${styleName}" `,
      );
    });

    test('rejects style IDs outside 0-60', () => {
      expect(() => new Table().setStyle(61)).toThrow(
        'Invalid table style: 61. Use 0 (none) or 1-60',
      );
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
