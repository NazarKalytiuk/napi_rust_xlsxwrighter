/**
 * Tests to achieve 100% coverage - Edge cases and enum variations
 */

const fs = require('fs');
const path = require('path');
const { Workbook, Format, Chart, Table, Sparkline } = require('../wrapper');
const { cleanupTestOutput } = require('./test-helpers');

describe('Complete Coverage - Edge Cases and Enum Variations', () => {
  const testOutputDir = path.join(__dirname, 'output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
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

      const styles = ['dashDot', 'dashDotDot', 'slantDashDot', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot'];
      styles.forEach((style, index) => {
        const format = new Format().setBorderTop(style).setBorderBottom(style);
        worksheet.write(index, 0, style, format);
      });

      workbook.save(path.join(testOutputDir, 'all_border_styles.xlsx'));
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
        'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical',
        'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal',
        'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis',
        'gray125', 'gray0625'
      ];

      patterns.forEach((pattern, index) => {
        const format = new Format()
          .setPattern(pattern)
          .setForegroundColor('#FF0000')
          .setBackgroundColor('#FFFF00');
        worksheet.write(index, 0, pattern, format);
      });

      workbook.save(path.join(testOutputDir, 'all_fill_patterns.xlsx'));
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
    test('should add threeArrowsGray icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();

      expect(() => {
        worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeArrowsGray');
      }).not.toThrow();

      workbook.save(path.join(testOutputDir, 'icons_three_arrows_gray.xlsx'));
    });

    test('should add threeFlags icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeFlags');
      workbook.save(path.join(testOutputDir, 'icons_three_flags.xlsx'));
    });

    test('should add threeTrafficLightsWithRim icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeTrafficLightsWithRim');
      workbook.save(path.join(testOutputDir, 'icons_traffic_rim.xlsx'));
    });

    test('should add threeSigns icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeSigns');
      workbook.save(path.join(testOutputDir, 'icons_three_signs.xlsx'));
    });

    test('should add threeSymbols icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeSymbols');
      workbook.save(path.join(testOutputDir, 'icons_three_symbols.xlsx'));
    });

    test('should add threeSymbolsCircled icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeSymbolsCircled');
      workbook.save(path.join(testOutputDir, 'icons_symbols_circled.xlsx'));
    });

    test('should add threeStars icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeStars');
      workbook.save(path.join(testOutputDir, 'icons_three_stars.xlsx'));
    });

    test('should add threeTriangles icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'threeTriangles');
      workbook.save(path.join(testOutputDir, 'icons_three_triangles.xlsx'));
    });

    test('should add fourArrowsGray icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fourArrowsGray');
      workbook.save(path.join(testOutputDir, 'icons_four_arrows_gray.xlsx'));
    });

    test('should add fourRedToBlack icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fourRedToBlack');
      workbook.save(path.join(testOutputDir, 'icons_red_to_black.xlsx'));
    });

    test('should add fourTrafficLights icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fourTrafficLights');
      workbook.save(path.join(testOutputDir, 'icons_four_traffic.xlsx'));
    });

    test('should add fiveArrowsGray icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fiveArrowsGray');
      workbook.save(path.join(testOutputDir, 'icons_five_arrows_gray.xlsx'));
    });

    test('should add fiveQuadrants icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fiveQuadrants');
      workbook.save(path.join(testOutputDir, 'icons_five_quadrants.xlsx'));
    });

    test('should add fiveBoxes icon set', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.addConditionalFormatIconSet(0, 0, 9, 0, 'fiveBoxes');
      workbook.save(path.join(testOutputDir, 'icons_five_boxes.xlsx'));
    });
  });

  describe('Conditional Formatting - All Date Rules', () => {
    test('should add yesterday date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('yellow');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'yesterday', format);
      workbook.save(path.join(testOutputDir, 'date_yesterday.xlsx'));
    });

    test('should add tomorrow date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('green');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'tomorrow', format);
      workbook.save(path.join(testOutputDir, 'date_tomorrow.xlsx'));
    });

    test('should add last7Days date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('blue');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'last7Days', format);
      workbook.save(path.join(testOutputDir, 'date_last7days.xlsx'));
    });

    test('should add lastWeek date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('purple');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'lastWeek', format);
      workbook.save(path.join(testOutputDir, 'date_lastweek.xlsx'));
    });

    test('should add nextWeek date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('orange');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'nextWeek', format);
      workbook.save(path.join(testOutputDir, 'date_nextweek.xlsx'));
    });

    test('should add lastMonth date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('red');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'lastMonth', format);
      workbook.save(path.join(testOutputDir, 'date_lastmonth.xlsx'));
    });

    test('should add thisMonth date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('cyan');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'thisMonth', format);
      workbook.save(path.join(testOutputDir, 'date_thismonth.xlsx'));
    });

    test('should add nextMonth date rule', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      const format = new Format().setBackgroundColor('lime');

      worksheet.addConditionalFormatDate(0, 0, 9, 0, 'nextMonth', format);
      workbook.save(path.join(testOutputDir, 'date_nextmonth.xlsx'));
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

      workbook.save(path.join(testOutputDir, 'table_style_1.xlsx'));
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
      workbook.save(path.join(testOutputDir, 'paper_tabloid.xlsx'));
    });

    test('should accept paper size 8 (A3)', () => {
      const workbook = new Workbook();
      const worksheet = workbook.addWorksheet();
      worksheet.setPaperSize(8);
      workbook.save(path.join(testOutputDir, 'paper_a3.xlsx'));
    });
  });
});
