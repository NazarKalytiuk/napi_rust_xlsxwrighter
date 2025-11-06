const { Workbook, Sparkline } = require('../wrapper.js');
const fs = require('fs');

describe('Sparklines', () => {
  let workbook, worksheet;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('Sparklines');
  });

  afterEach(() => {
    try {
      if (fs.existsSync('test-sparklines.xlsx')) {
        fs.unlinkSync('test-sparklines.xlsx');
      }
    } catch (err) {
      // Ignore cleanup errors
    }
  });

  describe('Sparkline Creation', () => {
    it('should create a basic line sparkline', () => {
      // Add sample data
      worksheet.write(0, 0, -2);
      worksheet.write(0, 1, 2);
      worksheet.write(0, 2, 3);
      worksheet.write(0, 3, -1);
      worksheet.write(0, 4, 0);

      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-sparklines.xlsx')).toBe(true);
    });

    it('should create a column sparkline', () => {
      worksheet.write(0, 0, 30);
      worksheet.write(0, 1, 20);
      worksheet.write(0, 2, 33);
      worksheet.write(0, 3, 20);
      worksheet.write(0, 4, 15);

      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setType('column');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should create a win/lose sparkline', () => {
      worksheet.write(0, 0, 1);
      worksheet.write(0, 1, -1);
      worksheet.write(0, 2, -1);
      worksheet.write(0, 3, 1);
      worksheet.write(0, 4, -1);

      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setType('winLose');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Sparkline Markers', () => {
    beforeEach(() => {
      // Write test data for all marker tests
      worksheet.write(0, 0, -2);
      worksheet.write(0, 1, 2);
      worksheet.write(0, 2, 3);
      worksheet.write(0, 3, -1);
      worksheet.write(0, 4, 0);
    });

    it('should show high point marker', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showHighPoint(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show low point marker', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showLowPoint(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show first point marker', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showFirstPoint(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show last point marker', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showLastPoint(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show negative point markers', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showNegativePoints(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show all markers', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showMarkers(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show axis', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showAxis(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Sparkline Colors', () => {
    beforeEach(() => {
      worksheet.write(0, 0, -2);
      worksheet.write(0, 1, 2);
      worksheet.write(0, 2, 3);
      worksheet.write(0, 3, -1);
      worksheet.write(0, 4, 0);
    });

    it('should set sparkline color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setSparklineColor('#FF0000');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set high point color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setHighPointColor('#00FF00');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set low point color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setLowPointColor('#0000FF');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set first point color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setFirstPointColor('#FFA500');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set last point color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setLastPointColor('#800080');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set negative points color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setNegativePointsColor('#FF00FF');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set markers color', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setMarkersColor('#000000');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Sparkline Styles', () => {
    beforeEach(() => {
      worksheet.write(0, 0, 30);
      worksheet.write(0, 1, 20);
      worksheet.write(0, 2, 33);
      worksheet.write(0, 3, 20);
      worksheet.write(0, 4, 15);
    });

    it('should set style 1', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setType('column').setStyle(1);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set style 12', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setType('column').setStyle(12);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set style 24', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setType('column').setStyle(24);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Sparkline Options', () => {
    beforeEach(() => {
      worksheet.write(0, 0, -2);
      worksheet.write(0, 1, 2);
      worksheet.write(0, 2, 3);
      worksheet.write(0, 3, -1);
      worksheet.write(0, 4, 0);
    });

    it('should set line weight', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setLineWeight(2.5);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set custom max value', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setCustomMax(5);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set custom min value', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setCustomMin(-5);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set group max', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setGroupMax(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set group min', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setGroupMin(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show hidden data', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showHiddenData(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should set right to left', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').setRightToLeft(true);

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show empty cells as gaps', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showEmptyCellsAs('gaps');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show empty cells as zero', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showEmptyCellsAs('zero');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should show empty cells as connected', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E1').showEmptyCellsAs('connected');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Sparkline Groups', () => {
    beforeEach(() => {
      // Row 1 data
      worksheet.write(0, 0, -2);
      worksheet.write(0, 1, 2);
      worksheet.write(0, 2, 3);
      worksheet.write(0, 3, -1);
      worksheet.write(0, 4, 0);

      // Row 2 data
      worksheet.write(1, 0, 30);
      worksheet.write(1, 1, 20);
      worksheet.write(1, 2, 33);
      worksheet.write(1, 3, 20);
      worksheet.write(1, 4, 15);

      // Row 3 data
      worksheet.write(2, 0, 1);
      worksheet.write(2, 1, -1);
      worksheet.write(2, 2, -1);
      worksheet.write(2, 3, 1);
      worksheet.write(2, 4, -1);
    });

    it('should create sparkline group', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E3').setType('line');

      expect(() => {
        worksheet.addSparklineGroup(0, 5, 2, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should create column sparkline group', () => {
      const sparkline = new Sparkline();
      sparkline.setRange('A1:E3').setType('column').setStyle(12);

      expect(() => {
        worksheet.addSparklineGroup(0, 5, 2, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Method Chaining', () => {
    it('should chain sparkline methods', () => {
      worksheet.write(0, 0, -2);
      worksheet.write(0, 1, 2);
      worksheet.write(0, 2, 3);
      worksheet.write(0, 3, -1);
      worksheet.write(0, 4, 0);

      const sparkline = new Sparkline();
      sparkline
        .setRange('A1:E1')
        .setType('line')
        .showMarkers(true)
        .showHighPoint(true)
        .showLowPoint(true)
        .setHighPointColor('#00FF00')
        .setLowPointColor('#FF0000');

      expect(() => {
        worksheet.addSparkline(0, 5, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });

  describe('Real World Scenarios', () => {
    it('should create dashboard with multiple sparklines', () => {
      // Sales data for 3 products across 5 months
      const products = [
        { name: 'Product A', data: [100, 120, 115, 130, 125] },
        { name: 'Product B', data: [80, 85, 90, 95, 100] },
        { name: 'Product C', data: [150, 140, 155, 160, 165] },
      ];

      products.forEach((product, row) => {
        // Write product name
        worksheet.write(row, 0, product.name);

        // Write sales data
        product.data.forEach((value, col) => {
          worksheet.write(row, col + 1, value);
        });

        // Add sparkline
        const sparkline = new Sparkline();
        sparkline
          .setRange(`B${row + 1}:F${row + 1}`)
          .setType('column')
          .setStyle(12);

        worksheet.addSparkline(row, 6, sparkline);
      });

      expect(() => {
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });

    it('should create trend analysis with markers', () => {
      worksheet.write(0, 0, 'Sales');

      const data = [50, 55, 52, 58, 60, 57, 63];
      data.forEach((value, col) => {
        worksheet.write(0, col + 1, value);
      });

      const sparkline = new Sparkline();
      sparkline
        .setRange('B1:H1')
        .setType('line')
        .showMarkers(true)
        .showHighPoint(true)
        .showLowPoint(true)
        .showFirstPoint(true)
        .showLastPoint(true)
        .setHighPointColor('#00FF00')
        .setLowPointColor('#FF0000');

      expect(() => {
        worksheet.addSparkline(0, 8, sparkline);
        workbook.save('test-sparklines.xlsx');
      }).not.toThrow();
    });
  });
});
