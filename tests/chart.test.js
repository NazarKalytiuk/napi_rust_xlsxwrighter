const { Workbook, Chart } = require('../wrapper.js');
const fs = require('fs');
const path = require('path');

describe('Chart', () => {
  let workbook;
  let worksheet;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('Data');

    // Add sample data
    worksheet.write(0, 0, 'Category');
    worksheet.write(0, 1, 'Values');
    worksheet.write(1, 0, 'A');
    worksheet.write(1, 1, 10);
    worksheet.write(2, 0, 'B');
    worksheet.write(2, 1, 20);
    worksheet.write(3, 0, 'C');
    worksheet.write(3, 1, 30);
  });

  describe('Constructor', () => {
    it('should create a column chart', () => {
      const chart = new Chart('column');
      expect(chart).toBeDefined();
      expect(chart._native).toBeDefined();
    });

    it('should create a line chart', () => {
      const chart = new Chart('line');
      expect(chart).toBeDefined();
    });

    it('should create a pie chart', () => {
      const chart = new Chart('pie');
      expect(chart).toBeDefined();
    });

    it('should create a bar chart', () => {
      const chart = new Chart('bar');
      expect(chart).toBeDefined();
    });

    it('should create an area chart', () => {
      const chart = new Chart('area');
      expect(chart).toBeDefined();
    });

    it('should create a scatter chart', () => {
      const chart = new Chart('scatter');
      expect(chart).toBeDefined();
    });

    it('should create a doughnut chart', () => {
      const chart = new Chart('doughnut');
      expect(chart).toBeDefined();
    });

    it('should create a radar chart', () => {
      const chart = new Chart('radar');
      expect(chart).toBeDefined();
    });

    it('should throw error for invalid chart type', () => {
      expect(() => {
        new Chart('invalid_type');
      }).toThrow();
    });
  });

  describe('addSeries', () => {
    it('should add series with categories and values', () => {
      const chart = new Chart('column');
      expect(() => {
        chart.addSeries({
          categories: 'Data!$A$2:$A$4',
          values: 'Data!$B$2:$B$4',
          name: 'Series 1'
        });
      }).not.toThrow();
    });

    it('should add series with only values', () => {
      const chart = new Chart('line');
      expect(() => {
        chart.addSeries({
          values: 'Data!$B$2:$B$4'
        });
      }).not.toThrow();
    });

    it('should add multiple series', () => {
      const chart = new Chart('column');
      expect(() => {
        chart.addSeries({
          categories: 'Data!$A$2:$A$4',
          values: 'Data!$B$2:$B$4',
          name: 'Series 1'
        });
        chart.addSeries({
          categories: 'Data!$A$2:$A$4',
          values: 'Data!$C$2:$C$4',
          name: 'Series 2'
        });
      }).not.toThrow();
    });
  });

  describe('setTitle', () => {
    it('should set chart title', () => {
      const chart = new Chart('column');
      const result = chart.setTitle('Test Chart');
      expect(result).toBe(chart); // Check method chaining
    });

    it('should support method chaining', () => {
      const chart = new Chart('line');
      expect(() => {
        chart.setTitle('Chart Title')
          .setXAxisName('X Axis')
          .setYAxisName('Y Axis');
      }).not.toThrow();
    });
  });

  describe('setXAxisName', () => {
    it('should set X-axis name', () => {
      const chart = new Chart('column');
      const result = chart.setXAxisName('X Axis');
      expect(result).toBe(chart);
    });
  });

  describe('setYAxisName', () => {
    it('should set Y-axis name', () => {
      const chart = new Chart('column');
      const result = chart.setYAxisName('Y Axis');
      expect(result).toBe(chart);
    });
  });

  describe('setStyle', () => {
    it('should set chart style', () => {
      const chart = new Chart('column');
      const result = chart.setStyle(11);
      expect(result).toBe(chart);
    });

    it('should accept styles from 1 to 48', () => {
      const chart = new Chart('column');
      expect(() => chart.setStyle(1)).not.toThrow();
      expect(() => chart.setStyle(48)).not.toThrow();
    });

    it('should throw error for style < 1', () => {
      const chart = new Chart('column');
      expect(() => chart.setStyle(0)).toThrow();
    });

    it('should throw error for style > 48', () => {
      const chart = new Chart('column');
      expect(() => chart.setStyle(49)).toThrow();
    });
  });

  describe('setLegendHidden', () => {
    it('should hide legend', () => {
      const chart = new Chart('column');
      const result = chart.setLegendHidden();
      expect(result).toBe(chart);
    });
  });

  describe('insertChart', () => {
    it('should insert chart into worksheet', () => {
      const chart = new Chart('column');
      chart.addSeries({
        categories: 'Data!$A$2:$A$4',
        values: 'Data!$B$2:$B$4'
      });

      expect(() => {
        worksheet.insertChart(1, 3, chart);
      }).not.toThrow();
    });

    it('should create valid Excel file with chart', () => {
      const chart = new Chart('column');
      chart.addSeries({
        categories: 'Data!$A$2:$A$4',
        values: 'Data!$B$2:$B$4',
        name: 'Test Series'
      });
      chart.setTitle('Test Chart');

      worksheet.insertChart(1, 3, chart);

      const filename = path.join(__dirname, 'temp-chart-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      // Cleanup
      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('insertChartWithOffset', () => {
    it('should insert chart with pixel offsets', () => {
      const chart = new Chart('line');
      chart.addSeries({
        values: 'Data!$B$2:$B$4'
      });

      expect(() => {
        worksheet.insertChartWithOffset(1, 3, chart, 25, 10);
      }).not.toThrow();
    });

    it('should create valid Excel file with positioned chart', () => {
      const chart = new Chart('pie');
      chart.addSeries({
        categories: 'Data!$A$2:$A$4',
        values: 'Data!$B$2:$B$4'
      });
      chart.setTitle('Pie Chart');

      worksheet.insertChartWithOffset(1, 3, chart, 50, 20);

      const filename = path.join(__dirname, 'temp-chart-offset-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      // Cleanup
      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Complex Chart Scenarios', () => {
    it('should create column chart with multiple series', () => {
      worksheet.write(0, 2, 'Series 2');
      worksheet.write(1, 2, 15);
      worksheet.write(2, 2, 25);
      worksheet.write(3, 2, 35);

      const chart = new Chart('column');
      chart.addSeries({
        categories: 'Data!$A$2:$A$4',
        values: 'Data!$B$2:$B$4',
        name: 'Series 1'
      });
      chart.addSeries({
        categories: 'Data!$A$2:$A$4',
        values: 'Data!$C$2:$C$4',
        name: 'Series 2'
      });
      chart.setTitle('Multi-Series Chart');
      chart.setXAxisName('Categories');
      chart.setYAxisName('Values');
      chart.setStyle(12);

      worksheet.insertChart(1, 4, chart);

      const filename = path.join(__dirname, 'temp-multi-series-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      // Cleanup
      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });

    it('should create stacked column chart', () => {
      const chart = new Chart('columnStacked');
      chart.addSeries({
        categories: 'Data!$A$2:$A$4',
        values: 'Data!$B$2:$B$4',
        name: 'Series 1'
      });
      chart.setTitle('Stacked Column Chart');

      worksheet.insertChart(1, 3, chart);

      const filename = path.join(__dirname, 'temp-stacked-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      // Cleanup
      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });

    it('should create chart with custom configuration', () => {
      const chart = new Chart('line');
      chart
        .addSeries({
          categories: 'Data!$A$2:$A$4',
          values: 'Data!$B$2:$B$4',
          name: 'Trend'
        });

      chart
        .setTitle('Sales Trend Analysis')
        .setXAxisName('Month')
        .setYAxisName('Revenue ($)')
        .setStyle(15)
        .setLegendHidden();

      worksheet.insertChartWithOffset(1, 3, chart, 30, 15);

      const filename = path.join(__dirname, 'temp-custom-test.xlsx');
      expect(() => workbook.save(filename)).not.toThrow();
      expect(fs.existsSync(filename)).toBe(true);

      // Cleanup
      if (fs.existsSync(filename)) {
        fs.unlinkSync(filename);
      }
    });
  });

  describe('Chart Types', () => {
    const chartTypes = [
      'area', 'bar', 'column', 'line', 'pie', 'scatter',
      'doughnut', 'radar', 'areaStacked', 'barStacked',
      'columnStacked', 'lineStacked'
    ];

    chartTypes.forEach(type => {
      it(`should create ${type} chart`, () => {
        const chart = new Chart(type);
        chart.addSeries({
          categories: 'Data!$A$2:$A$4',
          values: 'Data!$B$2:$B$4'
        });

        worksheet.insertChart(1, 3, chart);

        const filename = path.join(__dirname, `temp-${type}-test.xlsx`);
        expect(() => workbook.save(filename)).not.toThrow();
        expect(fs.existsSync(filename)).toBe(true);

        // Cleanup
        if (fs.existsSync(filename)) {
          fs.unlinkSync(filename);
        }
      });
    });
  });
});
