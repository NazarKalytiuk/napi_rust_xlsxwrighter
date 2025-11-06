const { Workbook, Table, TableColumn } = require('../wrapper.js');
const fs = require('fs');

describe('Tables', () => {
  let workbook, worksheet;

  beforeEach(() => {
    workbook = new Workbook();
    worksheet = workbook.addWorksheet('Tables');
  });

  afterEach(() => {
    try {
      if (fs.existsSync('test-tables.xlsx')) {
        fs.unlinkSync('test-tables.xlsx');
      }
    } catch (err) {
      // Ignore cleanup errors
    }
  });

  describe('Table Creation', () => {
    it('should create a basic table', () => {
      // Add sample data
      worksheet.write(0, 0, 'Product');
      worksheet.write(0, 1, 'Sales');
      worksheet.write(1, 0, 'Widget A');
      worksheet.write(1, 1, 100);
      worksheet.write(2, 0, 'Widget B');
      worksheet.write(2, 1, 200);

      const table = new Table();
      expect(() => {
        worksheet.addTable(0, 0, 2, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
      expect(fs.existsSync('test-tables.xlsx')).toBe(true);
    });

    it('should create a table with custom name', () => {
      worksheet.write(0, 0, 'Product');
      worksheet.write(0, 1, 'Price');
      worksheet.write(1, 0, 'Item 1');
      worksheet.write(1, 1, 29.99);

      const table = new Table();
      table.setName('PriceTable');

      expect(() => {
        worksheet.addTable(0, 0, 1, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should create a table with custom style', () => {
      worksheet.write(0, 0, 'Name');
      worksheet.write(0, 1, 'Value');
      worksheet.write(1, 0, 'A');
      worksheet.write(1, 1, 10);

      const table = new Table();
      table.setStyle(12); // Medium style 12

      expect(() => {
        worksheet.addTable(0, 0, 1, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });
  });

  describe('Table Formatting', () => {
    it('should set header row', () => {
      worksheet.write(0, 0, 'Header');
      worksheet.write(1, 0, 'Data');

      const table = new Table();
      table.setHeaderRow(true);

      expect(() => {
        worksheet.addTable(0, 0, 1, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should set total row', () => {
      worksheet.write(0, 0, 'Values');
      worksheet.write(1, 0, 100);
      worksheet.write(2, 0, 200);

      const table = new Table();
      table.setTotalRow(true);

      expect(() => {
        worksheet.addTable(0, 0, 2, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should set banded rows', () => {
      worksheet.write(0, 0, 'Item');
      worksheet.write(1, 0, 'A');
      worksheet.write(2, 0, 'B');

      const table = new Table();
      table.setBandedRows(true);

      expect(() => {
        worksheet.addTable(0, 0, 2, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should set banded columns', () => {
      worksheet.write(0, 0, 'A');
      worksheet.write(0, 1, 'B');
      worksheet.write(1, 0, 1);
      worksheet.write(1, 1, 2);

      const table = new Table();
      table.setBandedColumns(true);

      expect(() => {
        worksheet.addTable(0, 0, 1, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should set first column formatting', () => {
      worksheet.write(0, 0, 'Name');
      worksheet.write(0, 1, 'Value');
      worksheet.write(1, 0, 'Total');
      worksheet.write(1, 1, 100);

      const table = new Table();
      table.setFirstColumn(true);

      expect(() => {
        worksheet.addTable(0, 0, 1, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should set last column formatting', () => {
      worksheet.write(0, 0, 'Item');
      worksheet.write(0, 1, 'Status');
      worksheet.write(1, 0, 'A');
      worksheet.write(1, 1, 'Complete');

      const table = new Table();
      table.setLastColumn(true);

      expect(() => {
        worksheet.addTable(0, 0, 1, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });
  });

  describe('Table Columns', () => {
    it('should create table with column headers', () => {
      worksheet.write(0, 0, 'Product');
      worksheet.write(0, 1, 'Price');
      worksheet.write(0, 2, 'Quantity');
      worksheet.write(1, 0, 'Widget');
      worksheet.write(1, 1, 19.99);
      worksheet.write(1, 2, 5);

      const col1 = new TableColumn();
      col1.setHeader('Product');

      const col2 = new TableColumn();
      col2.setHeader('Price');

      const col3 = new TableColumn();
      col3.setHeader('Quantity');

      const table = new Table();
      table.setColumns([col1, col2, col3]);

      expect(() => {
        worksheet.addTable(0, 0, 1, 2, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should add total row with sum function', () => {
      worksheet.write(0, 0, 'Values');
      worksheet.write(1, 0, 100);
      worksheet.write(2, 0, 200);
      worksheet.write(3, 0, 300);

      const col = new TableColumn();
      col.setHeader('Values').setTotalFunction('sum');

      const table = new Table();
      table.setTotalRow(true).setColumns([col]);

      expect(() => {
        worksheet.addTable(0, 0, 3, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should add total row with average function', () => {
      worksheet.write(0, 0, 'Scores');
      worksheet.write(1, 0, 85);
      worksheet.write(2, 0, 92);
      worksheet.write(3, 0, 88);

      const col = new TableColumn();
      col.setHeader('Scores').setTotalFunction('average');

      const table = new Table();
      table.setTotalRow(true).setColumns([col]);

      expect(() => {
        worksheet.addTable(0, 0, 3, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should add total row with count function', () => {
      worksheet.write(0, 0, 'Items');
      worksheet.write(1, 0, 'A');
      worksheet.write(2, 0, 'B');
      worksheet.write(3, 0, 'C');

      const col = new TableColumn();
      col.setHeader('Items').setTotalFunction('count');

      const table = new Table();
      table.setTotalRow(true).setColumns([col]);

      expect(() => {
        worksheet.addTable(0, 0, 3, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should add total row with max function', () => {
      worksheet.write(0, 0, 'Values');
      worksheet.write(1, 0, 50);
      worksheet.write(2, 0, 75);
      worksheet.write(3, 0, 100);

      const col = new TableColumn();
      col.setHeader('Values').setTotalFunction('max');

      const table = new Table();
      table.setTotalRow(true).setColumns([col]);

      expect(() => {
        worksheet.addTable(0, 0, 3, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should add total row with custom label', () => {
      worksheet.write(0, 0, 'Category');
      worksheet.write(1, 0, 'A');
      worksheet.write(2, 0, 'B');

      const col = new TableColumn();
      col.setHeader('Category').setTotalLabel('Total Items');

      const table = new Table();
      table.setTotalRow(true).setColumns([col]);

      expect(() => {
        worksheet.addTable(0, 0, 2, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });
  });

  describe('Method Chaining', () => {
    it('should chain table methods', () => {
      worksheet.write(0, 0, 'Name');
      worksheet.write(0, 1, 'Value');
      worksheet.write(1, 0, 'Item');
      worksheet.write(1, 1, 100);

      const table = new Table();
      table
        .setName('ChainedTable')
        .setStyle(15)
        .setBandedRows(true)
        .setFirstColumn(true);

      expect(() => {
        worksheet.addTable(0, 0, 1, 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should chain column methods', () => {
      worksheet.write(0, 0, 'Sales');
      worksheet.write(1, 0, 1000);
      worksheet.write(2, 0, 2000);

      const col = new TableColumn();
      col
        .setHeader('Sales')
        .setTotalFunction('sum');

      const table = new Table();
      table.setTotalRow(true).setColumns([col]);

      expect(() => {
        worksheet.addTable(0, 0, 2, 0, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });
  });

  describe('Real World Scenarios', () => {
    it('should create sales report table', () => {
      const headers = ['Product', 'Q1', 'Q2', 'Q3', 'Q4'];
      const data = [
        ['Widget A', 1000, 1200, 1100, 1300],
        ['Widget B', 800, 900, 950, 1000],
        ['Widget C', 1500, 1600, 1700, 1800],
      ];

      // Write headers
      headers.forEach((header, col) => worksheet.write(0, col, header));

      // Write data
      data.forEach((row, rowIndex) => {
        row.forEach((value, col) => {
          worksheet.write(rowIndex + 1, col, value);
        });
      });

      // Create columns
      const columns = headers.map((header, index) => {
        const col = new TableColumn();
        col.setHeader(header);
        if (index > 0) {
          col.setTotalFunction('sum');
        } else {
          col.setTotalLabel('Total');
        }
        return col;
      });

      const table = new Table();
      table
        .setName('QuarterlySales')
        .setStyle(9)
        .setTotalRow(true)
        .setColumns(columns);

      expect(() => {
        worksheet.addTable(0, 0, data.length, headers.length - 1, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });

    it('should create employee roster table', () => {
      worksheet.write(0, 0, 'Name');
      worksheet.write(0, 1, 'Department');
      worksheet.write(0, 2, 'Salary');
      worksheet.write(1, 0, 'John Doe');
      worksheet.write(1, 1, 'Engineering');
      worksheet.write(1, 2, 75000);
      worksheet.write(2, 0, 'Jane Smith');
      worksheet.write(2, 1, 'Marketing');
      worksheet.write(2, 2, 65000);

      const table = new Table();
      table
        .setName('Employees')
        .setStyle(21)
        .setBandedRows(true);

      expect(() => {
        worksheet.addTable(0, 0, 2, 2, table);
        workbook.save('test-tables.xlsx');
      }).not.toThrow();
    });
  });
});
