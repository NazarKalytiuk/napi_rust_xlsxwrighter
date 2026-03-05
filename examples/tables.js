/**
 * Tables example - demonstrates Excel tables with styles, totals, and formulas
 */

const { Workbook, Format, Table, TableColumn } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();

// --- Basic Table ---
const ws1 = workbook.addWorksheet('Basic Table');

// Write table data
ws1.write(0, 0, 'Name');
ws1.write(0, 1, 'Department');
ws1.write(0, 2, 'Salary');
ws1.write(1, 0, 'Alice');
ws1.write(1, 1, 'Engineering');
ws1.write(1, 2, 85000);
ws1.write(2, 0, 'Bob');
ws1.write(2, 1, 'Marketing');
ws1.write(2, 2, 72000);
ws1.write(3, 0, 'Carol');
ws1.write(3, 1, 'Engineering');
ws1.write(3, 2, 92000);
ws1.write(4, 0, 'Dave');
ws1.write(4, 1, 'Sales');
ws1.write(4, 2, 68000);

ws1.setColumnWidth(0, 15);
ws1.setColumnWidth(1, 15);
ws1.setColumnWidth(2, 12);

const basicTable = new Table();
basicTable.setName('Employees');
basicTable.setStyle(5); // Light style 5
basicTable.setColumns([
  new TableColumn().setHeader('Name'),
  new TableColumn().setHeader('Department'),
  new TableColumn().setHeader('Salary'),
]);

ws1.addTable(0, 0, 4, 2, basicTable);

// --- Table with Totals ---
const ws2 = workbook.addWorksheet('Table with Totals');

ws2.write(0, 0, 'Product');
ws2.write(0, 1, 'Quantity');
ws2.write(0, 2, 'Price');
ws2.write(0, 3, 'Total');

const products = [
  ['Laptop', 10, 999.99],
  ['Mouse', 50, 29.99],
  ['Keyboard', 30, 79.99],
  ['Monitor', 15, 449.99],
  ['Headset', 25, 149.99],
];

for (let i = 0; i < products.length; i++) {
  ws2.write(i + 1, 0, products[i][0]);
  ws2.write(i + 1, 1, products[i][1]);
  ws2.write(i + 1, 2, products[i][2]);
}
// Total row placeholder (filled by Excel)
ws2.write(6, 0, 'Total');

ws2.setColumnWidth(0, 12);
ws2.setColumnWidth(1, 10);
ws2.setColumnWidth(2, 10);
ws2.setColumnWidth(3, 12);

const totalsTable = new Table();
totalsTable.setName('ProductSales');
totalsTable.setStyle(25); // Medium style 4
totalsTable.setTotalRow(true);
totalsTable.setColumns([
  new TableColumn().setHeader('Product').setTotalLabel('Total'),
  new TableColumn().setHeader('Quantity').setTotalFunction('sum'),
  new TableColumn().setHeader('Price').setTotalFunction('average'),
  new TableColumn().setHeader('Total').setFormula('[@Quantity]*[@Price]').setTotalFunction('sum'),
]);

ws2.addTable(0, 0, 6, 3, totalsTable);

// --- Styled Table with Banding ---
const ws3 = workbook.addWorksheet('Styled Table');

ws3.write(0, 0, 'Region');
ws3.write(0, 1, 'Q1');
ws3.write(0, 2, 'Q2');
ws3.write(0, 3, 'Q3');
ws3.write(0, 4, 'Q4');

const regions = ['North', 'South', 'East', 'West'];
for (let i = 0; i < regions.length; i++) {
  ws3.write(i + 1, 0, regions[i]);
  for (let q = 0; q < 4; q++) {
    ws3.write(i + 1, q + 1, Math.floor(Math.random() * 50000) + 10000);
  }
}

ws3.setColumnWidth(0, 10);

const styledTable = new Table();
styledTable.setName('RegionalSales');
styledTable.setStyle(52); // Dark style 3
styledTable.setBandedRows(true);
styledTable.setBandedColumns(true);
styledTable.setFirstColumn(true);
styledTable.setLastColumn(true);
styledTable.setColumns([
  new TableColumn().setHeader('Region'),
  new TableColumn().setHeader('Q1'),
  new TableColumn().setHeader('Q2'),
  new TableColumn().setHeader('Q3'),
  new TableColumn().setHeader('Q4'),
]);

ws3.addTable(0, 0, 4, 4, styledTable);

workbook.save('tables.xlsx');
console.log('Created tables.xlsx with 3 table examples: basic, totals/formulas, styled');
