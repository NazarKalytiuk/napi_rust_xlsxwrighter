/**
 * Charts example - demonstrates creating various chart types
 */

const { Workbook, Format, Chart } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();

// --- Column Chart ---
const ws1 = workbook.addWorksheet('Column Chart');

// Write data for the chart
const headerFmt = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');
ws1.write(0, 0, 'Quarter', headerFmt);
ws1.write(0, 1, 'Revenue', headerFmt);
ws1.write(0, 2, 'Expenses', headerFmt);

const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
const revenue = [45000, 52000, 49000, 61000];
const expenses = [32000, 35000, 33000, 38000];

for (let i = 0; i < 4; i++) {
  ws1.write(i + 1, 0, quarters[i]);
  ws1.write(i + 1, 1, revenue[i]);
  ws1.write(i + 1, 2, expenses[i]);
}

const colChart = new Chart('column');
colChart.setTitle('Quarterly Performance');
colChart.setXAxisName('Quarter');
colChart.setYAxisName('Amount ($)');
colChart.setStyle(10);
colChart.addSeries({
  categories: "'Column Chart'!$A$2:$A$5",
  values: "'Column Chart'!$B$2:$B$5",
  name: 'Revenue',
});
colChart.addSeries({
  categories: "'Column Chart'!$A$2:$A$5",
  values: "'Column Chart'!$C$2:$C$5",
  name: 'Expenses',
});

ws1.insertChart(6, 0, colChart);

// --- Line Chart ---
const ws2 = workbook.addWorksheet('Line Chart');

ws2.write(0, 0, 'Month', headerFmt);
ws2.write(0, 1, 'Users', headerFmt);

const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
const users = [1200, 1800, 2400, 3100, 3900, 5200];

for (let i = 0; i < 6; i++) {
  ws2.write(i + 1, 0, months[i]);
  ws2.write(i + 1, 1, users[i]);
}

const lineChart = new Chart('line');
lineChart.setTitle('User Growth');
lineChart.setXAxisName('Month');
lineChart.setYAxisName('Users');
lineChart.addSeries({
  categories: "'Line Chart'!$A$2:$A$7",
  values: "'Line Chart'!$B$2:$B$7",
  name: 'Active Users',
});

ws2.insertChart(8, 0, lineChart);

// --- Pie Chart ---
const ws3 = workbook.addWorksheet('Pie Chart');

ws3.write(0, 0, 'Category', headerFmt);
ws3.write(0, 1, 'Share', headerFmt);

const categories = ['Desktop', 'Mobile', 'Tablet', 'Other'];
const shares = [45, 35, 15, 5];

for (let i = 0; i < 4; i++) {
  ws3.write(i + 1, 0, categories[i]);
  ws3.write(i + 1, 1, shares[i]);
}

const pieChart = new Chart('pie');
pieChart.setTitle('Traffic by Device');
pieChart.addSeries({
  categories: "'Pie Chart'!$A$2:$A$5",
  values: "'Pie Chart'!$B$2:$B$5",
});

ws3.insertChart(6, 0, pieChart);

// --- Bar Chart ---
const ws4 = workbook.addWorksheet('Bar Chart');

ws4.write(0, 0, 'Product', headerFmt);
ws4.write(0, 1, 'Sales', headerFmt);

const products = ['Widget A', 'Widget B', 'Widget C', 'Widget D', 'Widget E'];
const sales = [320, 280, 450, 190, 370];

for (let i = 0; i < 5; i++) {
  ws4.write(i + 1, 0, products[i]);
  ws4.write(i + 1, 1, sales[i]);
}

const barChart = new Chart('bar');
barChart.setTitle('Product Sales');
barChart.setLegendHidden();
barChart.addSeries({
  categories: "'Bar Chart'!$A$2:$A$6",
  values: "'Bar Chart'!$B$2:$B$6",
  name: 'Sales',
});

ws4.insertChart(7, 0, barChart);

// --- Scatter Chart ---
const ws5 = workbook.addWorksheet('Scatter Chart');

ws5.write(0, 0, 'X', headerFmt);
ws5.write(0, 1, 'Y', headerFmt);

for (let i = 0; i < 10; i++) {
  ws5.write(i + 1, 0, i + 1);
  ws5.write(i + 1, 1, (i + 1) * (i + 1));
}

const scatterChart = new Chart('scatter');
scatterChart.setTitle('X vs Y (Quadratic)');
scatterChart.setXAxisName('X');
scatterChart.setYAxisName('Y');
scatterChart.setLegendHidden();
scatterChart.addSeries({
  categories: "'Scatter Chart'!$A$2:$A$11",
  values: "'Scatter Chart'!$B$2:$B$11",
});

ws5.insertChart(12, 0, scatterChart);

workbook.save('charts.xlsx');
console.log('Created charts.xlsx with 5 chart types: column, line, pie, bar, scatter');
