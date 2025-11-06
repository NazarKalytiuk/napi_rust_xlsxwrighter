const { Workbook, Chart } = require('./wrapper.js');

// Create a workbook
const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Charts Demo');

// Add data for the chart
worksheet.write(0, 0, 'Month');
worksheet.write(0, 1, 'Sales');
worksheet.write(0, 2, 'Expenses');

const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
const sales = [1000, 1500, 2100, 2800, 3200, 3800];
const expenses = [800, 900, 1200, 1600, 1800, 2000];

for (let i = 0; i < months.length; i++) {
  worksheet.write(i + 1, 0, months[i]);
  worksheet.write(i + 1, 1, sales[i]);
  worksheet.write(i + 1, 2, expenses[i]);
}

// Create a column chart
console.log('Creating column chart...');
const columnChart = new Chart('column');
columnChart.addSeries({
  categories: 'Charts Demo!$A$2:$A$7',
  values: 'Charts Demo!$B$2:$B$7',
  name: 'Sales'
});
columnChart.addSeries({
  categories: 'Charts Demo!$A$2:$A$7',
  values: 'Charts Demo!$C$2:$C$7',
  name: 'Expenses'
});
columnChart.setTitle('Monthly Sales vs Expenses');
columnChart.setXAxisName('Month');
columnChart.setYAxisName('Amount ($)');
columnChart.setStyle(11);

worksheet.insertChartWithOffset(1, 4, columnChart, 25, 10);

// Create a line chart
console.log('Creating line chart...');
const lineChart = new Chart('line');
lineChart.addSeries({
  categories: 'Charts Demo!$A$2:$A$7',
  values: 'Charts Demo!$B$2:$B$7',
  name: 'Sales Trend'
});
lineChart.setTitle('Sales Trend');
lineChart.setXAxisName('Month');
lineChart.setYAxisName('Revenue ($)');
lineChart.setStyle(12);

worksheet.insertChartWithOffset(17, 4, lineChart, 25, 10);

// Create a pie chart
console.log('Creating pie chart...');
const worksheet2 = workbook.addWorksheet('Pie Chart');
worksheet2.write(0, 0, 'Category');
worksheet2.write(0, 1, 'Value');
worksheet2.write(1, 0, 'Product A');
worksheet2.write(1, 1, 35);
worksheet2.write(2, 0, 'Product B');
worksheet2.write(2, 1, 25);
worksheet2.write(3, 0, 'Product C');
worksheet2.write(3, 1, 20);
worksheet2.write(4, 0, 'Product D');
worksheet2.write(4, 1, 15);
worksheet2.write(5, 0, 'Product E');
worksheet2.write(5, 1, 5);

const pieChart = new Chart('pie');
pieChart.addSeries({
  categories: 'Pie Chart!$A$2:$A$6',
  values: 'Pie Chart!$B$2:$B$6',
  name: 'Market Share'
});
pieChart.setTitle('Product Market Share');
pieChart.setStyle(10);

worksheet2.insertChartWithOffset(1, 3, pieChart, 25, 10);

// Save the workbook
console.log('Saving workbook...');
workbook.save('test-chart.xlsx');
console.log('✓ Chart demo saved to test-chart.xlsx');
console.log('  - Column chart: Monthly Sales vs Expenses');
console.log('  - Line chart: Sales Trend');
console.log('  - Pie chart: Product Market Share');
