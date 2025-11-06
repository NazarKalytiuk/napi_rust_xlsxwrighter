// Comprehensive example of conditional formatting and page setup features
const { Workbook, Format } = require('../wrapper.js');

const workbook = new Workbook();

// =============================================================================
// Example 1: 2-Color Scale - Sales Performance
// =============================================================================
const salesSheet = workbook.addWorksheet('Sales Performance');

// Add headers
salesSheet.writeString(0, 0, 'Month');
salesSheet.writeString(0, 1, 'Sales ($)');
salesSheet.setColumnWidth(0, 15);
salesSheet.setColumnWidth(1, 15);

// Write sales data
const salesData = [
  ['January', 45000],
  ['February', 52000],
  ['March', 38000],
  ['April', 61000],
  ['May', 55000],
  ['June', 48000],
  ['July', 71000],
  ['August', 68000],
  ['September', 59000],
  ['October', 44000],
  ['November', 77000],
  ['December', 82000]
];

for (let i = 0; i < salesData.length; i++) {
  salesSheet.writeString(i + 1, 0, salesData[i][0]);
  salesSheet.writeNumber(i + 1, 1, salesData[i][1]);
}

// Apply 2-color scale: red (low) to green (high)
salesSheet.addConditionalFormat2ColorScale(1, 1, 12, 1, 'red', 'green');

// =============================================================================
// Example 2: 3-Color Scale - Temperature Data
// =============================================================================
const tempSheet = workbook.addWorksheet('Temperature Data');

// Add headers
tempSheet.writeString(0, 0, 'City');
tempSheet.writeString(0, 1, 'Temp (°F)');
tempSheet.setColumnWidth(0, 15);
tempSheet.setColumnWidth(1, 15);

// Write temperature data
const tempData = [
  ['New York', 45],
  ['Los Angeles', 72],
  ['Chicago', 32],
  ['Houston', 85],
  ['Phoenix', 95],
  ['Philadelphia', 48],
  ['San Antonio', 78],
  ['San Diego', 68],
  ['Dallas', 81],
  ['San Jose', 65]
];

for (let i = 0; i < tempData.length; i++) {
  tempSheet.writeString(i + 1, 0, tempData[i][0]);
  tempSheet.writeNumber(i + 1, 1, tempData[i][1]);
}

// Apply 3-color scale: blue (cold) -> yellow (mild) -> red (hot)
tempSheet.addConditionalFormat3ColorScale(1, 1, 10, 1, 'blue', 'yellow', 'red');

// =============================================================================
// Example 3: Data Bars - Project Progress
// =============================================================================
const projectSheet = workbook.addWorksheet('Project Progress');

// Add headers
projectSheet.writeString(0, 0, 'Task');
projectSheet.writeString(0, 1, 'Completion %');
projectSheet.setColumnWidth(0, 25);
projectSheet.setColumnWidth(1, 20);

// Write project data
const projectData = [
  ['Requirements Analysis', 100],
  ['Design', 100],
  ['Backend Development', 75],
  ['Frontend Development', 60],
  ['Testing', 30],
  ['Documentation', 45],
  ['Deployment', 0],
  ['User Training', 0]
];

for (let i = 0; i < projectData.length; i++) {
  projectSheet.writeString(i + 1, 0, projectData[i][0]);
  projectSheet.writeNumber(i + 1, 1, projectData[i][1]);
}

// Apply data bars with blue color
projectSheet.addConditionalFormatDataBar(1, 1, 8, 1, 'blue');

// =============================================================================
// Example 4: Cell Rules - Budget Alerts
// =============================================================================
const budgetSheet = workbook.addWorksheet('Budget Alerts');

// Add headers
budgetSheet.writeString(0, 0, 'Department');
budgetSheet.writeString(0, 1, 'Budget');
budgetSheet.writeString(0, 2, 'Spent');
budgetSheet.writeString(0, 3, 'Remaining');
budgetSheet.setColumnWidth(0, 20);
budgetSheet.setColumnWidth(1, 15);
budgetSheet.setColumnWidth(2, 15);
budgetSheet.setColumnWidth(3, 15);

// Write budget data
const budgetData = [
  ['Marketing', 100000, 85000, 15000],
  ['Engineering', 150000, 145000, 5000],
  ['Sales', 80000, 65000, 15000],
  ['HR', 50000, 48000, 2000],
  ['Operations', 120000, 95000, 25000],
  ['Research', 90000, 92000, -2000]
];

for (let i = 0; i < budgetData.length; i++) {
  budgetSheet.writeString(i + 1, 0, budgetData[i][0]);
  budgetSheet.writeNumber(i + 1, 1, budgetData[i][1]);
  budgetSheet.writeNumber(i + 1, 2, budgetData[i][2]);
  budgetSheet.writeNumber(i + 1, 3, budgetData[i][3]);
}

// Highlight remaining budget < 10000 in red
const alertFormat = new Format()
  .setBold()
  .setBackgroundColor('red')
  .setFontColor('white');

budgetSheet.addConditionalFormatCell(1, 3, 6, 3, 'lessThan', 10000, alertFormat);

// Highlight over-budget (negative remaining) with special format
const overBudgetFormat = new Format()
  .setBold()
  .setBackgroundColor('purple')
  .setFontColor('white');

budgetSheet.addConditionalFormatCell(1, 3, 6, 3, 'lessThan', 0, overBudgetFormat);

// =============================================================================
// Example 5: Text Rules - Status Tracking
// =============================================================================
const statusSheet = workbook.addWorksheet('Status Tracking');

// Add headers
statusSheet.writeString(0, 0, 'Task ID');
statusSheet.writeString(0, 1, 'Task Name');
statusSheet.writeString(0, 2, 'Status');
statusSheet.setColumnWidth(0, 10);
statusSheet.setColumnWidth(1, 30);
statusSheet.setColumnWidth(2, 15);

// Write status data
const statusData = [
  ['T001', 'Database Migration', 'Completed'],
  ['T002', 'API Endpoints', 'In Progress'],
  ['T003', 'UI Components', 'Pending'],
  ['T004', 'Authentication', 'Completed'],
  ['T005', 'Error Handling', 'In Progress'],
  ['T006', 'Documentation', 'Pending'],
  ['T007', 'Unit Tests', 'Completed'],
  ['T008', 'Integration Tests', 'Pending']
];

for (let i = 0; i < statusData.length; i++) {
  statusSheet.writeString(i + 1, 0, statusData[i][0]);
  statusSheet.writeString(i + 1, 1, statusData[i][1]);
  statusSheet.writeString(i + 1, 2, statusData[i][2]);
}

// Format completed tasks with green background
const completedFormat = new Format()
  .setBold()
  .setBackgroundColor('green')
  .setFontColor('white');

statusSheet.addConditionalFormatText(1, 2, 8, 2, 'contains', 'Completed', completedFormat);

// Format in-progress tasks with yellow background
const inProgressFormat = new Format()
  .setBold()
  .setBackgroundColor('yellow');

statusSheet.addConditionalFormatText(1, 2, 8, 2, 'contains', 'In Progress', inProgressFormat);

// Format pending tasks with gray background
const pendingFormat = new Format()
  .setBackgroundColor('silver');

statusSheet.addConditionalFormatText(1, 2, 8, 2, 'contains', 'Pending', pendingFormat);

// =============================================================================
// Example 6: Page Setup - Print-Ready Report
// =============================================================================
const reportSheet = workbook.addWorksheet('Quarterly Report');

// Set page orientation to landscape
reportSheet.setLandscape();

// Set margins (left, right, top, bottom, header, footer) in inches
reportSheet.setMargins(0.75, 0.75, 1.0, 1.0, 0.5, 0.5);

// Set paper size (9 = A4)
reportSheet.setPaperSize(9);

// Add header and footer with special codes
// &C = center, &L = left, &R = right
// &P = page number, &N = total pages, &D = date, &T = time
reportSheet.setHeader('&LQuarterly Report&C&D&RPage &P of &N');
reportSheet.setFooter('&LConfidential&CCompany Name&R&T');

// Write report data
reportSheet.writeString(0, 0, 'Quarter');
reportSheet.writeString(0, 1, 'Revenue');
reportSheet.writeString(0, 2, 'Expenses');
reportSheet.writeString(0, 3, 'Profit');
reportSheet.setColumnWidth(0, 15);
reportSheet.setColumnWidth(1, 15);
reportSheet.setColumnWidth(2, 15);
reportSheet.setColumnWidth(3, 15);

const reportData = [
  ['Q1 2024', 250000, 180000, 70000],
  ['Q2 2024', 280000, 195000, 85000],
  ['Q3 2024', 310000, 210000, 100000],
  ['Q4 2024', 340000, 225000, 115000]
];

for (let i = 0; i < reportData.length; i++) {
  reportSheet.writeString(i + 1, 0, reportData[i][0]);
  reportSheet.writeNumber(i + 1, 1, reportData[i][1]);
  reportSheet.writeNumber(i + 1, 2, reportData[i][2]);
  reportSheet.writeNumber(i + 1, 3, reportData[i][3]);
}

// Set print area to only include the data
reportSheet.setPrintArea(0, 0, 4, 3);

// Enable gridlines for printing
reportSheet.setPrintGridlines(true);

// Center content on page
reportSheet.setPrintCenterHorizontally(true);
reportSheet.setPrintCenterVertically(true);

// Add 3-color scale to profit column
reportSheet.addConditionalFormat3ColorScale(1, 3, 4, 3, 'red', 'yellow', 'green');

// =============================================================================
// Example 7: Combined Features - Executive Dashboard
// =============================================================================
const dashboardSheet = workbook.addWorksheet('Executive Dashboard');

// Page setup for professional printing
dashboardSheet.setPortrait();
dashboardSheet.setMargins(1.0, 1.0, 1.5, 1.0, 0.75, 0.5);
dashboardSheet.setHeader('&C&B&18Executive Dashboard');
dashboardSheet.setFooter('&LGenerated: &D &T&RPage &P');
dashboardSheet.setPrintGridlines(false);
dashboardSheet.setPrintCenterHorizontally(true);

// Title
const titleFormat = new Format()
  .setBold()
  .setFontSize(16)
  .setAlign('center');

dashboardSheet.mergeRange(0, 0, 0, 3, 'Key Performance Indicators', titleFormat);
dashboardSheet.setRowHeight(0, 30);

// Add KPI data
dashboardSheet.writeString(2, 0, 'Metric');
dashboardSheet.writeString(2, 1, 'Current');
dashboardSheet.writeString(2, 2, 'Target');
dashboardSheet.writeString(2, 3, '% of Target');

const kpiData = [
  ['Customer Satisfaction', 4.2, 4.5, 93.3],
  ['Response Time (hrs)', 2.5, 2.0, 80.0],
  ['Sales Growth (%)', 15.5, 12.0, 129.2],
  ['Employee Retention (%)', 92.0, 95.0, 96.8],
  ['Cost Reduction (%)', 8.5, 10.0, 85.0]
];

for (let i = 0; i < kpiData.length; i++) {
  dashboardSheet.writeString(i + 3, 0, kpiData[i][0]);
  dashboardSheet.writeNumber(i + 3, 1, kpiData[i][1]);
  dashboardSheet.writeNumber(i + 3, 2, kpiData[i][2]);
  dashboardSheet.writeNumber(i + 3, 3, kpiData[i][3]);
}

// Set column widths
dashboardSheet.setColumnWidth(0, 25);
dashboardSheet.setColumnWidth(1, 12);
dashboardSheet.setColumnWidth(2, 12);
dashboardSheet.setColumnWidth(3, 15);

// Add data bars to "% of Target" column
dashboardSheet.addConditionalFormatDataBar(3, 3, 7, 3, 'green');

// Highlight metrics below 90% target in orange
const warningFormat = new Format()
  .setBold()
  .setBackgroundColor('orange')
  .setFontColor('white');

dashboardSheet.addConditionalFormatCell(3, 3, 7, 3, 'lessThan', 90, warningFormat);

// Highlight metrics above 120% target in bright green
const excellentFormat = new Format()
  .setBold()
  .setBackgroundColor('green')
  .setFontColor('white');

dashboardSheet.addConditionalFormatCell(3, 3, 7, 3, 'greaterThan', 120, excellentFormat);

// =============================================================================
// Save the workbook
// =============================================================================
workbook.save('conditional-formatting-and-page-setup.xlsx');

console.log('✓ Comprehensive example created: conditional-formatting-and-page-setup.xlsx');
console.log('\nFeatures demonstrated:');
console.log('  • 2-color scale conditional formatting');
console.log('  • 3-color scale conditional formatting');
console.log('  • Data bar conditional formatting');
console.log('  • Cell rule conditional formatting (greaterThan, lessThan)');
console.log('  • Text rule conditional formatting (containsText)');
console.log('  • Page orientation (portrait/landscape)');
console.log('  • Page margins');
console.log('  • Paper size');
console.log('  • Headers and footers with special codes');
console.log('  • Print area');
console.log('  • Print gridlines');
console.log('  • Print centering');
console.log('  • Combined features for professional dashboards');
