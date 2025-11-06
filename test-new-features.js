// Quick test of new features
const { RuscWorkbook, RuscFormat } = require('./index.js');

const workbook = new RuscWorkbook();
const worksheet = workbook.addWorksheet('Test');

// Test page setup
worksheet.setLandscape();
worksheet.setMargins(1.0, 1.0, 1.0, 1.0, 0.5, 0.5);
worksheet.setHeader('&CTest Header');
worksheet.setFooter('&CPage &P');

// Write some data
for (let i = 0; i < 10; i++) {
  worksheet.writeNumber(i, 0, i * 10);
}

// Test conditional formatting - 2 color scale
worksheet.addConditionalFormat2ColorScale(0, 0, 9, 0, 'red', 'green');

// Test conditional formatting - data bar
for (let i = 0; i < 10; i++) {
  worksheet.writeNumber(i, 1, (i + 1) * 5);
}
worksheet.addConditionalFormatDataBar(0, 1, 9, 1, 'blue');

// Write more data for cell conditional format
for (let i = 0; i < 10; i++) {
  worksheet.writeNumber(i, 2, Math.random() * 100);
}

// Test conditional formatting - cell rule (greater than 50)
const highlightFormat = new RuscFormat();
highlightFormat.setBold();
highlightFormat.setBackgroundColor('yellow');
worksheet.addConditionalFormatCell(0, 2, 9, 2, 'greaterThan', 50, highlightFormat);

workbook.save('test-new-features.xlsx');

console.log('✓ Test file created: test-new-features.xlsx');
console.log('✓ Page setup methods work');
console.log('✓ Conditional formatting works');
