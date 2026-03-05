/**
 * Document Properties example - demonstrates setting workbook metadata
 */

const { Workbook, Format, DocProperties } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();

// Set document properties (visible in File > Properties in Excel)
const props = new DocProperties()
  .setAuthor('Jane Smith')
  .setTitle('Q4 2025 Sales Report')
  .setSubject('Quarterly Sales Analysis')
  .setCompany('Acme Corp')
  .setManager('John Director')
  .setCategory('Reports')
  .setKeywords('sales, quarterly, 2025, Q4')
  .setComment('Auto-generated report from CRM data')
  .setStatus('Final')
  .setHyperlinkBase('https://intranet.acme.com/reports/');

workbook.setProperties(props);

// Add some content
const worksheet = workbook.addWorksheet('Sales Report');
worksheet.setColumnWidth(0, 15);
worksheet.setColumnWidth(1, 15);

const headerFmt = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');

worksheet.write(0, 0, 'Region', headerFmt);
worksheet.write(0, 1, 'Revenue', headerFmt);
worksheet.write(1, 0, 'North');
worksheet.write(1, 1, 245000);
worksheet.write(2, 0, 'South');
worksheet.write(2, 1, 198000);
worksheet.write(3, 0, 'East');
worksheet.write(3, 1, 312000);
worksheet.write(4, 0, 'West');
worksheet.write(4, 1, 276000);

workbook.save('doc-properties.xlsx');
console.log('Created doc-properties.xlsx');
console.log('Open in Excel and check File > Properties to see the metadata');
