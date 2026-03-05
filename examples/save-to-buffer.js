/**
 * Save to Buffer example - demonstrates saving to a Buffer for HTTP responses,
 * streams, or any scenario where you don't want to write to disk.
 */

const { Workbook, Format } = require('@nazarkalytiuk/rusc-xlsx');
const fs = require('fs');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Data');

const headerFmt = new Format().setBold();

worksheet.write(0, 0, 'Name', headerFmt);
worksheet.write(0, 1, 'Score', headerFmt);
worksheet.write(1, 0, 'Alice');
worksheet.write(1, 1, 95);
worksheet.write(2, 0, 'Bob');
worksheet.write(2, 1, 87);

// Save to a Node.js Buffer instead of a file
const buffer = workbook.saveToBuffer();

console.log('Buffer type:', Buffer.isBuffer(buffer));
console.log('Buffer size:', buffer.length, 'bytes');

// You can do anything with this buffer:

// 1. Write to disk manually
fs.writeFileSync('buffer-output.xlsx', buffer);
console.log('Wrote buffer to buffer-output.xlsx');

// 2. In Express/Koa, send as HTTP response:
//   res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//   res.setHeader('Content-Disposition', 'attachment; filename="report.xlsx"');
//   res.send(buffer);

// 3. Upload to S3/cloud storage:
//   await s3.putObject({ Bucket: 'reports', Key: 'report.xlsx', Body: buffer });

// 4. Convert to base64 for email attachment:
const base64 = buffer.toString('base64');
console.log('Base64 length:', base64.length, 'chars (first 50:', base64.substring(0, 50) + '...)');
