/**
 * DateTime example - demonstrates all datetime writing methods
 */

const { Workbook, Format } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('DateTime');

const headerFmt = new Format().setBold().setBackgroundColor('#4472C4').setFontColor('white');
const dateFmt = new Format().setNumFormat('yyyy-mm-dd');
const dateTimeFmt = new Format().setNumFormat('yyyy-mm-dd hh:mm:ss');
const timeFmt = new Format().setNumFormat('hh:mm:ss AM/PM');

worksheet.setColumnWidth(0, 30);
worksheet.setColumnWidth(1, 25);

worksheet.write(0, 0, 'Method', headerFmt);
worksheet.write(0, 1, 'Result', headerFmt);

let row = 1;

// From JavaScript Date object
worksheet.write(row, 0, 'From JS Date (new Date())');
worksheet.writeDatetime(row, 1, new Date(2025, 0, 15, 9, 30, 0), dateTimeFmt);
row++;

// From Unix timestamp (seconds)
worksheet.write(row, 0, 'From Unix timestamp');
worksheet.writeDatetime(row, 1, 1700000000, dateTimeFmt);
row++;

// From ISO string
worksheet.write(row, 0, 'From ISO string');
worksheet.writeDatetimeFromString(row, 1, '2025-06-15T14:30:00', dateTimeFmt);
row++;

// From year/month/day
worksheet.write(row, 0, 'From year/month/day');
worksheet.writeDatetimeFromYmd(row, 1, 2025, 12, 25, dateFmt);
row++;

// From year/month/day/hour/min/sec
worksheet.write(row, 0, 'From year/month/day/h/m/s');
worksheet.writeDatetimeFromYmdHms(row, 1, 2025, 7, 4, 18, 0, 0, dateTimeFmt);
row++;

// From object with date components
worksheet.write(row, 0, 'From object {year, month, day}');
worksheet.writeDatetime(row, 1, { year: 2025, month: 3, day: 14 }, dateFmt);
row++;

// From object with time components
worksheet.write(row, 0, 'From object with time');
worksheet.writeDatetime(row, 1, {
  year: 2025, month: 11, day: 28,
  hour: 8, min: 15, sec: 30
}, dateTimeFmt);
row++;

// Time-only format
row++;
worksheet.write(row, 0, 'Time-only display:');
worksheet.writeDatetimeFromYmdHms(row, 1, 2025, 1, 1, 14, 45, 30, timeFmt);
row++;

// Different date formats
row++;
worksheet.write(row, 0, 'Custom format: dd/mm/yyyy');
const euFmt = new Format().setNumFormat('dd/mm/yyyy');
worksheet.writeDatetimeFromYmd(row, 1, 2025, 3, 14, euFmt);
row++;

worksheet.write(row, 0, 'Custom format: mmm d, yyyy');
const longFmt = new Format().setNumFormat('mmm d, yyyy');
worksheet.writeDatetimeFromYmd(row, 1, 2025, 3, 14, longFmt);
row++;

worksheet.write(row, 0, 'Custom format: ddd, mmmm dd');
const dayFmt = new Format().setNumFormat('ddd, mmmm dd');
worksheet.writeDatetimeFromYmd(row, 1, 2025, 3, 14, dayFmt);
row++;

workbook.save('datetime.xlsx');
console.log('Created datetime.xlsx with various datetime formats and input methods');
