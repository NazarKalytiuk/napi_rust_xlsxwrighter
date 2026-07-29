const AdmZip = require('adm-zip');
const {
  Workbook,
  rustMemoryStats,
} = require('../wrapper');

const memoryModes = [
  ['constant memory', 'addWorksheetWithConstantMemory'],
  ['low memory', 'addWorksheetWithLowMemory'],
];

function readWorksheetXml(buffer) {
  const archive = new AdmZip(buffer);
  return archive.readAsText('xl/worksheets/sheet1.xml');
}

function extractRows(xml) {
  return [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)]
    .map(([, id, contents]) => ({ id: Number(id), contents }));
}

describe.each(memoryModes)('Worksheet.setDefaultRowHeight in %s mode', (_mode, addWorksheet) => {
  test('keeps row IDs unique and preserves the final data row', () => {
    const workbook = new Workbook();
    const worksheet = workbook[addWorksheet]('Data');

    worksheet.setDefaultRowHeight(15.001);
    for (let row = 0; row < 3; row++) {
      worksheet.writeRow(row, 0, [`Row ${row}`, 100 + row]);
    }

    const xml = readWorksheetXml(workbook.saveToBuffer());
    const rows = extractRows(xml);
    const rowIds = rows.map(({ id }) => id);

    expect(xml).toContain('<sheetFormatPr defaultRowHeight="15.001" customHeight="1"/>');
    expect(rowIds).toEqual([1, 2, 3]);
    expect(new Set(rowIds).size).toBe(rowIds.length);
    expect(rows.at(-1).contents).toContain('<v>102</v>');
  });

  test('does not retain metadata for each written row', () => {
    const workbook = new Workbook();
    const worksheet = workbook[addWorksheet]('Data');

    worksheet.setDefaultRowHeight(15.001);
    for (let row = 0; row < 10; row++) {
      worksheet.writeNumber(row, 0, row);
    }

    const retainedBytesBefore = rustMemoryStats()[0];
    for (let row = 10; row < 20000; row++) {
      worksheet.writeNumber(row, 0, row);
    }
    const retainedBytesAfter = rustMemoryStats()[0];

    expect(retainedBytesAfter - retainedBytesBefore).toBeLessThan(64 * 1024);
  });
});
