/**
 * Tests for column and row settings, outline groups and formatting cells after
 * they are written
 */

import { describe, expect, test } from 'vitest';
import { Format, Workbook } from '../wrapper.ts';
import { readCellStyle, readXlsxPart } from './test-helpers.mjs';

/** The XML of a workbook's first worksheet. */
function sheetXml(workbook) {
  return readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
}

/** The opening `<row>` tag of a 1-based row number. */
function rowTag(xml, row) {
  return new RegExp(`<row r="${row}"[^>]*>`).exec(xml)?.[0];
}

/** The opening `<c>` tag of a cell. */
function cellTag(xml, cell) {
  return new RegExp(`<c r="${cell}"[^>]*>`).exec(xml)?.[0];
}

describe('Column width, format and visibility', () => {
  test('sets column widths in pixels and for column ranges', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setColumnWidthPixels(0, 100);
    worksheet.setColumnRangeWidth(1, 2, 20);
    worksheet.setColumnRangeWidthPixels(3, 4, 50);

    const xml = sheetXml(workbook);
    expect(xml).toContain('<col min="1" max="1" width="14.28515625" customWidth="1"/>');
    expect(xml).toContain('<col min="2" max="3" width="20.7109375" customWidth="1"/>');
    expect(xml).toContain('<col min="4" max="5" width="7.140625" customWidth="1"/>');
  });

  test('marks an auto-fit width as best fit, unlike a manual width', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setColumnAutofitWidth(0, 80);
    worksheet.setColumnWidthPixels(1, 80);

    const xml = sheetXml(workbook);
    expect(xml).toContain('<col min="1" max="1" width="11.42578125" bestFit="1" customWidth="1"/>');
    expect(xml).toContain('<col min="2" max="2" width="11.42578125" customWidth="1"/>');
  });

  test('applies a column format to cells without a format of their own', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setColumnFormat(0, new Format().setNumFormat('0.0%'));
    worksheet.write(0, 0, 0.5);
    worksheet.write(1, 0, 0.25, new Format().setBold());

    const buffer = workbook.saveToBuffer();
    expect(readXlsxPart(buffer, 'xl/worksheets/sheet1.xml')).toContain(
      '<col min="1" max="1" width="9.140625" style="1"/>',
    );
    expect(readCellStyle(buffer, 'A1').xf).toContain('numFmtId="164"');
    expect(readXlsxPart(buffer, 'xl/styles.xml')).toContain(
      '<numFmt numFmtId="164" formatCode="0.0%"/>',
    );
    expect(readCellStyle(buffer, 'A2').font).toContain('<b/>');
  });

  test('applies a format to a column range', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setColumnRangeFormat(1, 2, new Format().setItalic());

    expect(sheetXml(workbook)).toContain('<col min="2" max="3" width="9.140625" style="1"/>');
  });

  test('hides a column and a column range', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setColumnHidden(0);
    worksheet.setColumnRangeHidden(2, 3);

    const xml = sheetXml(workbook);
    expect(xml).toContain('<col min="1" max="1" width="0" hidden="1" customWidth="1"/>');
    expect(xml).toContain('<col min="3" max="4" width="0" hidden="1" customWidth="1"/>');
  });

  test('rejects a column range whose first column is after its last', () => {
    const worksheet = new Workbook().addWorksheet();

    expect(() => worksheet.setColumnRangeWidth(3, 1, 10)).toThrow(
      /^Failed to set column range width: /,
    );
    expect(() => worksheet.setColumnRangeHidden(3, 1)).toThrow(/^Failed to hide column range: /);
  });
});

describe('Row height, format and visibility', () => {
  test('sets a row height in pixels', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setRowHeightPixels(0, 40);
    worksheet.write(0, 0, 'tall');

    expect(rowTag(sheetXml(workbook), 1)).toBe('<row r="1" spans="1:1" ht="30" customHeight="1">');
  });

  test('sets the default row height in pixels', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setDefaultRowHeightPixels(30);

    expect(sheetXml(workbook)).toContain(
      '<sheetFormatPr defaultRowHeight="22.5" customHeight="1"/>',
    );
  });

  test('applies a row format to cells without a format of their own', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setRowFormat(0, new Format().setBold());
    worksheet.write(0, 0, 'header');

    const buffer = workbook.saveToBuffer();
    expect(rowTag(readXlsxPart(buffer, 'xl/worksheets/sheet1.xml'), 1)).toBe(
      '<row r="1" spans="1:1" s="1" customFormat="1">',
    );
    expect(readCellStyle(buffer, 'A1').font).toContain('<b/>');
  });

  test('hides a row and unhides it again', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setRowHidden(0);
    worksheet.setRowHidden(1);
    worksheet.setRowUnhidden(1);
    worksheet.writeRows(0, 0, [['hidden'], ['shown']]);

    const xml = sheetXml(workbook);
    expect(rowTag(xml, 1)).toBe('<row r="1" spans="1:1" hidden="1">');
    expect(rowTag(xml, 2)).toBe('<row r="2" spans="1:1">');
  });

  test('hides the unused rows with one worksheet setting', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.hideUnusedRows(true);
    worksheet.write(0, 0, 'used');

    expect(sheetXml(workbook)).toContain('<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/>');
  });
});

describe('Outline groups', () => {
  test('groups rows into nested outline levels', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.groupRows(1, 4);
    worksheet.groupRows(2, 3);

    const xml = sheetXml(workbook);
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"/>');
    expect(rowTag(xml, 2)).toBe('<row r="2" outlineLevel="1"/>');
    expect(rowTag(xml, 3)).toBe('<row r="3" outlineLevel="2"/>');
    expect(rowTag(xml, 5)).toBe('<row r="5" outlineLevel="1"/>');
  });

  test('groups rows collapsed, hiding them', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.groupRowsCollapsed(1, 2);

    const xml = sheetXml(workbook);
    expect(rowTag(xml, 2)).toBe('<row r="2" hidden="1" outlineLevel="1"/>');
    expect(rowTag(xml, 3)).toBe('<row r="3" hidden="1" outlineLevel="1"/>');
    expect(rowTag(xml, 4)).toBe('<row r="4" collapsed="1"/>');
  });

  test('groups columns, open and collapsed', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.groupColumns(1, 2);
    worksheet.groupColumnsCollapsed(4, 5);

    const xml = sheetXml(workbook);
    expect(xml).toContain('<col min="2" max="3" width="9.140625" outlineLevel="1"/>');
    expect(xml).toContain(
      '<col min="5" max="6" width="0" hidden="1" customWidth="1" outlineLevel="1"/>',
    );
    expect(xml).toContain('<col min="7" max="7" width="9.140625" collapsed="1"/>');
  });

  test('places the outline symbols above and to the left', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.groupRows(1, 2);
    worksheet.groupSymbolsAbove(true);
    worksheet.groupSymbolsToLeft(true);

    expect(sheetXml(workbook)).toContain(
      '<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>',
    );
  });

  test('rejects groups nested deeper than 7', () => {
    const worksheet = new Workbook().addWorksheet();
    for (let level = 1; level <= 7; level++) {
      worksheet.groupRows(level, 20 - level);
    }

    expect(() => worksheet.groupRows(8, 12)).toThrow(
      "Failed to group rows: Worksheet group level exceeds Excel's limit of 8 levels.",
    );
  });
});

describe('Formatting cells after they are written', () => {
  test('sets the format of a written cell and of an empty cell', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.write(0, 0, 'data');
    worksheet.setCellFormat(0, 0, new Format().setBold());
    worksheet.setCellFormat(2, 2, new Format().setItalic());

    const buffer = workbook.saveToBuffer();
    const xml = readXlsxPart(buffer, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A1" s="1" t="s"><v>0</v></c>');
    expect(readCellStyle(buffer, 'A1').font).toContain('<b/>');
    expect(xml).toContain('<c r="C3" s="2"/>');
    expect(readCellStyle(buffer, 'C3').font).toContain('<i/>');
  });

  test('sets the format of a range, keeping its data', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.writeRows(0, 0, [
      ['a', 1],
      ['b', 2],
    ]);
    worksheet.setRangeFormat(0, 0, 0, 1, new Format().setBold());

    const buffer = workbook.saveToBuffer();
    const xml = readXlsxPart(buffer, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="A1" s="1" t="s"><v>0</v></c>');
    expect(xml).toContain('<c r="B1" s="1"><v>1.0</v></c>');
    expect(readCellStyle(buffer, 'B1').font).toContain('<b/>');
    expect(cellTag(xml, 'A2')).toBe('<c r="A2" t="s">');
  });

  test('draws the border format around the outside of a range', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.setRangeFormatWithBorder(
      0,
      0,
      2,
      2,
      new Format().setItalic(),
      new Format().setBorder('thick'),
    );

    const buffer = workbook.saveToBuffer();
    const corner = readCellStyle(buffer, 'A1');
    expect(corner.font).toContain('<i/>');
    expect(corner.border).toContain('<left style="thick">');
    expect(corner.border).toContain('<top style="thick">');
    expect(corner.border).not.toContain('<right style="thick">');
    const centre = readCellStyle(buffer, 'B2');
    expect(centre.font).toContain('<i/>');
    expect(centre.border).toBe('<border><left/><right/><top/><bottom/><diagonal/></border>');
  });

  test('clears a cell, or only its format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.write(0, 0, 'removed', new Format().setBold());
    worksheet.write(1, 0, 'kept', new Format().setBold());
    worksheet.clearCell(0, 0);
    worksheet.clearCellFormat(1, 0);

    const xml = sheetXml(workbook);
    expect(cellTag(xml, 'A1')).toBeUndefined();
    expect(xml).toContain('<c r="A2" t="s"><v>0</v></c>');
  });

  test('formats only the row being written in constant memory mode', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheetWithConstantMemory();
    const bold = new Format().setBold();

    worksheet.write(0, 0, 'current');
    worksheet.setCellFormat(0, 0, bold);
    worksheet.write(1, 0, 'next');
    worksheet.setCellFormat(0, 1, bold);

    const xml = sheetXml(workbook);
    expect(xml).toContain(
      '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>current</t></is></c></row>',
    );
    expect(cellTag(xml, 'B1')).toBeUndefined();
  });
});
