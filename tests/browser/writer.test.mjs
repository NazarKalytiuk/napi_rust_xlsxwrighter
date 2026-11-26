/**
 * The browser entry (wrapper.browser.mjs) on the WebAssembly build, in real
 * browsers. The Node.js suite already runs every API against the same `.wasm`
 * (`npm run test:wasm`); these tests cover what differs in a browser: the
 * loader, the missing file system, the Buffer polyfill, memory growth and Web
 * Workers. Workbooks are checked in Node.js with ExcelJS (see the commands in
 * vitest.browser.config.mts).
 */

import { describe, expect, test } from 'vitest';
import { commands } from 'vitest/browser';
import { __napiBindingTarget } from '../../browser.js';
import {
  Button,
  Chart,
  DataValidation,
  DocProperties,
  Format,
  Formula,
  Image,
  Note,
  ProtectionOptions,
  RichText,
  Shape,
  ShapeFont,
  ShapeFormat,
  ShapeSolidFill,
  ShapeText,
  Sparkline,
  Table,
  TableColumn,
  Workbook,
  rustMemoryStats,
} from '../../wrapper.browser.mjs';
import imageUrl from '../test-image.png?url';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

async function imageBytes() {
  const response = await fetch(imageUrl);
  return new Uint8Array(await response.arrayBuffer());
}

describe('Browser entry', () => {
  test('loads the WebAssembly build', () => {
    expect(__napiBindingTarget).toBe('wasm32-wasip1');
    expect(globalThis.process).toBeUndefined();
  });

  test('writes a workbook that uses every part of the API', async () => {
    const workbook = new Workbook();
    workbook.setProperties(new DocProperties().setTitle('Browser report').setAuthor('Vitest'));

    const data = workbook.addWorksheet('Data');
    const bold = new Format().setBold().setBackgroundColor('yellow');
    const money = new Format().setNumFormat('#,##0.00');
    data.writeRow(0, 0, ['Month', 'Sales', 'Active', 'Updated'], bold);
    const updated = new Date(Date.UTC(2024, 0, 15, 12, 30));
    for (let month = 1; month <= 6; month++) {
      data.writeRow(month, 0, [`M${month}`, month * 100, month % 2 === 0, updated]);
      data.write(month, 1, month * 100, money);
    }
    data.writeFormula(7, 1, new Formula('=SUM(B2:B7)').setResult('2100'));
    data.writeUrlWithText(8, 0, 'https://example.com/', 'Example');
    data.writeRichString(
      9,
      0,
      new RichText().addSegment('plain ').addSegment('bold', new Format().setBold()),
    );
    data.mergeRange(10, 0, 10, 3, 'Merged', bold);
    data.setColumnWidth(0, 20);
    data.setFreezePanes(1, 0);
    data.autofilter(0, 0, 6, 3);
    data.insertNote(0, 5, new Note('A note').setAuthor('Vitest'));
    data.addDataValidation(1, 1, 6, 1, new DataValidation().allowWholeNumberBetween(0, 1000));
    data.addConditionalFormatDataBar(1, 1, 6, 1, 'blue');
    data.addSparkline(0, 6, new Sparkline().setRange('Data!B2:B7').setType('column'));

    const chart = new Chart('column');
    chart.addSeries({ categories: 'Data!$A$2:$A$7', values: 'Data!$B$2:$B$7', name: 'Sales' });
    chart.setTitle('Sales');
    data.insertChart(12, 0, chart);
    data.insertImage(12, 8, Image.fromBuffer(await imageBytes()));
    data.insertButton(30, 0, new Button().setCaption('Run'));
    data.insertShape(
      30,
      4,
      Shape.textbox()
        .setText('Textbox')
        .setFormat(new ShapeFormat().setSolidFill(new ShapeSolidFill().setColor('yellow')))
        .setFont(new ShapeFont().setName('Arial').setSize(8).setBold())
        .setTextOptions(new ShapeText().setDirection('rotate270')),
    );

    const tables = workbook.addWorksheet('Table');
    tables.writeRow(0, 0, ['Item', 'Qty']);
    tables.writeRows(1, 0, [
      ['A', 1],
      ['B', 2],
    ]);
    tables.addTable(
      0,
      0,
      3,
      1,
      new Table()
        .setName('Items')
        .setTotalRow(true)
        .setColumns([
          new TableColumn().setHeader('Item'),
          new TableColumn().setHeader('Qty').setTotalFunction('sum'),
        ]),
    );
    tables.protectWithOptions(new ProtectionOptions().setFormatColumns(true));
    workbook.defineName('Sales', '=Data!$B$2:$B$7');

    const buffer = workbook.saveToBuffer();
    const base64 = buffer.toString('base64');
    const { parts, sheets } = await commands.inspectXlsx(base64);

    expect(sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Table']);
    const cells = sheets[0].cells;
    expect(cells.A1.text).toBe('Month');
    expect(cells.B7.text).toBe('600');
    expect(cells.C3.text).toBe('true');
    expect(cells.D2.date).toBe(updated.toISOString());
    expect(cells.B2.numFmt).toBe('#,##0.00');
    expect(cells.B8.formula).toBe('SUM(B2:B7)');
    expect(cells.A9.hyperlink).toBe('https://example.com/');
    expect(cells.A10.text).toBe('plain bold');
    expect(cells.A11.text).toBe('Merged');
    expect(sheets[1].cells.B2.text).toBe('1');

    expect(parts).toEqual(
      expect.arrayContaining([
        'docProps/core.xml',
        'xl/charts/chart1.xml',
        'xl/comments1.xml',
        'xl/media/image1.png',
        'xl/tables/table1.xml',
      ]),
    );
    expect(await commands.readXlsxPart(base64, 'docProps/core.xml')).toContain('Browser report');
    expect(await commands.readXlsxPart(base64, 'xl/workbook.xml')).toContain(
      '<definedName name="Sales">Data!$B$2:$B$7</definedName>',
    );
    const dataSheet = await commands.readXlsxPart(base64, 'xl/worksheets/sheet1.xml');
    for (const element of [
      '<dataValidation ',
      '<conditionalFormatting ',
      'x14:sparkline',
      '<pane ',
    ]) {
      expect(dataSheet).toContain(element);
    }
    expect(await commands.readXlsxPart(base64, 'xl/worksheets/sheet2.xml')).toContain(
      '<sheetProtection ',
    );
    const drawing = await commands.readXlsxPart(base64, 'xl/drawings/drawing1.xml');
    for (const element of [
      '<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>',
      '<a:rPr lang="en-US" sz="800" b="1">',
      '<a:bodyPr vert="vert270" ',
    ]) {
      expect(drawing).toContain(element);
    }
  });

  test('returns a Buffer that browser APIs accept as a Uint8Array', async () => {
    const workbook = new Workbook();
    workbook.addWorksheet().writeString(0, 0, 'Blob');
    const buffer = workbook.saveToBuffer();

    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.readUInt32LE(0)).toBe(0x04034b50); // ZIP local file header "PK\x03\x04"
    const blob = new Blob([buffer], { type: XLSX_MIME });
    expect(blob.size).toBe(buffer.length);
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(new Uint8Array(buffer));
  });

  test('writes workbooks large enough to grow WebAssembly memory', async () => {
    const [before] = rustMemoryStats();
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    const rows = 50_000;
    for (let row = 0; row < rows; row++) {
      worksheet.writeRow(row, 0, [`row ${row}`, row, row * 0.5, row % 2 === 0, `text ${row % 97}`]);
    }
    // The loader starts with 1024 pages (64 MiB), of which the module's stack
    // and static data take 980, so this much heap means the memory has grown.
    const [during] = rustMemoryStats();
    expect(during - before).toBeGreaterThan(16 * 1024 * 1024);

    const { sheets } = await commands.inspectXlsx(workbook.saveToBuffer().toString('base64'));
    expect(sheets[0].rowCount).toBe(rows);
    expect(sheets[0].cells[`A${rows}`].text).toBe(`row ${rows - 1}`);
    expect(sheets[0].cells[`C${rows}`].text).toBe(String((rows - 1) * 0.5));
  });

  test('reports file system operations as errors and keeps working', async () => {
    const workbook = new Workbook();
    workbook.addWorksheet().writeString(0, 0, 'still works');

    expect(() => workbook.save('report.xlsx')).toThrow(/Failed to save workbook/);
    expect(() => new Image('logo.png')).toThrow(/Failed to load image/);
    expect(() => workbook.setTempdir('/tmp')).toThrow(/Failed to set tempdir/);
    expect(() => workbook.addWorksheetWithConstantMemory()).toThrow(
      /Failed to add constant memory worksheet/,
    );
    expect(() => workbook.addWorksheetWithLowMemory()).toThrow(
      /Failed to add low memory worksheet/,
    );
    await expect(workbook.saveToStream({})).rejects.toThrow(/saveToStream\(\) needs Node\.js/);

    const { sheets } = await commands.inspectXlsx(workbook.saveToBuffer().toString('base64'));
    expect(sheets.map((sheet) => sheet.cells.A1.text)).toEqual(['still works']);
  });

  test('reports invalid input as errors without aborting the module', async () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    expect(() => worksheet.setName('bad/name')).toThrow(/cannot contain invalid characters/);
    expect(() => worksheet.writeNumber(2_000_000, 0, 1)).toThrow(/exceeds Excel's allowed limits/);
    expect(() => worksheet.write(0, 0, new Date(Number.NaN))).toThrow(/Invalid Date/);
    worksheet.writeString(0, 0, 'after errors');

    const { sheets } = await commands.inspectXlsx(workbook.saveToBuffer().toString('base64'));
    expect(sheets[0].cells.A1.text).toBe('after errors');
  });

  test('runs in a module Web Worker', async () => {
    const worker = new Worker(new URL('./fixtures/export-worker.mjs', import.meta.url), {
      type: 'module',
    });
    try {
      const rows = [
        ['from', 'worker'],
        [1, 2],
      ];
      const bytes = await new Promise((resolve, reject) => {
        worker.addEventListener('message', ({ data }) => {
          if (data.type === 'ready') {
            // oxlint-disable-next-line unicorn/require-post-message-target-origin -- Worker#postMessage has no target origin
            worker.postMessage({ rows });
          } else if (data.type === 'workbook') {
            resolve(data.bytes);
          } else {
            reject(new Error(data.error));
          }
        });
        worker.addEventListener('error', (event) =>
          reject(event.error ?? new Error(event.message ?? 'The worker failed to load')),
        );
      });

      expect(bytes).toBeInstanceOf(Uint8Array);
      let binary = '';
      for (const byte of bytes) {
        binary += String.fromCharCode(byte);
      }
      const { sheets } = await commands.inspectXlsx(btoa(binary));
      expect(sheets[0].cells.A1.text).toBe('from');
      expect(sheets[0].cells.B2.text).toBe('2');
    } finally {
      worker.terminate();
    }
  });
});
