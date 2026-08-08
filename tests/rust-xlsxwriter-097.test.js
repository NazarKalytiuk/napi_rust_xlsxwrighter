const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const AdmZip = require('adm-zip');
const {
  Format,
  Image,
  Table,
  Workbook,
} = require('../wrapper.js');

function workbookPart(buffer, part) {
  return new AdmZip(buffer).readAsText(part);
}

function firstColumnWidth(buffer) {
  const xml = workbookPart(buffer, 'xl/worksheets/sheet1.xml');
  const match = xml.match(/<col\b[^>]*\bwidth="([0-9.]+)"/);
  if (!match) {
    throw new Error('Worksheet has no explicit column width');
  }
  return Number(match[1]);
}

describe('rust_xlsxwriter 0.97 API', () => {
  test('writes typed font schemes and rejects invalid schemes', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Fonts');
    const body = new Format().setFontName('Aptos').setFontScheme('body');
    const headings = new Format().setFontName('Aptos Display').setFontScheme('headings');
    const noTheme = new Format().setFontName('Arial').setFontScheme('none');

    expect(body.setBold()).toBe(body);
    expect(() => new Format().setFontScheme('invalid')).toThrow('Invalid font scheme');

    worksheet.write(0, 0, 'Body', body);
    worksheet.write(1, 0, 'Heading', headings);
    worksheet.write(2, 0, 'None', noTheme);

    const styles = workbookPart(workbook.saveToBuffer(), 'xl/styles.xml');
    expect(styles).toContain('<scheme val="minor"/>');
    expect(styles).toContain('<scheme val="major"/>');
    const noSchemeFont = styles.match(/<font>(?:(?!<\/font>)[\s\S])*<name val="Arial"\/>(?:(?!<\/font>)[\s\S])*<\/font>/);
    expect(noSchemeFont).not.toBeNull();
    expect(noSchemeFont[0]).not.toContain('<scheme');
  });

  test('uses the Excel 2023 theme before worksheets are added', () => {
    const workbook = new Workbook();
    workbook.useExcel2023Theme();
    workbook.addWorksheet('Theme').write(0, 0, 'Aptos');

    const buffer = workbook.saveToBuffer();
    expect(workbookPart(buffer, 'xl/theme/theme1.xml')).toContain('Aptos');
    expect(workbookPart(buffer, 'xl/styles.xml')).toContain('Aptos Narrow');

    const lateWorkbook = new Workbook();
    lateWorkbook.addWorksheet('Too late');
    expect(() => lateWorkbook.useExcel2023Theme()).toThrow('before adding worksheets');
  });

  test('sets a custom default workbook format before worksheets are added', () => {
    const workbook = new Workbook();
    const format = new Format()
      .setFontName('Arial')
      .setFontSize(11)
      .setFontScheme('none');

    workbook.setDefaultFormat(format, 20, 72);
    workbook.addWorksheet('Default').write(0, 0, 'Arial');

    const buffer = workbook.saveToBuffer();
    expect(workbookPart(buffer, 'xl/styles.xml')).toContain('val="Arial"');

    const lateWorkbook = new Workbook();
    lateWorkbook.addWorksheet('Too late');
    expect(() => lateWorkbook.setDefaultFormat(format, 20, 72)).toThrow('before adding worksheets');
  });


  test('loads a custom XML theme and preserves it in the workbook', () => {
    const themePath = path.join(os.tmpdir(), `rusc-xlsx-theme-${process.pid}-${Date.now()}.xml`);
    const themeXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Rusc Test Theme"></a:theme>';
    fs.writeFileSync(themePath, themeXml);

    try {
      const workbook = new Workbook();
      workbook.useCustomTheme(themePath);
      workbook.addWorksheet('Theme').write(0, 0, 'Custom');

      expect(workbookPart(workbook.saveToBuffer(), 'xl/theme/theme1.xml')).toContain('Rusc Test Theme');
    } finally {
      fs.rmSync(themePath, { force: true });
    }

    expect(() => new Workbook().useCustomTheme(`${themePath}.missing`)).toThrow('Failed to use custom theme');
  });

  test('writes table alternative text metadata', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Table');
    worksheet.writeRow(0, 0, ['Name', 'Value']);
    worksheet.writeRow(1, 0, ['Alpha', 1]);

    const table = new Table()
      .setName('AccessibleTable')
      .setAltTextTitle('Sales table')
      .setAltText('Quarterly sales values');

    worksheet.addTable(0, 0, 1, 1, table);

    const tableXml = workbookPart(workbook.saveToBuffer(), 'xl/tables/table1.xml');
    expect(tableXml).toContain('altText="Sales table"');
    expect(tableXml).toContain('altTextSummary="Quarterly sales values"');
  });

  test('limits autofit width and inspected rows', () => {
    const unlimited = new Workbook();
    const unlimitedSheet = unlimited.addWorksheet('Unlimited');
    unlimitedSheet.write(0, 0, 'Moderately long heading');
    unlimitedSheet.write(1, 0, 'x'.repeat(200));
    unlimitedSheet.autofit();

    const widthLimited = new Workbook();
    const widthLimitedSheet = widthLimited.addWorksheet('Width');
    widthLimitedSheet.write(0, 0, 'Moderately long heading');
    widthLimitedSheet.write(1, 0, 'x'.repeat(200));
    widthLimitedSheet.setAutofitMaxWidth(100);
    widthLimitedSheet.autofit();

    const rowLimited = new Workbook();
    const rowLimitedSheet = rowLimited.addWorksheet('Rows');
    rowLimitedSheet.write(0, 0, 'Moderately long heading');
    rowLimitedSheet.write(1, 0, 'x'.repeat(200));
    rowLimitedSheet.setAutofitMaxRow(1);
    rowLimitedSheet.autofit();

    const unlimitedWidth = firstColumnWidth(unlimited.saveToBuffer());
    expect(firstColumnWidth(widthLimited.saveToBuffer())).toBeLessThan(unlimitedWidth);
    expect(firstColumnWidth(rowLimited.saveToBuffer())).toBeLessThan(unlimitedWidth);
  });

  test('centers a proportionally fitted image in its cell', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Image');
    worksheet.setColumnWidth(0, 40);
    worksheet.setRowHeight(0, 100);

    const image = new Image(path.join(__dirname, 'test-image.png'));
    worksheet.insertImageFitToCellCentered(0, 0, image);

    const drawing = workbookPart(workbook.saveToBuffer(), 'xl/drawings/drawing1.xml');
    const offsets = [...drawing.matchAll(/<xdr:(?:colOff|rowOff)>(\d+)<\/xdr:(?:colOff|rowOff)>/g)]
      .map(([, value]) => Number(value));

    expect(offsets.some(value => value > 0)).toBe(true);
  });
});
