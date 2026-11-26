/**
 * Tests for Format functionality
 */

import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { Workbook, Format } from '../wrapper.ts';
import { cleanupTestOutput, createTestOutputDir, readCellStyle } from './test-helpers.mjs';

describe('Format', () => {
  let testOutputDir;

  beforeAll(() => {
    testOutputDir = createTestOutputDir();
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  test('should create a new format', () => {
    const format = new Format();
    expect(format).toBeInstanceOf(Format);
    expect(format._native).toBeDefined();
  });

  test('should set bold', () => {
    const format = new Format();
    const result = format.setBold();
    expect(result).toBe(format); // Check chaining
  });

  test('should set italic', () => {
    const format = new Format();
    const result = format.setItalic();
    expect(result).toBe(format);
  });

  test('should set font size', () => {
    const format = new Format();
    const result = format.setFontSize(14);
    expect(result).toBe(format);
  });

  test('should set font color with named color', () => {
    const format = new Format();
    expect(() => {
      format.setFontColor('red');
    }).not.toThrow();
  });

  test('should set font color with hex color', () => {
    const format = new Format();
    expect(() => {
      format.setFontColor('#FF0000');
    }).not.toThrow();
  });

  test('should set background color', () => {
    const format = new Format();
    expect(() => {
      format.setBackgroundColor('yellow');
    }).not.toThrow();
  });

  test('should set horizontal alignment', () => {
    const format = new Format();
    expect(() => {
      format.setAlign('center');
    }).not.toThrow();
  });

  test('should set vertical alignment', () => {
    const format = new Format();
    expect(() => {
      format.setVerticalAlign('center');
    }).not.toThrow();
  });

  test('should set border', () => {
    const format = new Format();
    expect(() => {
      format.setBorder('thin');
    }).not.toThrow();
  });

  test('should set number format', () => {
    const format = new Format();
    expect(() => {
      format.setNumFormat('#,##0.00');
    }).not.toThrow();
  });

  test('should chain multiple format operations', () => {
    const format = new Format()
      .setBold()
      .setItalic()
      .setFontSize(16)
      .setFontColor('blue')
      .setBackgroundColor('yellow')
      .setAlign('center')
      .setBorder('thin');

    expect(format).toBeInstanceOf(Format);
  });

  test('should write data with format', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    const format = new Format().setBold().setFontColor('red');

    expect(() => {
      worksheet.write(0, 0, 'Formatted Text', format);
      worksheet.write(1, 0, 42, format);
      worksheet.write(2, 0, true, format);
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-formatted.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });

  test('should handle invalid color', () => {
    const format = new Format();
    expect(() => {
      format.setFontColor('invalidcolor123');
    }).toThrow('Invalid color: invalidcolor123');
  });

  test('should handle invalid alignment', () => {
    const format = new Format();
    expect(() => {
      format.setAlign('invalidalign');
    }).toThrow('Invalid alignment: invalidalign');
  });

  test('should apply number formatting', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    const currencyFormat = new Format().setNumFormat('$#,##0.00');
    const percentFormat = new Format().setNumFormat('0.00%');

    expect(() => {
      worksheet.write(0, 0, 1234.56, currencyFormat);
      worksheet.write(1, 0, 0.85, percentFormat);
    }).not.toThrow();

    const filename = path.join(testOutputDir, 'test-number-format.xlsx');
    workbook.save(filename);
    expect(fs.existsSync(filename)).toBe(true);
  });
});

/** Style parts of A1 in a workbook whose only cell uses `format`. */
function styleOf(format) {
  const workbook = new Workbook();
  workbook.addWorksheet().write(0, 0, 'x', format);
  return readCellStyle(workbook.saveToBuffer(), 'A1');
}

describe('Format XLSX output', () => {
  test.each([
    ['left', 'left'],
    ['center', 'center'],
    ['right', 'right'],
    ['fill', 'fill'],
    ['justify', 'justify'],
    ['centerAcross', 'centerContinuous'],
    ['distributed', 'distributed'],
  ])('writes horizontal alignment %s as %s', (align, horizontal) => {
    expect(styleOf(new Format().setAlign(align)).xf).toContain(
      `<alignment horizontal="${horizontal}"/>`,
    );
  });

  test.each(['top', 'center', 'justify', 'distributed'])(
    'writes vertical alignment %s',
    (align) => {
      expect(styleOf(new Format().setVerticalAlign(align)).xf).toContain(
        `<alignment vertical="${align}"/>`,
      );
    },
  );

  test('leaves bottom, the Excel default vertical alignment, implicit', () => {
    const { xf } = styleOf(new Format().setVerticalAlign('bottom'));
    expect(xf).toContain('applyAlignment="1"');
    expect(xf).not.toContain('vertical=');
  });

  test.each([
    ['single', '<u/>'],
    ['double', '<u val="double"/>'],
    ['singleAccounting', '<u val="singleAccounting"/>'],
    ['doubleAccounting', '<u val="doubleAccounting"/>'],
  ])('writes %s underline', (underline, element) => {
    expect(styleOf(new Format().setUnderline(underline)).font).toContain(element);
  });

  test.each([
    ['black', '000000'],
    ['blue', '0000FF'],
    ['brown', '800000'],
    ['cyan', '00FFFF'],
    ['gray', '808080'],
    ['green', '008000'],
    ['lime', '00FF00'],
    ['magenta', 'FF00FF'],
    ['navy', '000080'],
    ['orange', 'FF6600'],
    ['pink', 'FFC0CB'],
    ['purple', '800080'],
    ['red', 'FF0000'],
    ['silver', 'C0C0C0'],
    ['white', 'FFFFFF'],
    ['yellow', 'FFFF00'],
    ['#1E90FF', '1E90FF'],
  ])('writes font color %s as RGB %s', (color, rgb) => {
    expect(styleOf(new Format().setFontColor(color)).font).toContain(`<color rgb="FF${rgb}"/>`);
  });

  test.each([
    'thin',
    'medium',
    'thick',
    'double',
    'dotted',
    'dashed',
    'dashDot',
    'dashDotDot',
    'slantDashDot',
    'mediumDashed',
    'mediumDashDot',
    'mediumDashDotDot',
    'hair',
  ])('writes %s border on all four sides', (style) => {
    const { border } = styleOf(new Format().setBorder(style).setBorderColor('navy'));
    for (const side of ['left', 'right', 'top', 'bottom']) {
      expect(border).toContain(`<${side} style="${style}"><color rgb="FF000080"/></${side}>`);
    }
  });

  test('writes hair borders set side by side and on the diagonal', () => {
    const { border } = styleOf(
      new Format()
        .setBorderTop('hair')
        .setBorderBottom('hair')
        .setBorderLeft('hair')
        .setBorderRight('hair')
        .setBorderDiagonal('hair')
        .setBorderDiagonalType('both'),
    );
    for (const side of ['left', 'right', 'top', 'bottom', 'diagonal']) {
      expect(border).toContain(`<${side} style="hair"><color auto="1"/></${side}>`);
    }
  });

  test.each([
    ['theme:0:0', '<color theme="0"/>'],
    ['theme:0:1', '<color theme="0" tint="-4.9989318521683403E-2"/>'],
    ['theme:1:5', '<color theme="1" tint="4.9989318521683403E-2"/>'],
    ['theme:2:3', '<color theme="2" tint="-0.499984740745262"/>'],
    ['theme:4:2', '<color theme="4" tint="0.59999389629810485"/>'],
    ['theme:9:5', '<color theme="9" tint="-0.499984740745262"/>'],
  ])('writes font color %s as a theme color', (color, element) => {
    expect(styleOf(new Format().setFontColor(color)).font).toContain(element);
  });

  test('writes theme fill and border colors', () => {
    const solid = styleOf(new Format().setBackgroundColor('theme:4:2'));
    expect(solid.fill).toContain('<fgColor theme="4" tint="0.59999389629810485"/>');

    const pattern = styleOf(
      new Format()
        .setPattern('darkGrid')
        .setForegroundColor('theme:5:0')
        .setBackgroundColor('theme:6:1'),
    );
    expect(pattern.fill).toContain('<fgColor theme="5"/>');
    expect(pattern.fill).toContain('<bgColor theme="6" tint="0.79998168889431442"/>');

    const { border } = styleOf(
      new Format()
        .setBorder('thin')
        .setBorderColor('theme:7:0')
        .setBorderTopColor('theme:8:4')
        .setBorderDiagonal('thin')
        .setBorderDiagonalColor('theme:3:0'),
    );
    expect(border).toContain('<left style="thin"><color theme="7"/></left>');
    expect(border).toContain(
      '<top style="thin"><color theme="8" tint="-0.249977111117893"/></top>',
    );
    expect(border).toContain('<diagonal style="thin"><color theme="3"/></diagonal>');
  });

  test.each([
    'theme:10:0',
    'theme:0:6',
    'theme:4',
    'theme:4:2:1',
    'theme:a:0',
    'theme:-1:0',
    'theme::',
  ])('rejects the theme color %s', (color) => {
    expect(() => new Format().setFontColor(color)).toThrow(
      `Invalid color: ${color}. Theme colors are 'theme:<color>:<shade>' with a color of 0-9 and a shade of 0-5`,
    );
  });

  test.each(['#12345', '#1234567', '#GGGGGG', '#+1+2+3', '12345F'])(
    'rejects the hex color %s',
    (color) => {
      expect(() => new Format().setBackgroundColor(color)).toThrow(`Invalid color: ${color}`);
    },
  );

  test('rejects an invalid diagonal border color', () => {
    expect(() => new Format().setBorderDiagonalColor('nocolor')).toThrow('Invalid color: nocolor');
  });

  test('writes shrink-to-fit and hidden protection', () => {
    const { xf } = styleOf(new Format().setAlign('center').setShrink().setHidden());
    expect(xf).toContain('<alignment horizontal="center" shrinkToFit="1"/>');
    expect(xf).toContain('<protection hidden="1"/>');
  });
});
