import { describe, expect, test } from 'vitest';
import { Workbook } from '../wrapper.ts';
import { readXlsxPart } from './test-helpers.mjs';

function workbookWithSheet() {
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet();
  return { workbook, worksheet };
}

describe('rust_xlsxwriter 0.100-0.102 API', () => {
  test.each([
    ['automatic', '<calcPr calcId="124519" fullCalcOnLoad="1"/>'],
    ['manual', '<calcPr calcId="124519" calcMode="manual" calcOnSave="0"/>'],
    ['partial', '<calcPr calcId="124519" calcMode="autoNoTable" fullCalcOnLoad="1"/>'],
  ])('writes the %s calculation mode', (mode, calcPr) => {
    const { workbook } = workbookWithSheet();
    workbook.setCalculationMode(mode);

    expect(readXlsxPart(workbook.saveToBuffer(), 'xl/workbook.xml')).toContain(calcPr);
  });

  test('rejects unknown calculation modes', () => {
    const { workbook } = workbookWithSheet();

    expect(() => workbook.setCalculationMode('sometimes')).toThrow(
      'Invalid calculation mode: sometimes',
    );
  });

  test('writes the worksheet tab ratio and rejects values outside 0-100', () => {
    const { workbook } = workbookWithSheet();
    workbook.setTabRatio(80);

    expect(readXlsxPart(workbook.saveToBuffer(), 'xl/workbook.xml')).toMatch(
      /<workbookView [^>]*tabRatio="800"/,
    );
    expect(() => workbook.setTabRatio(101)).toThrow('Tab ratio must be between 0 and 100, got 101');
  });

  test('hides zero values in worksheet cells', () => {
    const { workbook, worksheet } = workbookWithSheet();
    worksheet.write(0, 0, 0);
    const defaultSheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');

    worksheet.setHiddenZeroes(true);
    const hiddenSheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');

    expect(defaultSheet).not.toContain('showZeros');
    expect(hiddenSheet).toMatch(/<sheetView [^>]*showZeros="0"/);
  });

  test('applies the zip compression level and rejects values outside 1-9', () => {
    const sizeAtLevel = (level) => {
      const { workbook, worksheet } = workbookWithSheet();
      for (let row = 0; row < 2000; row++) {
        worksheet.writeRow(row, 0, ['repeated text value', row, row * 2]);
      }
      workbook.setZipCompressionLevel(level);
      return workbook.saveToBuffer().length;
    };

    expect(sizeAtLevel(9)).toBeLessThan(sizeAtLevel(1));

    const { workbook } = workbookWithSheet();
    expect(() => workbook.setZipCompressionLevel(0)).toThrow(
      'Zip compression level must be between 1 and 9, got 0',
    );
    expect(() => workbook.setZipCompressionLevel(10)).toThrow(
      'Zip compression level must be between 1 and 9, got 10',
    );
  });
});
