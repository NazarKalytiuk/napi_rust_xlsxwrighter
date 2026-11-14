/**
 * Tests for turning off Excel warnings and for ZIP64 (large file) workbooks
 */

import { describe, expect, test } from 'vitest';
import { Workbook } from '../wrapper.ts';
import { readXlsxPart } from './test-helpers.mjs';

/** The `<ignoredErrors>` element of a workbook's first worksheet. */
function ignoredErrorsXml(workbook) {
  const sheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
  return /<ignoredErrors>.*<\/ignoredErrors>/.exec(sheet)?.[0];
}

/** The extra field of the first local file header of a zip file. */
function firstEntryExtraField(zip) {
  expect(zip.readUInt32LE(0)).toBe(0x04034b50);
  const nameLength = zip.readUInt16LE(26);
  const extraLength = zip.readUInt16LE(28);
  return zip.subarray(30 + nameLength, 30 + nameLength + extraLength);
}

describe('Worksheet.ignoreError and ignoreErrorRange', () => {
  test.each([
    ['numberStoredAsText', 'numberStoredAsText'],
    ['formulaError', 'evalError'],
    ['formulaDiffers', 'formula'],
    ['formulaRefersToEmptyCells', 'emptyCellReference'],
    ['formulaOmitsCells', 'formulaRange'],
    ['dataValidationError', 'listDataValidation'],
    ['unlockedCellsWithFormula', 'unlockedFormula'],
    ['inconsistentColumnFormula', 'calculatedColumn'],
  ])('turns off %s warnings with the %s attribute', (errorType, attribute) => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.ignoreError(0, 0, errorType);

    expect(ignoredErrorsXml(workbook)).toBe(
      `<ignoredErrors><ignoredError sqref="A1" ${attribute}="1"/></ignoredErrors>`,
    );
  });

  test('joins the cells and ranges of one warning type', () => {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();

    worksheet.ignoreError(0, 0, 'numberStoredAsText');
    worksheet.ignoreErrorRange(0, 2, 4, 3, 'numberStoredAsText');

    expect(ignoredErrorsXml(workbook)).toBe(
      '<ignoredErrors><ignoredError sqref="A1 C1:D5" numberStoredAsText="1"/></ignoredErrors>',
    );
  });

  test('rejects a second rule for the same cell', () => {
    const worksheet = new Workbook().addWorksheet();
    worksheet.ignoreError(0, 0, 'numberStoredAsText');

    expect(() => worksheet.ignoreError(0, 0, 'formulaError')).toThrow(
      "Failed to ignore error: Parameter error: 'Excel only allows on rule per cell. Cell/Range 'A1' was previously used.'.",
    );
  });

  test('rejects an inverted range', () => {
    const worksheet = new Workbook().addWorksheet();

    expect(() => worksheet.ignoreErrorRange(3, 0, 1, 0, 'formulaError')).toThrow(
      'Failed to ignore error: First row or column in range is greater than last row or column.',
    );
  });

  test('rejects an unknown warning type', () => {
    const worksheet = new Workbook().addWorksheet();

    expect(() => worksheet.ignoreError(0, 0, 'numberAsText')).toThrow(
      'Invalid ignore error type: numberAsText',
    );
  });

  test('rejects the two-digit text year warning that rust_xlsxwriter writes wrongly', () => {
    const worksheet = new Workbook().addWorksheet();

    expect(() => worksheet.ignoreErrorRange(0, 0, 1, 1, 'twoDigitTextYear')).toThrow(
      "Unsupported ignore error type: twoDigitTextYear. rust_xlsxwriter 0.102.0 writes it with an attribute name Excel doesn't define",
    );
  });
});

describe('Workbook.useZipLargeFile', () => {
  test('writes ZIP64 extra fields that zip readers still read', () => {
    const workbook = new Workbook();
    workbook.addWorksheet().write(0, 0, 'large');

    workbook.useZipLargeFile(true);
    const zip = workbook.saveToBuffer();

    const extra = firstEntryExtraField(zip);
    expect(extra.readUInt16LE(0)).toBe(0x0001); // ZIP64 extended information
    expect(readXlsxPart(zip, 'xl/sharedStrings.xml')).toContain('<t>large</t>');
  });

  test('writes no ZIP64 extra fields by default or when turned off', () => {
    const workbook = new Workbook();
    workbook.addWorksheet().write(0, 0, 'small');

    expect(firstEntryExtraField(workbook.saveToBuffer())).toHaveLength(0);
    workbook.useZipLargeFile(true);
    workbook.useZipLargeFile(false);
    expect(firstEntryExtraField(workbook.saveToBuffer())).toHaveLength(0);
  });
});
