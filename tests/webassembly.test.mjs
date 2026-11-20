/**
 * Behaviour that must not depend on whether the native addon or the
 * WebAssembly build is loaded. `npm test` runs this file against the native
 * addon and `npm run test:wasm` against the WebAssembly build.
 */

import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { __napiBindingTarget } from '../index.js';
import { Image, Workbook } from '../wrapper.ts';
import { cleanupTestOutput, createTestOutputDir, readXlsxPart } from './test-helpers.mjs';

describe('Native and WebAssembly builds', () => {
  let testOutputDir;

  beforeAll(() => {
    testOutputDir = createTestOutputDir();
  });

  afterAll(() => cleanupTestOutput(testOutputDir));

  test('loads the build the test run asks for', () => {
    expect(__napiBindingTarget).toBe(process.env.NAPI_RS_WASI_FLAVOR ?? 'native');
  });

  test('resolves relative paths against the working directory', () => {
    // WASI resolves relative paths against `/`; the wrapper resolves them first.
    const absolute = path.join(testOutputDir, 'relative.xlsx');
    const relative = path.relative(process.cwd(), absolute);
    expect(path.isAbsolute(relative)).toBe(false);

    const workbook = new Workbook();
    workbook.addWorksheet().writeString(0, 0, 'relative');
    workbook.save(relative);

    expect(readXlsxPart(absolute, 'xl/sharedStrings.xml')).toContain('relative');
  });

  test('reads images from a relative path and from a plain Uint8Array', () => {
    const imagePath = path.relative(
      process.cwd(),
      path.join(import.meta.dirname, 'test-image.png'),
    );
    const bytes = new Uint8Array(fs.readFileSync(imagePath));
    expect(Buffer.isBuffer(bytes)).toBe(false);

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet();
    worksheet.insertImage(0, 0, new Image(imagePath));
    worksheet.insertImage(10, 0, Image.fromBuffer(bytes));
    const buffer = workbook.saveToBuffer();

    expect(Buffer.isBuffer(buffer)).toBe(true);
    expect(readXlsxPart(buffer, 'xl/drawings/drawing1.xml').match(/<xdr:pic>/g)).toHaveLength(2);
  });

  test.each(['constant', 'low'])(
    'creates %s memory worksheets in the default temporary directory',
    (mode) => {
      const workbook = new Workbook();
      const worksheet =
        mode === 'constant'
          ? workbook.addWorksheetWithConstantMemory()
          : workbook.addWorksheetWithLowMemory();
      for (let row = 0; row < 100; row++) {
        worksheet.writeString(row, 0, `row ${row}`);
      }

      const sheet = readXlsxPart(workbook.saveToBuffer(), 'xl/worksheets/sheet1.xml');
      expect(sheet.match(/<row /g)).toHaveLength(100);
    },
  );
});
