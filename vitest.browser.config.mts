import { fileURLToPath } from 'node:url';

import { playwright } from '@vitest/browser-playwright';
import AdmZip from 'adm-zip';
import ExcelJS from 'exceljs';
import { defineConfig } from 'vitest/config';
import type { BrowserCommand } from 'vitest/node';

/** What {@link inspectXlsx} reports about one cell. */
interface InspectedCell {
  text: string;
  formula: string | null;
  hyperlink: string | null;
  date: string | null;
  numFmt: string | null;
}

/**
 * Opens a workbook the browser wrote with ExcelJS, in Node.js, so browser
 * tests check their output with the same reader the Node.js tests use.
 */
const inspectXlsx: BrowserCommand<[base64: string]> = async (_context, base64) => {
  const bytes = Buffer.from(base64, 'base64');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  return {
    parts: new AdmZip(bytes)
      .getEntries()
      .map((entry) => entry.entryName)
      .toSorted(),
    sheets: workbook.worksheets.map((worksheet) => {
      const cells: Record<string, InspectedCell> = {};
      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          cells[cell.address] = {
            text: cell.text,
            formula: cell.formula ?? null,
            hyperlink: cell.hyperlink ?? null,
            date: cell.value instanceof Date ? cell.value.toISOString() : null,
            numFmt: cell.numFmt ?? null,
          };
        });
      });
      return { name: worksheet.name, rowCount: worksheet.rowCount, cells };
    }),
  };
};

/** Returns one part (for example `xl/styles.xml`) of a workbook the browser wrote. */
const readXlsxPart: BrowserCommand<[base64: string, part: string]> = (_context, base64, part) => {
  const entry = new AdmZip(Buffer.from(base64, 'base64')).getEntry(part);
  if (!entry) {
    throw new Error(`Workbook has no part ${part}`);
  }
  return entry.getData().toString('utf8');
};

// Runs tests/browser against the browser entry (wrapper.browser.mjs) and the
// local WebAssembly build: `npm run build:wasm && npm run build:wrapper` first.
export default defineConfig({
  resolve: {
    alias: [
      {
        // ./browser.js imports the published WebAssembly package; the tests
        // use the loader `npm run build:wasm` writes next to it instead.
        find: /^@nazarkalytiuk\/rusc-xlsx-wasm32-wasip1$/,
        replacement: fileURLToPath(new URL('./rusc_xlsx.wasip1-browser.js', import.meta.url)),
      },
    ],
  },
  test: {
    include: ['tests/browser/**/*.test.mjs'],
    browser: {
      enabled: true,
      headless: true,
      provider: playwright(),
      instances: [{ browser: 'chromium' }, { browser: 'firefox' }, { browser: 'webkit' }],
      commands: { inspectXlsx, readXlsxPart },
    },
  },
});
