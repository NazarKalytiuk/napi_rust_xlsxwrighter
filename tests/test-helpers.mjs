/**
 * Shared test utilities
 */

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import AdmZip from 'adm-zip';

const transientCleanupErrors = new Set(['EBUSY', 'ENOTEMPTY', 'EPERM']);

/**
 * Read the bytes of one part (for example `xl/vbaProject.bin`) from a saved workbook.
 * @param {string | Buffer} xlsx - Path to an .xlsx file or its contents
 * @param {string} part - Path of the part inside the xlsx zip container
 * @returns {Buffer} The part's contents
 */
export function readXlsxEntry(xlsx, part) {
  const entry = new AdmZip(xlsx).getEntry(part);
  if (!entry) {
    throw new Error(`Workbook has no part ${part}`);
  }
  return entry.getData();
}

/**
 * Read one XML part (for example `xl/styles.xml`) from a saved workbook.
 * @param {string | Buffer} xlsx - Path to an .xlsx file or its contents
 * @param {string} part - Path of the part inside the xlsx zip container
 * @returns {string} The part's text
 */
export function readXlsxPart(xlsx, part) {
  return readXlsxEntry(xlsx, part).toString('utf8');
}

/**
 * Resolve the style parts applied to one cell of the first worksheet.
 * @param {string | Buffer} xlsx - Path to an .xlsx file or its contents
 * @param {string} cell - A1-style reference of the cell
 * @returns {{ xf: string, font: string, fill: string, border: string }} The
 *   cell's `<xf>` from `cellXfs` and the `<font>`, `<fill>` and `<border>` it uses
 */
export function readCellStyle(xlsx, cell) {
  const sheet = readXlsxPart(xlsx, 'xl/worksheets/sheet1.xml');
  const cellTag = new RegExp(`<c r="${cell}"[^>]*>`).exec(sheet)?.[0];
  if (!cellTag) {
    throw new Error(`Worksheet has no cell ${cell}`);
  }
  const styles = readXlsxPart(xlsx, 'xl/styles.xml');
  const elements = (list, tag) => {
    const body = new RegExp(`<${list}\\b[^>]*>([\\s\\S]*?)</${list}>`).exec(styles)?.[1] ?? '';
    return [...body.matchAll(new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map(
      ([element]) => element,
    );
  };
  const xf = elements('cellXfs', 'xf')[Number(/\bs="(\d+)"/.exec(cellTag)?.[1] ?? 0)];
  const id = (name) => Number(new RegExp(`\\b${name}="(\\d+)"`).exec(xf)[1]);
  return {
    xf,
    font: elements('fonts', 'font')[id('fontId')],
    fill: elements('fills', 'fill')[id('fillId')],
    border: elements('borders', 'border')[id('borderId')],
  };
}

/**
 * Create a unique output directory for one test suite. Each suite gets its own
 * directory so test files can run in parallel without removing each other's
 * output during cleanup.
 * @returns {string} Absolute path to the created directory
 */
export function createTestOutputDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'rusc-xlsx-test-'));
}

/**
 * Safely cleanup test output directory with retry logic for Windows file locking
 * @param {string} testOutputDir - Path to test output directory
 */
export async function cleanupTestOutput(testOutputDir) {
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      await fs.promises.rm(testOutputDir, { recursive: true, force: true });
      return;
    } catch (error) {
      if (!transientCleanupErrors.has(error.code)) {
        throw error;
      }

      if (attempt === 4) {
        console.warn(`Warning: Could not clean up test output directory (${error.code})`);
        return;
      }

      await new Promise((resolve) => setTimeout(resolve, 100 * 2 ** attempt));
    }
  }
}
