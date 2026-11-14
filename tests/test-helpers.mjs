/**
 * Shared test utilities
 */

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import AdmZip from 'adm-zip';

const transientCleanupErrors = new Set(['EBUSY', 'ENOTEMPTY', 'EPERM']);

/**
 * Read one XML part (for example `xl/styles.xml`) from a saved workbook.
 * @param {string | Buffer} xlsx - Path to an .xlsx file or its contents
 * @param {string} part - Path of the part inside the xlsx zip container
 * @returns {string} The part's text
 */
export function readXlsxPart(xlsx, part) {
  const entry = new AdmZip(xlsx).getEntry(part);
  if (!entry) {
    throw new Error(`Workbook has no part ${part}`);
  }
  return entry.getData().toString('utf8');
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
