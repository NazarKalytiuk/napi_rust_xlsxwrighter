/**
 * Shared test utilities
 */

const fs = require('fs');

const transientCleanupErrors = new Set(['EBUSY', 'ENOTEMPTY', 'EPERM']);

/**
 * Safely cleanup test output directory with retry logic for Windows file locking
 * @param {string} testOutputDir - Path to test output directory
 */
async function cleanupTestOutput(testOutputDir) {
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      await fs.promises.rm(testOutputDir, { recursive: true, force: true });
      return;
    } catch (error) {
      if (!transientCleanupErrors.has(error.code)) {
        throw error;
      }

      if (attempt === 4) {
        console.warn(
          `Warning: Could not clean up test output directory (${error.code})`,
        );
        return;
      }

      await new Promise(resolve => setTimeout(resolve, 100 * 2 ** attempt));
    }
  }
}

module.exports = {
  cleanupTestOutput,
};
