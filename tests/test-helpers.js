/**
 * Shared test utilities
 */

const fs = require('fs');

/**
 * Safely cleanup test output directory with retry logic for Windows file locking
 * @param {string} testOutputDir - Path to test output directory
 */
async function cleanupTestOutput(testOutputDir) {
  // Wait a bit for file handles to be released (especially on Windows)
  await new Promise(resolve => setTimeout(resolve, 100));

  // Retry cleanup with exponential backoff (Windows file locking)
  if (fs.existsSync(testOutputDir)) {
    let retries = 5;
    let delay = 100;

    while (retries > 0) {
      try {
        fs.rmSync(testOutputDir, { recursive: true, force: true });
        break;
      } catch (err) {
        if (err.code === 'EBUSY' && retries > 1) {
          await new Promise(resolve => setTimeout(resolve, delay));
          delay *= 2; // Exponential backoff
          retries--;
        } else if (err.code === 'EBUSY') {
          // Last retry failed, but don't fail the tests
          console.warn('Warning: Could not clean up test output directory (files still locked)');
          break;
        } else {
          throw err;
        }
      }
    }
  }
}

module.exports = {
  cleanupTestOutput,
};
