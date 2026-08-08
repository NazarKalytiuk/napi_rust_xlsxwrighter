/**
 * Shared test utilities
 */

const fs = require('fs');

/**
 * Safely cleanup test output directory with retry logic for Windows file locking
 * @param {string} testOutputDir - Path to test output directory
 */
async function cleanupTestOutput(testOutputDir) {
  await fs.promises.rm(testOutputDir, {
    recursive: true,
    force: true,
    maxRetries: 5,
    retryDelay: 100,
  });
}

module.exports = {
  cleanupTestOutput,
};
