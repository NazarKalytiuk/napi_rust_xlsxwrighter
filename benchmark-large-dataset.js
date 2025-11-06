/**
 * Large Dataset Performance Benchmark
 *
 * This script demonstrates constant memory mode performance
 * with 900,000 rows x 50 columns (~45 million cells)
 *
 * Measures:
 * - Execution time
 * - Memory usage (RSS, Heap Used, External)
 * - File size
 *
 * Run: node benchmark-large-dataset.js
 */

const { Workbook } = require('./wrapper.js');
const fs = require('fs');
const path = require('path');

// Configuration
const ROWS = 900000;
const COLUMNS = 50;
const OUTPUT_FILE = 'benchmark-900k-rows.xlsx';
const TEMP_DIR = path.join(__dirname, 'temp');

// Ensure temp directory exists
if (!fs.existsSync(TEMP_DIR)) {
  fs.mkdirSync(TEMP_DIR, { recursive: true });
}

// Helper function to format bytes
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// Helper function to format time
function formatTime(ms) {
  if (ms < 1000) return `${ms}ms`;
  const seconds = Math.floor(ms / 1000);
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = seconds % 60;
  if (minutes === 0) return `${seconds}s`;
  return `${minutes}m ${remainingSeconds}s`;
}

// Helper function to get memory stats
function getMemoryStats() {
  const usage = process.memoryUsage();
  return {
    rss: usage.rss,
    heapUsed: usage.heapUsed,
    heapTotal: usage.heapTotal,
    external: usage.external,
  };
}

// Helper function to format memory stats
function formatMemoryStats(stats) {
  return {
    rss: formatBytes(stats.rss),
    heapUsed: formatBytes(stats.heapUsed),
    heapTotal: formatBytes(stats.heapTotal),
    external: formatBytes(stats.external),
  };
}

// Main benchmark function
async function runBenchmark() {
  console.log('='.repeat(70));
  console.log('Large Dataset Performance Benchmark');
  console.log('='.repeat(70));
  console.log(`Rows:    ${ROWS.toLocaleString()}`);
  console.log(`Columns: ${COLUMNS}`);
  console.log(`Cells:   ${(ROWS * COLUMNS).toLocaleString()}`);
  console.log(`Mode:    Constant Memory`);
  console.log('='.repeat(70));
  console.log('');

  // Initial memory
  const initialMemory = getMemoryStats();
  console.log('Initial Memory:');
  console.log(formatMemoryStats(initialMemory));
  console.log('');

  // Force garbage collection if available
  if (global.gc) {
    console.log('Running garbage collection...');
    global.gc();
  }

  const startTime = Date.now();
  let lastProgressTime = startTime;
  let progressInterval;

  try {
    // Create workbook
    console.log('Creating workbook...');
    const workbook = new Workbook();

    // Set custom temp directory
    workbook.setTempdir(TEMP_DIR);
    console.log(`Temp directory: ${TEMP_DIR}`);
    console.log('');

    // Add worksheet with constant memory mode
    console.log('Adding worksheet in constant memory mode...');
    const worksheet = workbook.addWorksheetWithConstantMemory('Data');
    console.log('');

    // Write header row
    console.log('Writing header row...');
    for (let col = 0; col < COLUMNS; col++) {
      worksheet.write(0, col, `Column ${col + 1}`);
    }
    console.log('');

    // Start progress monitoring
    let rowsWritten = 1;
    console.log('Writing data rows...');
    console.log('Progress: (each dot = 10,000 rows)');

    progressInterval = setInterval(() => {
      const elapsed = Date.now() - startTime;
      const rowsPerSecond = Math.round(rowsWritten / (elapsed / 1000));
      const currentMem = getMemoryStats();
      process.stdout.write(`\r[${rowsWritten.toLocaleString()}/${ROWS.toLocaleString()}] ` +
        `${rowsPerSecond.toLocaleString()} rows/s | ` +
        `Heap: ${formatBytes(currentMem.heapUsed)} | ` +
        `RSS: ${formatBytes(currentMem.rss)}     `);
    }, 500);

    // Write data rows
    for (let row = 1; row <= ROWS; row++) {
      rowsWritten = row;

      for (let col = 0; col < COLUMNS; col++) {
        // Alternate between strings and numbers for realistic data
        if (col % 2 === 0) {
          worksheet.write(row, col, `Row ${row} Col ${col}`);
        } else {
          worksheet.write(row, col, row * col * 0.123);
        }
      }

      // Show dots every 10,000 rows
      if (row % 10000 === 0) {
        const now = Date.now();
        if (now - lastProgressTime > 100) { // Throttle dot output
          process.stdout.write('.');
          lastProgressTime = now;
        }
      }
    }

    clearInterval(progressInterval);
    console.log('\n');

    const writeTime = Date.now() - startTime;
    console.log(`Data written in: ${formatTime(writeTime)}`);

    // Memory after writing
    const afterWriteMemory = getMemoryStats();
    console.log('Memory after writing:');
    console.log(formatMemoryStats(afterWriteMemory));
    console.log('');

    // Save workbook
    console.log('Saving workbook...');
    const saveStartTime = Date.now();
    workbook.save(OUTPUT_FILE);
    const saveTime = Date.now() - saveStartTime;
    console.log(`Saved in: ${formatTime(saveTime)}`);
    console.log('');

    // Get file size
    const stats = fs.statSync(OUTPUT_FILE);
    const fileSize = stats.size;

    // Final memory
    const finalMemory = getMemoryStats();

    // Calculate memory deltas
    const memoryIncrease = {
      rss: finalMemory.rss - initialMemory.rss,
      heapUsed: finalMemory.heapUsed - initialMemory.heapUsed,
      external: finalMemory.external - initialMemory.external,
    };

    // Total time
    const totalTime = Date.now() - startTime;

    // Print results
    console.log('='.repeat(70));
    console.log('RESULTS');
    console.log('='.repeat(70));
    console.log('');

    console.log('Performance:');
    console.log(`  Total Time:        ${formatTime(totalTime)}`);
    console.log(`  Write Time:        ${formatTime(writeTime)}`);
    console.log(`  Save Time:         ${formatTime(saveTime)}`);
    console.log(`  Rows/second:       ${Math.round(ROWS / (totalTime / 1000)).toLocaleString()}`);
    console.log(`  Cells/second:      ${Math.round((ROWS * COLUMNS) / (totalTime / 1000)).toLocaleString()}`);
    console.log('');

    console.log('Memory Usage:');
    console.log(`  Initial RSS:       ${formatBytes(initialMemory.rss)}`);
    console.log(`  Final RSS:         ${formatBytes(finalMemory.rss)}`);
    console.log(`  RSS Increase:      ${formatBytes(memoryIncrease.rss)}`);
    console.log(`  Peak Heap Used:    ${formatBytes(finalMemory.heapUsed)}`);
    console.log(`  Heap Increase:     ${formatBytes(memoryIncrease.heapUsed)}`);
    console.log(`  External Memory:   ${formatBytes(finalMemory.external)}`);
    console.log('');

    console.log('Output:');
    console.log(`  File:              ${OUTPUT_FILE}`);
    console.log(`  File Size:         ${formatBytes(fileSize)}`);
    console.log(`  Bytes/row:         ${Math.round(fileSize / ROWS)} bytes`);
    console.log('');

    console.log('Efficiency:');
    console.log(`  Memory/Row:        ${formatBytes(memoryIncrease.rss / ROWS)}`);
    console.log(`  Memory/Cell:       ${formatBytes(memoryIncrease.rss / (ROWS * COLUMNS))}`);
    console.log('');

    console.log('='.repeat(70));
    console.log('✅ Benchmark completed successfully!');
    console.log('='.repeat(70));

  } catch (error) {
    clearInterval(progressInterval);
    console.error('');
    console.error('❌ Error during benchmark:');
    console.error(error);
    process.exit(1);
  } finally {
    // Cleanup temp directory
    console.log('');
    console.log('Cleaning up temp files...');
    if (fs.existsSync(TEMP_DIR)) {
      const tempFiles = fs.readdirSync(TEMP_DIR);
      for (const file of tempFiles) {
        fs.unlinkSync(path.join(TEMP_DIR, file));
      }
      fs.rmdirSync(TEMP_DIR);
    }
    console.log('Done!');
  }
}

// Run benchmark
console.log('');
console.log('Starting benchmark...');
console.log('(Run with --expose-gc flag for accurate memory measurements)');
console.log('Example: node --expose-gc benchmark-large-dataset.js');
console.log('');

runBenchmark().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
