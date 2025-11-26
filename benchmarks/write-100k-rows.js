/**
 * Performance Benchmark: Writing 100,000 rows
 *
 * Compares rusc-xlsx vs xlsx (SheetJS) vs exceljs
 */

const { Workbook, Format } = require('..');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const NUM_ROWS = 100000;
const NUM_COLS = 6;

// Test data generator
function generateRowData(row) {
  const products = ['Widget A', 'Widget B', 'Widget C', 'Gadget X', 'Gadget Y'];
  const quantity = Math.floor(Math.random() * 100) + 1;
  const price = Math.random() * 100 + 10;
  return [
    row,
    products[row % products.length],
    quantity,
    price,
    quantity * price,
    `2024-01-${String((row % 28) + 1).padStart(2, '0')}`
  ];
}

// Memory usage helper
function getMemoryUsage() {
  const used = process.memoryUsage();
  return {
    heapUsed: Math.round(used.heapUsed / 1024 / 1024 * 100) / 100,
    heapTotal: Math.round(used.heapTotal / 1024 / 1024 * 100) / 100,
    rss: Math.round(used.rss / 1024 / 1024 * 100) / 100
  };
}

// Cleanup helper
function cleanup(filename) {
  try {
    if (fs.existsSync(filename)) {
      fs.unlinkSync(filename);
    }
  } catch (e) {
    // ignore
  }
}

// Force garbage collection if available
function forceGC() {
  if (global.gc) {
    global.gc();
  }
}

// ============================================
// rusc-xlsx (Constant Memory Mode)
// ============================================
async function benchmarkRuscXlsx() {
  forceGC();
  const filename = 'benchmark-rusc-xlsx.xlsx';
  const startMem = getMemoryUsage();
  const startTime = Date.now();

  const workbook = new Workbook();
  const worksheet = workbook.addWorksheetWithConstantMemory('Data');

  // Headers
  const headers = ['ID', 'Product', 'Quantity', 'Price', 'Total', 'Date'];
  headers.forEach((header, col) => {
    worksheet.write(0, col, header);
  });

  // Data rows
  for (let row = 1; row <= NUM_ROWS; row++) {
    const data = generateRowData(row);
    data.forEach((value, col) => {
      worksheet.write(row, col, value);
    });
  }

  workbook.save(filename);

  const endTime = Date.now();
  const endMem = getMemoryUsage();
  const fileSize = fs.statSync(filename).size;

  cleanup(filename);

  return {
    library: 'rusc-xlsx (constant memory)',
    time: endTime - startTime,
    startHeap: startMem.heapUsed,
    endHeap: endMem.heapUsed,
    peakRss: endMem.rss,
    fileSize: Math.round(fileSize / 1024 / 1024 * 100) / 100
  };
}

// ============================================
// rusc-xlsx (Normal Mode)
// ============================================
async function benchmarkRuscXlsxNormal() {
  forceGC();
  const filename = 'benchmark-rusc-xlsx-normal.xlsx';
  const startMem = getMemoryUsage();
  const startTime = Date.now();

  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet('Data');

  // Headers
  const headers = ['ID', 'Product', 'Quantity', 'Price', 'Total', 'Date'];
  headers.forEach((header, col) => {
    worksheet.write(0, col, header);
  });

  // Data rows
  for (let row = 1; row <= NUM_ROWS; row++) {
    const data = generateRowData(row);
    data.forEach((value, col) => {
      worksheet.write(row, col, value);
    });
  }

  workbook.save(filename);

  const endTime = Date.now();
  const endMem = getMemoryUsage();
  const fileSize = fs.statSync(filename).size;

  cleanup(filename);

  return {
    library: 'rusc-xlsx (normal)',
    time: endTime - startTime,
    startHeap: startMem.heapUsed,
    endHeap: endMem.heapUsed,
    peakRss: endMem.rss,
    fileSize: Math.round(fileSize / 1024 / 1024 * 100) / 100
  };
}

// ============================================
// xlsx (SheetJS)
// ============================================
async function benchmarkXlsx() {
  forceGC();
  const filename = 'benchmark-xlsx.xlsx';
  const startMem = getMemoryUsage();
  const startTime = Date.now();

  // Build data array
  const data = [['ID', 'Product', 'Quantity', 'Price', 'Total', 'Date']];
  for (let row = 1; row <= NUM_ROWS; row++) {
    data.push(generateRowData(row));
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  XLSX.writeFile(wb, filename);

  const endTime = Date.now();
  const endMem = getMemoryUsage();
  const fileSize = fs.statSync(filename).size;

  cleanup(filename);

  return {
    library: 'xlsx (SheetJS)',
    time: endTime - startTime,
    startHeap: startMem.heapUsed,
    endHeap: endMem.heapUsed,
    peakRss: endMem.rss,
    fileSize: Math.round(fileSize / 1024 / 1024 * 100) / 100
  };
}

// ============================================
// exceljs
// ============================================
async function benchmarkExcelJS() {
  forceGC();
  const filename = 'benchmark-exceljs.xlsx';
  const startMem = getMemoryUsage();
  const startTime = Date.now();

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');

  // Headers
  worksheet.addRow(['ID', 'Product', 'Quantity', 'Price', 'Total', 'Date']);

  // Data rows
  for (let row = 1; row <= NUM_ROWS; row++) {
    worksheet.addRow(generateRowData(row));
  }

  await workbook.xlsx.writeFile(filename);

  const endTime = Date.now();
  const endMem = getMemoryUsage();
  const fileSize = fs.statSync(filename).size;

  cleanup(filename);

  return {
    library: 'exceljs',
    time: endTime - startTime,
    startHeap: startMem.heapUsed,
    endHeap: endMem.heapUsed,
    peakRss: endMem.rss,
    fileSize: Math.round(fileSize / 1024 / 1024 * 100) / 100
  };
}

// ============================================
// exceljs (streaming)
// ============================================
async function benchmarkExcelJSStreaming() {
  forceGC();
  const filename = 'benchmark-exceljs-stream.xlsx';
  const startMem = getMemoryUsage();
  const startTime = Date.now();

  const options = {
    filename: filename,
    useStyles: false,
    useSharedStrings: false
  };
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
  const worksheet = workbook.addWorksheet('Data');

  // Headers
  worksheet.addRow(['ID', 'Product', 'Quantity', 'Price', 'Total', 'Date']).commit();

  // Data rows
  for (let row = 1; row <= NUM_ROWS; row++) {
    worksheet.addRow(generateRowData(row)).commit();
  }

  await worksheet.commit();
  await workbook.commit();

  const endTime = Date.now();
  const endMem = getMemoryUsage();
  const fileSize = fs.statSync(filename).size;

  cleanup(filename);

  return {
    library: 'exceljs (streaming)',
    time: endTime - startTime,
    startHeap: startMem.heapUsed,
    endHeap: endMem.heapUsed,
    peakRss: endMem.rss,
    fileSize: Math.round(fileSize / 1024 / 1024 * 100) / 100
  };
}

// ============================================
// Run all benchmarks
// ============================================
async function runBenchmarks() {
  console.log('='.repeat(70));
  console.log('Performance Benchmark: Writing 100,000 rows to Excel');
  console.log('='.repeat(70));
  console.log(`Rows: ${NUM_ROWS.toLocaleString()}`);
  console.log(`Columns: ${NUM_COLS}`);
  console.log(`Node.js: ${process.version}`);
  console.log(`Platform: ${process.platform} ${process.arch}`);
  console.log('='.repeat(70));
  console.log('');

  const results = [];

  // Run benchmarks with delays between to allow memory to settle
  console.log('Running rusc-xlsx (constant memory)...');
  results.push(await benchmarkRuscXlsx());
  await new Promise(r => setTimeout(r, 1000));

  console.log('Running rusc-xlsx (normal)...');
  results.push(await benchmarkRuscXlsxNormal());
  await new Promise(r => setTimeout(r, 1000));

  console.log('Running xlsx (SheetJS)...');
  results.push(await benchmarkXlsx());
  await new Promise(r => setTimeout(r, 1000));

  console.log('Running exceljs...');
  results.push(await benchmarkExcelJS());
  await new Promise(r => setTimeout(r, 1000));

  console.log('Running exceljs (streaming)...');
  results.push(await benchmarkExcelJSStreaming());

  console.log('\n');
  console.log('='.repeat(70));
  console.log('RESULTS');
  console.log('='.repeat(70));
  console.log('');

  // Sort by time
  results.sort((a, b) => a.time - b.time);

  // Print results table
  console.log('| Library                      | Time (ms) | Time (s) | File Size (MB) |');
  console.log('|------------------------------|-----------|----------|----------------|');

  results.forEach((r, i) => {
    const timeS = (r.time / 1000).toFixed(2);
    const rank = i === 0 ? ' (fastest)' : '';
    console.log(`| ${(r.library + rank).padEnd(28)} | ${String(r.time).padStart(9)} | ${timeS.padStart(8)} | ${String(r.fileSize).padStart(14)} |`);
  });

  console.log('');
  console.log('Speed comparison (relative to fastest):');
  const fastest = results[0].time;
  results.forEach(r => {
    const ratio = (r.time / fastest).toFixed(2);
    console.log(`  ${r.library}: ${ratio}x`);
  });

  console.log('');
  console.log('='.repeat(70));
}

runBenchmarks().catch(console.error);
