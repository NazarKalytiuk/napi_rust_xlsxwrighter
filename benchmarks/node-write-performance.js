'use strict';

// The JavaScript side of benchmarks/pure-rust/src/bin/write_performance.rs: the
// workloads of rust_xlsxwriter's benches/write_performance.rs (v0.102.0)
// written through this package, ExcelJS or SheetJS, with the same
// configurations, sample sizes and per-iteration work (create a workbook,
// populate it, save it to a temporary file, remove the file).
//
// Rust drops each workbook at the end of its iteration. A JavaScript workbook
// is freed only when V8 collects it, and this package's Rust data only after
// Node runs the N-API finalizer, which happens once the event loop turns. So
// after every iteration this runner collects garbage and yields once, outside
// the timed region, and for this package checks that the Rust heap is back
// where the configuration started. Needs `node --expose-gc`.
//
// Every (group, size) pair prints one JSON line with the time of each measured
// iteration. --validate-dir DIR saves each workbook once to DIR instead,
// --memory GROUP writes the 1M-cell workbook of GROUP once and prints the peak
// RSS, and --describe prints the library version and supported modes.

const fs = require('fs');
const os = require('os');
const path = require('path');
const { performance } = require('perf_hooks');
const {
  COLS,
  CONFIGS,
  GROUPS,
  MEMORY_ROWS,
  MODES,
  sampleSize,
  valueAt,
  workbookName,
} = require('./write-performance-workloads.js');

const DEFAULT_WARM_UP_MS = 1_000;
const LIBRARIES = ['rusc-xlsx', 'exceljs', 'sheetjs'];
const WRITE_MODES = ['cell', 'row'];

// Rust allocations that may legitimately outlive a configuration.
const RETAINED_RUST_BYTES_TOLERANCE = 1024 * 1024;

function rowValues(row, cols, valueOf) {
  // oxlint-disable-next-line unicorn/no-new-array -- preallocating the row is ~5x faster than Array.from and this runs inside the timed loop
  const values = new Array(cols);
  for (let col = 0; col < cols; col++) {
    values[col] = valueOf(row, col);
  }
  return values;
}

// One write_*() per upstream group through the per-cell typed API that mirrors
// write_string()/write_number().
const CELL_WRITERS = {
  numbers(worksheet, rows, cols) {
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        worksheet.writeNumber(row, col, 12345.0);
      }
    }
  },
  mixed(worksheet, rows, cols) {
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        if (col % 2 === 1) {
          worksheet.writeString(row, col, 'Foo');
        } else {
          worksheet.writeNumber(row, col, 12345.0);
        }
      }
    }
  },
  string_static(worksheet, rows, cols) {
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        worksheet.writeString(row, col, 'Foo');
      }
    }
  },
  string_dynamic(worksheet, rows, cols) {
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        worksheet.writeString(row, col, `Cell (${row}, ${col})`);
      }
    }
  },
};

function ruscXlsxLibrary(writeMode) {
  const { Workbook, rustMemoryStats } = require('../wrapper.js');
  const addWorksheet = {
    default: (workbook) => workbook.addWorksheet(),
    constant: (workbook) => workbook.addWorksheetWithConstantMemory(),
    low: (workbook) => workbook.addWorksheetWithLowMemory(),
  };

  const writes =
    writeMode === 'cell' ? '`writeNumber()`/`writeString()` per cell' : '`writeRow()` per row';
  return {
    version: require('../package.json').version,
    modes: Object.keys(addWorksheet),
    apis: {
      default: `\`addWorksheet()\`, ${writes}`,
      constant: `\`addWorksheetWithConstantMemory()\`, ${writes}`,
      low: `\`addWorksheetWithLowMemory()\`, ${writes}`,
    },
    retainedNativeBytes: () => rustMemoryStats()[0],
    write(rows, cols, mode, group, file) {
      const workbook = new Workbook();
      const worksheet = addWorksheet[mode](workbook);
      if (writeMode === 'cell') {
        CELL_WRITERS[group](worksheet, rows, cols);
      } else {
        const valueOf = valueAt(group);
        for (let row = 0; row < rows; row++) {
          worksheet.writeRow(row, 0, rowValues(row, cols, valueOf));
        }
      }
      workbook.save(file);
    },
  };
}

// default: the in-memory Workbook; constant: the streaming WorkbookWriter with
// inline strings; low: the streaming WorkbookWriter with a shared string table.
// DEFLATE level 6 as rust_xlsxwriter uses.
function excelJsLibrary() {
  const ExcelJS = require('exceljs');

  return {
    version: require('exceljs/package.json').version,
    modes: MODES,
    apis: {
      default: '`Workbook`, `addRow()` per row, `xlsx.writeFile()`',
      constant: '`stream.xlsx.WorkbookWriter` (inline strings), `addRow().commit()` per row',
      low: '`stream.xlsx.WorkbookWriter` with `useSharedStrings`, `addRow().commit()` per row',
    },
    async write(rows, cols, mode, group, file) {
      const valueOf = valueAt(group);
      if (mode === 'default') {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sheet1');
        for (let row = 0; row < rows; row++) {
          worksheet.addRow(rowValues(row, cols, valueOf));
        }
        await workbook.xlsx.writeFile(file, {
          zip: { compression: 'DEFLATE', compressionOptions: { level: 6 } },
        });
        return;
      }

      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: file,
        useStyles: false,
        useSharedStrings: mode === 'low',
        zip: { zlib: { level: 6 } },
      });
      const worksheet = workbook.addWorksheet('Sheet1');
      for (let row = 0; row < rows; row++) {
        worksheet.addRow(rowValues(row, cols, valueOf)).commit();
      }
      worksheet.commit();
      await workbook.commit();
    },
  };
}

// SheetJS writes only in memory. bookSST gives the shared string table that
// rust_xlsxwriter writes by default; without it strings become t="str" cells,
// the type of a formula's string result.
function sheetJsLibrary() {
  const XLSX = require('xlsx');

  return {
    version: XLSX.version,
    modes: ['default'],
    apis: {
      default:
        '`aoa_to_sheet()` (dense) from row arrays, `writeFile()` with `compression` and `bookSST`',
    },
    write(rows, cols, _mode, group, file) {
      const valueOf = valueAt(group);
      // oxlint-disable-next-line unicorn/no-new-array -- preallocating the rows is faster than Array.from and this runs inside the timed loop
      const data = new Array(rows);
      for (let row = 0; row < rows; row++) {
        data[row] = rowValues(row, cols, valueOf);
      }
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(data, { dense: true }),
        'Sheet1',
      );
      XLSX.writeFile(workbook, file, { compression: true, bookSST: true });
    },
  };
}

function loadLibrary(config) {
  switch (config.library) {
    case 'rusc-xlsx':
      return ruscXlsxLibrary(config.writeMode);
    case 'exceljs':
      return excelJsLibrary();
    case 'sheetjs':
      return sheetJsLibrary();
    default:
      throw new Error(`Unknown library: ${config.library}`);
  }
}

function parseArgs(argv) {
  const config = {
    library: 'rusc-xlsx',
    mode: 'default',
    writeMode: 'cell',
    warmUpMs: DEFAULT_WARM_UP_MS,
    validateDir: null,
    memoryGroup: null,
    describe: false,
  };

  for (let index = 0; index < argv.length; index++) {
    const flag = argv[index];
    if (flag === '--describe') {
      config.describe = true;
      continue;
    }

    const value = argv[++index];
    if (value === undefined) {
      throw new Error(`Missing value for ${flag}`);
    }
    switch (flag) {
      case '--library':
        if (!LIBRARIES.includes(value)) {
          throw new Error(`library must be one of: ${LIBRARIES.join(', ')} (got ${value})`);
        }
        config.library = value;
        break;
      case '--mode':
        if (!MODES.includes(value)) {
          throw new Error(`mode must be one of: ${MODES.join(', ')} (got ${value})`);
        }
        config.mode = value;
        break;
      case '--write-mode':
        if (!WRITE_MODES.includes(value)) {
          throw new Error(`write-mode must be one of: ${WRITE_MODES.join(', ')} (got ${value})`);
        }
        config.writeMode = value;
        break;
      case '--warm-up-ms': {
        const parsed = Number(value);
        if (!Number.isSafeInteger(parsed) || parsed < 0) {
          throw new Error('warm-up-ms must be a non-negative integer');
        }
        config.warmUpMs = parsed;
        break;
      }
      case '--validate-dir':
        config.validateDir = path.resolve(value);
        break;
      case '--memory':
        valueAt(value);
        config.memoryGroup = value;
        break;
      default:
        throw new Error(`Unknown argument: ${flag}`);
    }
  }

  return config;
}

function maxRssBytes() {
  return process.resourceUsage().maxRSS * 1_024;
}

function tempWorkbookPath() {
  return path.join(os.tmpdir(), `bench_write_perf_${process.pid}.xlsx`);
}

// Create a workbook, populate it and save it to a temporary file.
async function writeWorkbook(library, rows, mode, group) {
  const tempFile = tempWorkbookPath();
  await library.write(rows, COLS, mode, group, tempFile);
  fs.rmSync(tempFile, { force: true });
}

// Frees the workbooks of finished iterations, as Rust's drop does.
async function releaseWorkbooks() {
  globalThis.gc();
  await new Promise((resolve) => setImmediate(resolve));
}

async function measureMemory(library, config) {
  const baselineRssBytes = maxRssBytes();
  const tempFile = tempWorkbookPath();
  await library.write(MEMORY_ROWS, COLS, config.mode, config.memoryGroup, tempFile);
  const peakRssBytes = maxRssBytes();
  const outputBytes = fs.statSync(tempFile).size;
  fs.rmSync(tempFile, { force: true });

  return {
    group: config.memoryGroup,
    cells: MEMORY_ROWS * COLS,
    baselineRssBytes,
    peakRssBytes,
    outputBytes,
  };
}

async function measureGroup(library, config, group) {
  const results = [];
  for (const [rows, sizeName] of CONFIGS) {
    const cellCount = rows * COLS;

    await releaseWorkbooks();
    const retainedBefore = library.retainedNativeBytes?.();

    // At least one warm-up iteration, then as many as fit the warm-up time.
    const warmUpStart = performance.now();
    do {
      await writeWorkbook(library, rows, config.mode, group);
      await releaseWorkbooks();
    } while (performance.now() - warmUpStart < config.warmUpMs);

    const samplesMs = [];
    for (let sample = 0; sample < sampleSize(cellCount); sample++) {
      const start = performance.now();
      await writeWorkbook(library, rows, config.mode, group);
      samplesMs.push(performance.now() - start);
      await releaseWorkbooks();
    }

    if (retainedBefore !== undefined) {
      const retainedBytes = library.retainedNativeBytes() - retainedBefore;
      if (retainedBytes > RETAINED_RUST_BYTES_TOLERANCE) {
        throw new Error(
          `${group}/${sizeName}: ${retainedBytes} bytes of Rust memory were not freed between iterations`,
        );
      }
    }

    results.push({ group, size: sizeName, rows, columns: COLS, cells: cellCount, samplesMs });
  }
  return results;
}

async function run(config) {
  if (typeof globalThis.gc !== 'function') {
    throw new Error('Run with `node --expose-gc` so workbooks can be freed between iterations');
  }

  const library = loadLibrary(config);
  if (config.describe) {
    const { version, modes, apis } = library;
    process.stdout.write(`${JSON.stringify({ version, modes, apis })}\n`);
    return;
  }
  if (!library.modes.includes(config.mode)) {
    throw new Error(`${config.library} has no ${config.mode} mode`);
  }

  const identity = { implementation: config.library, mode: config.mode };

  if (config.memoryGroup) {
    const result = await measureMemory(library, config);
    process.stdout.write(`${JSON.stringify({ ...identity, ...result })}\n`);
    return;
  }

  for (const [group] of GROUPS) {
    if (config.validateDir) {
      for (const [rows] of CONFIGS) {
        const file = path.join(config.validateDir, workbookName(group, rows * COLS));
        await library.write(rows, COLS, config.mode, group, file);
        await releaseWorkbooks();
      }
      continue;
    }

    for (const result of await measureGroup(library, config, group)) {
      process.stdout.write(`${JSON.stringify({ ...identity, ...result })}\n`);
    }
  }
}

if (process.argv.includes('--help') || process.argv.includes('-h')) {
  process.stdout.write(
    'Usage: node --expose-gc node-write-performance.js [--library rusc-xlsx|exceljs|sheetjs] ' +
      '[--mode default|constant|low] [--write-mode cell|row] [--warm-up-ms N] ' +
      '[--validate-dir DIR | --memory GROUP | --describe]\n',
  );
} else {
  run(parseArgs(process.argv.slice(2))).catch((error) => {
    process.stderr.write(`Node write_performance failed: ${error.stack || error.message}\n`);
    process.exitCode = 1;
  });
}
