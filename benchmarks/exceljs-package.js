'use strict';

const fs = require('fs');
const path = require('path');
const { performance } = require('perf_hooks');
const ExcelJS = require('exceljs');
const packageJson = require('exceljs/package.json');

const RUNTIME = process.versions.bun ? 'bun' : 'node';
const RUNTIME_VERSION = process.versions.bun || process.version;
const DEFAULT_ROWS = 200_000;
const DEFAULT_COLUMNS = 20;
const MAX_EXCEL_ROWS = 1_048_576;
const MAX_EXCEL_COLUMNS = 16_384;

function parsePositiveInteger(value, name, maximum) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed <= 0 || parsed > maximum) {
    throw new Error(`${name} must be an integer from 1 to ${maximum}`);
  }
  return parsed;
}

function parseArgs(argv) {
  const config = {
    rows: DEFAULT_ROWS,
    columns: DEFAULT_COLUMNS,
    output: path.resolve('exceljs-generation.xlsx'),
  };

  for (let index = 0; index < argv.length; index += 2) {
    const flag = argv[index];
    const value = argv[index + 1];
    if (value === undefined) {
      throw new Error(`Missing value for ${flag}`);
    }

    switch (flag) {
      case '--rows':
        config.rows = parsePositiveInteger(value, 'rows', MAX_EXCEL_ROWS);
        break;
      case '--columns':
        config.columns = parsePositiveInteger(value, 'columns', MAX_EXCEL_COLUMNS);
        break;
      case '--output':
        config.output = path.resolve(value);
        break;
      default:
        throw new Error(`Unknown argument: ${flag}`);
    }
  }

  return config;
}

function makeRow(row, columns) {
  // oxlint-disable-next-line unicorn/no-new-array -- preallocating the row is ~5x faster than Array.from and this runs inside the timed loop
  const values = new Array(columns);
  for (let column = 0; column < columns; column++) {
    values[column] = column % 2 === 0 ? `Row ${row} Col ${column}` : row * column * 0.123;
  }
  return values;
}

function maxRssBytes() {
  const maxRss = process.resourceUsage().maxRSS;
  // Node reports KiB; Bun's Node compatibility API reports bytes.
  return RUNTIME === 'bun' ? maxRss : maxRss * 1_024;
}

async function run(config) {
  const outputDirectory = path.dirname(config.output);
  fs.mkdirSync(outputDirectory, { recursive: true });

  const totalStart = performance.now();
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: config.output,
    useStyles: false,
    useSharedStrings: false,
    zip: { zlib: { level: 6 } },
  });
  const worksheet = workbook.addWorksheet('Data');

  const writeStart = performance.now();
  for (let row = 0; row < config.rows; row++) {
    worksheet.addRow(makeRow(row, config.columns)).commit();
  }
  worksheet.commit();
  const writeMs = performance.now() - writeStart;

  const saveStart = performance.now();
  await workbook.commit();
  const saveMs = performance.now() - saveStart;
  const totalMs = performance.now() - totalStart;

  const cells = config.rows * config.columns;
  const result = {
    implementation: `${RUNTIME}-exceljs/stream`,
    library: 'exceljs',
    libraryVersion: packageJson.version,
    runtime: RUNTIME,
    runtimeVersion: RUNTIME_VERSION,
    mode: 'streaming',
    writeMode: 'row',
    batchRows: null,
    rows: config.rows,
    columns: config.columns,
    cells,
    writeMs,
    saveMs,
    totalMs,
    maxRssBytes: maxRssBytes(),
    outputBytes: fs.statSync(config.output).size,
  };

  process.stdout.write(`${JSON.stringify(result)}\n`);
}

async function main() {
  if (process.argv.includes('--help') || process.argv.includes('-h')) {
    process.stdout.write(
      `Usage: ${RUNTIME} exceljs-package.js [--rows N] [--columns N] [--output PATH]\n`,
    );
    return;
  }
  await run(parseArgs(process.argv.slice(2)));
}

main().catch((error) => {
  process.stderr.write(`${RUNTIME} ExcelJS benchmark failed: ${error.stack || error.message}\n`);
  process.exitCode = 1;
});
