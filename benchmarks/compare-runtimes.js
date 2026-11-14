'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const { performance } = require('perf_hooks');
const { spawnSync } = require('child_process');
const AdmZip = require('adm-zip');

const ROOT = path.resolve(__dirname, '..');
const rootPackageJson = require('../package.json');
const exceljsPackageJson = require('exceljs/package.json');
const BUN_BINARY = process.platform === 'win32' ? 'bun.exe' : 'bun';
const RUNTIMES = ['node', 'bun'];
const REVERSED_RUNTIMES = ['bun', 'node'];
const BENCHMARKS = {
  'rusc-xlsx': {
    label: '@nazarkalytiuk/rusc-xlsx',
    version: rootPackageJson.version,
    runner: path.join(__dirname, 'node-package.js'),
    mode: 'constant memory',
    sharedImplementation: 'release N-API binary',
    needsBuild: true,
  },
  exceljs: {
    label: 'ExcelJS',
    version: exceljsPackageJson.version,
    runner: path.join(__dirname, 'exceljs-package.js'),
    mode: 'streaming row writer',
    sharedImplementation: 'installed ExcelJS package',
    needsBuild: false,
  },
};
const DEFAULT_ROWS = 200_000;
const DEFAULT_COLUMNS = 20;
const DEFAULT_RUNS = 3;
const DEFAULT_BATCH_ROWS = 1_000;

function parsePositiveInteger(value, name, maximum = Number.MAX_SAFE_INTEGER) {
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
    runs: DEFAULT_RUNS,
    library: 'rusc-xlsx',
    writeMode: 'row',
    batchRows: DEFAULT_BATCH_ROWS,
    keepFiles: false,
    skipBuild: false,
  };

  for (let index = 0; index < argv.length; index++) {
    const flag = argv[index];
    switch (flag) {
      case '--library': {
        const library = argv[++index];
        if (!Object.hasOwn(BENCHMARKS, library)) {
          throw new Error('library must be one of: rusc-xlsx, exceljs');
        }
        config.library = library;
        break;
      }
      case '--rows':
        config.rows = parsePositiveInteger(argv[++index], 'rows', 1_048_576);
        break;
      case '--columns':
        config.columns = parsePositiveInteger(argv[++index], 'columns', 16_384);
        break;
      case '--runs':
        config.runs = parsePositiveInteger(argv[++index], 'runs', 100);
        break;
      case '--write-mode': {
        const writeMode = argv[++index];
        if (!['cell', 'row', 'batch'].includes(writeMode)) {
          throw new Error('write-mode must be one of: cell, row, batch');
        }
        config.writeMode = writeMode;
        break;
      }
      case '--batch-rows':
        config.batchRows = parsePositiveInteger(argv[++index], 'batch-rows', 1_048_576);
        break;
      case '--keep-files':
        config.keepFiles = true;
        break;
      case '--skip-build':
        config.skipBuild = true;
        break;
      case '--help':
      case '-h':
        config.help = true;
        break;
      default:
        throw new Error(`Unknown argument: ${flag}`);
    }
  }
  if (config.library === 'exceljs' && config.writeMode !== 'row') {
    throw new Error('ExcelJS benchmark only supports the streaming row writer');
  }

  return config;
}

function printHelp() {
  process.stdout.write(
    [
      'Compare an XLSX library on Node.js and Bun.',
      '',
      'Usage:',
      '  npm run benchmark:runtimes -- [options]',
      '  npm run benchmark:runtimes:exceljs -- [options]',
      '',
      'Options:',
      '  --library NAME    Library: rusc-xlsx or exceljs (default: rusc-xlsx)',
      '  --rows N          Rows per workbook (default: 200000)',
      '  --columns N       Columns per workbook (default: 20)',
      '  --runs N          Measured runs per runtime (default: 3)',
      '  --write-mode MODE rusc-xlsx API: cell, row, or batch (default: row)',
      '  --batch-rows N    rusc-xlsx rows per writeRows() call (default: 1000)',
      '  --skip-build      Reuse the existing rusc-xlsx release N-API binary',
      '  --keep-files      Keep generated XLSX files',
      '  --help            Show this help',
      '',
    ].join('\n'),
  );
}

function runInherited(command, args) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    stdio: 'inherit',
    env: process.env,
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} exited with status ${result.status}`);
  }
}

function buildReleaseBinary() {
  const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  process.stdout.write('Building the N-API package in release mode...\n');
  runInherited(npm, ['run', 'build']);
}

function readCommandVersion(command) {
  const result = spawnSync(command, ['--version'], {
    cwd: ROOT,
    encoding: 'utf8',
    env: process.env,
  });

  if (result.error?.code === 'ENOENT') {
    throw new Error('Bun is required for this benchmark but was not found in PATH');
  }
  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(
      `${command} --version failed with status ${result.status}\n${result.stderr || result.stdout}`,
    );
  }

  return result.stdout.trim();
}

function executeRunner(command, args) {
  const processStart = performance.now();
  const result = spawnSync(command, args, {
    cwd: ROOT,
    encoding: 'utf8',
    maxBuffer: 1024 * 1024,
    env: process.env,
  });
  const processWallMs = performance.now() - processStart;

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(
      `${command} failed with status ${result.status}\n${result.stderr || result.stdout}`,
    );
  }

  const lines = result.stdout.trim().split(/\r?\n/).filter(Boolean);
  const resultLine = lines.at(-1);
  try {
    return { ...JSON.parse(resultLine), processWallMs };
  } catch {
    throw new Error(`Runner returned invalid JSON:\n${result.stdout}`);
  }
}

function runtimeCommand(runtime) {
  return runtime === 'node' ? process.execPath : BUN_BINARY;
}

function runRuntime(runtime, config, output) {
  const args = [
    BENCHMARKS[config.library].runner,
    '--rows',
    String(config.rows),
    '--columns',
    String(config.columns),
    '--output',
    output,
  ];
  if (config.library === 'rusc-xlsx') {
    args.push('--write-mode', config.writeMode, '--batch-rows', String(config.batchRows));
  }
  return executeRunner(runtimeCommand(runtime), args);
}

function worksheetFingerprint(filename) {
  const zip = new AdmZip(filename);
  const entry = zip.getEntry('xl/worksheets/sheet1.xml');
  if (!entry) {
    throw new Error(`${filename} does not contain xl/worksheets/sheet1.xml`);
  }
  return {
    crc: entry.header.crc >>> 0,
    size: entry.header.size,
  };
}

function validateResult(result, runtime, config, output) {
  const expectedCells = config.rows * config.columns;
  if (
    result.library !== config.library ||
    result.libraryVersion !== BENCHMARKS[config.library].version ||
    result.runtime !== runtime ||
    result.rows !== config.rows ||
    result.columns !== config.columns ||
    result.cells !== expectedCells
  ) {
    throw new Error(`${runtime} reported an unexpected workload or runtime`);
  }
  for (const metric of ['writeMs', 'saveMs', 'totalMs', 'processWallMs']) {
    if (!Number.isFinite(result[metric]) || result[metric] < 0) {
      throw new Error(`${runtime} reported an invalid ${metric}`);
    }
  }
  if (!Number.isFinite(result.maxRssBytes) || result.maxRssBytes <= 0) {
    throw new Error(`${runtime} reported an invalid maxRssBytes`);
  }
  const outputBytes = fs.statSync(output).size;
  if (result.outputBytes !== outputBytes || outputBytes === 0) {
    throw new Error(`${runtime} reported an invalid output size`);
  }
}

function validatePair(pair, files, config) {
  for (const runtime of RUNTIMES) {
    validateResult(pair[runtime], runtime, config, files[runtime]);
  }

  const nodeSheet = worksheetFingerprint(files.node);
  const bunSheet = worksheetFingerprint(files.bun);
  if (nodeSheet.crc !== bunSheet.crc || nodeSheet.size !== bunSheet.size) {
    throw new Error(
      `Generated worksheet data differs: Node.js ${JSON.stringify(nodeSheet)}, ` +
        `Bun ${JSON.stringify(bunSheet)}`,
    );
  }

  return nodeSheet;
}

function median(values) {
  const sorted = values.toSorted((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

function formatMilliseconds(value) {
  return `${value.toFixed(1)} ms`;
}

function formatMiB(bytes) {
  return `${(bytes / 1024 / 1024).toFixed(1)} MiB`;
}

function formatRate(cells, milliseconds) {
  return Math.round(cells / (milliseconds / 1_000)).toLocaleString('en-US');
}

function formatRatio(value, baseline) {
  return `${(value / baseline).toFixed(2)}x`;
}

function formatComparison(value, baseline) {
  const ratio = value / baseline;
  const percentDifference = (ratio - 1) * 100;
  return (
    `${ratio.toFixed(2)}x ` +
    `(${percentDifference >= 0 ? '+' : ''}${percentDifference.toFixed(1)}%)`
  );
}

function runtimeLabel(runtime, version) {
  return runtime === 'node' ? `Node.js ${version}` : `Bun ${version}`;
}

function printResults(results, config) {
  const runRows = [];
  for (let index = 0; index < config.runs; index++) {
    for (const runtime of RUNTIMES) {
      const result = results[runtime][index];
      runRows.push({
        Run: index + 1,
        Runtime: runtimeLabel(runtime, result.runtimeVersion),
        Generation: formatMilliseconds(result.totalMs),
        'Process wall': formatMilliseconds(result.processWallMs),
        Write: formatMilliseconds(result.writeMs),
        Save: formatMilliseconds(result.saveMs),
        'Cells/s': formatRate(result.cells, result.totalMs),
        'Max RSS': formatMiB(result.maxRssBytes),
        XLSX: formatMiB(result.outputBytes),
      });
    }
  }

  process.stdout.write('\nMeasured runs:\n');
  console.table(runRows);

  const summaries = {};
  for (const runtime of RUNTIMES) {
    const runtimeResults = results[runtime];
    summaries[runtime] = {
      runtimeVersion: runtimeResults[0].runtimeVersion,
      totalMs: median(runtimeResults.map((result) => result.totalMs)),
      processWallMs: median(runtimeResults.map((result) => result.processWallMs)),
      writeMs: median(runtimeResults.map((result) => result.writeMs)),
      saveMs: median(runtimeResults.map((result) => result.saveMs)),
      maxRssBytes: median(runtimeResults.map((result) => result.maxRssBytes)),
      outputBytes: median(runtimeResults.map((result) => result.outputBytes)),
    };
  }

  const nodeSummary = summaries.node;
  const summaryRows = RUNTIMES.map((runtime) => {
    const summary = summaries[runtime];
    return {
      Runtime: runtimeLabel(runtime, summary.runtimeVersion),
      'Median generation': formatMilliseconds(summary.totalMs),
      'Median process wall': formatMilliseconds(summary.processWallMs),
      'Median write': formatMilliseconds(summary.writeMs),
      'Median save': formatMilliseconds(summary.saveMs),
      'Median cells/s': formatRate(config.rows * config.columns, summary.totalMs),
      'Median max RSS': formatMiB(summary.maxRssBytes),
      'Median XLSX': formatMiB(summary.outputBytes),
      'Generation / Node': formatRatio(summary.totalMs, nodeSummary.totalMs),
      'Wall / Node': formatRatio(summary.processWallMs, nodeSummary.processWallMs),
      'RSS / Node': formatRatio(summary.maxRssBytes, nodeSummary.maxRssBytes),
      'XLSX / Node': formatRatio(summary.outputBytes, nodeSummary.outputBytes),
    };
  });

  process.stdout.write('\nMedian summary:\n');
  console.table(summaryRows);
  process.stdout.write(
    `Bun / Node.js median generation ratio: ` +
      `${formatComparison(summaries.bun.totalMs, nodeSummary.totalMs)}.\n`,
  );
  process.stdout.write(
    `Bun / Node.js median process-wall ratio: ` +
      `${formatComparison(summaries.bun.processWallMs, nodeSummary.processWallMs)}.\n`,
  );
  process.stdout.write(
    `Bun / Node.js median max-RSS ratio: ` +
      `${formatComparison(summaries.bun.maxRssBytes, nodeSummary.maxRssBytes)}.\n`,
  );
  process.stdout.write(
    `Bun / Node.js median XLSX-size ratio: ` +
      `${formatComparison(summaries.bun.outputBytes, nodeSummary.outputBytes)}.\n`,
  );
}

function writeApiDescription(config) {
  if (config.library === 'exceljs') {
    return 'stream.xlsx.WorkbookWriter + addRow().commit()';
  }
  if (config.writeMode === 'cell') {
    return 'write()';
  }
  if (config.writeMode === 'row') {
    return 'writeRow()';
  }
  return `writeRows() in batches of ${config.batchRows.toLocaleString('en-US')} rows`;
}

function main() {
  const config = parseArgs(process.argv.slice(2));
  if (config.help) {
    printHelp();
    return;
  }
  const benchmark = BENCHMARKS[config.library];
  if (process.versions.bun) {
    throw new Error('Run the benchmark orchestrator with Node.js, not Bun');
  }

  const versions = {
    node: process.version,
    bun: readCommandVersion(BUN_BINARY),
  };
  if (benchmark.needsBuild && !config.skipBuild) {
    buildReleaseBinary();
  }

  const workDirectory = fs.mkdtempSync(
    path.join(os.tmpdir(), `${config.library}-runtime-benchmark-`),
  );
  const results = { node: [], bun: [] };
  let validatedFingerprint;

  process.stdout.write(`\n${benchmark.label}: Node.js vs Bun XLSX generation benchmark\n`);
  process.stdout.write(`Library: ${benchmark.label} ${benchmark.version}\n`);
  process.stdout.write(`Rows: ${config.rows.toLocaleString('en-US')}\n`);
  process.stdout.write(`Columns: ${config.columns.toLocaleString('en-US')}\n`);
  process.stdout.write(`Cells: ${(config.rows * config.columns).toLocaleString('en-US')}\n`);
  process.stdout.write('Dataset: 50% strings, 50% numbers\n');
  process.stdout.write(`Mode: ${benchmark.mode}\n`);
  process.stdout.write(`Package API: ${writeApiDescription(config)}\n`);
  if (config.library === 'exceljs') {
    process.stdout.write('Compression: ZIP DEFLATE level 6\n');
  }
  process.stdout.write(`Runs: ${config.runs}\n`);
  process.stdout.write(`Node.js: ${versions.node}\n`);
  process.stdout.write(`Bun: ${versions.bun}\n`);
  process.stdout.write(`CPU: ${os.cpus()[0]?.model || 'unknown'}\n`);
  process.stdout.write(
    `Both runtimes use the same ${benchmark.sharedImplementation} and generated workload. ` +
      'Generation excludes process startup; process wall includes it. ' +
      'Setup and output validation are excluded.\n\n',
  );

  try {
    for (let run = 0; run < config.runs; run++) {
      const files = {
        node: path.join(workDirectory, `run-${run + 1}-node.xlsx`),
        bun: path.join(workDirectory, `run-${run + 1}-bun.xlsx`),
      };
      const order = run % 2 === 0 ? RUNTIMES : REVERSED_RUNTIMES;
      const pair = {};

      for (const runtime of order) {
        process.stdout.write(`[run ${run + 1}/${config.runs}] ${runtime}... `);
        pair[runtime] = runRuntime(runtime, config, files[runtime]);
        process.stdout.write(`${formatMilliseconds(pair[runtime].totalMs)}\n`);
      }

      validatedFingerprint = validatePair(pair, files, config);
      for (const runtime of RUNTIMES) {
        results[runtime].push(pair[runtime]);
        if (!config.keepFiles) {
          fs.rmSync(files[runtime], { force: true });
        }
      }
    }

    process.stdout.write(
      `Worksheet equivalence verified for every run ` +
        `(CRC ${validatedFingerprint.crc}, ` +
        `${validatedFingerprint.size.toLocaleString('en-US')} XML bytes).\n`,
    );
    printResults(results, config);

    if (config.keepFiles) {
      process.stdout.write(`Generated files kept in ${workDirectory}\n`);
    }
  } finally {
    if (!config.keepFiles) {
      fs.rmSync(workDirectory, { recursive: true, force: true });
    }
  }
}

try {
  main();
} catch (error) {
  process.stderr.write(`Runtime benchmark failed: ${error.stack || error.message}\n`);
  process.exitCode = 1;
}
