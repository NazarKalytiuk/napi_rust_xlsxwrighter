'use strict';

// Writes the same workbooks through direct rust_xlsxwriter calls, this package,
// ExcelJS and SheetJS, with the same configurations, sample sizes and timing,
// then measures the peak memory and file size of the largest workbook of each
// group in a fresh process. The workloads are rust_xlsxwriter's own
// write_performance benchmark (benches/write_performance.rs, v0.102.0) and a
// 30-column business-report export (benchmarks/export-workload.js).

const fs = require('fs');
const os = require('os');
const path = require('path');
const v8 = require('v8');
const { spawnSync } = require('child_process');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');
const {
  ROOT,
  buildReleaseBinaries,
  cargoMetadata,
  pureRustBinary,
  rustXlsxwriterBuild,
} = require('./release-builds.js');
const {
  DEFAULT_EXPORT_ROWS,
  MODES,
  WORKLOADS,
  valueAt,
  workbookName,
  workload,
} = require('./write-performance-workloads.js');

const NODE_RUNNER = path.join(__dirname, 'node-write-performance.js');
const IMPLEMENTATIONS = ['pure-rust', 'rusc-xlsx', 'exceljs', 'sheetjs'];
const LIBRARY_LABELS = {
  'pure-rust': 'rust_xlsxwriter',
  'rusc-xlsx': 'rusc-xlsx',
  exceljs: 'ExcelJS',
  sheetjs: 'SheetJS',
};
// This package's write APIs: per-cell typed writes and writeRow().
const RUSC_XLSX_APIS = ['cell', 'row'];
const MODE_TITLES = {
  default: 'Default (in-memory) worksheet',
  constant: 'Constant memory',
  low: 'Low memory',
};
const WORKLOAD_TITLES = {
  'write-performance': "rust_xlsxwriter's write_performance benchmark",
  export: 'business-report export',
};
const DEFAULT_WARM_UP_MS = 1_000;
// Workbook parts that record when the file was created.
const TIMESTAMPED_PARTS = new Set(['docProps/core.xml']);
// V8's fatal error when a runner exceeds the JavaScript heap limit.
const HEAP_OUT_OF_MEMORY = 'JavaScript heap out of memory';
const OUT_OF_MEMORY = 'out of memory';
// Excel's serial date of the Unix epoch, and the accepted rounding difference
// between libraries' conversions (well under a millisecond).
const UNIX_EPOCH_SERIAL = 25_569;
const DATE_SERIAL_TOLERANCE = 1e-9;

// The Rust API in each mode; the JavaScript runners describe their own.
function pureRustApis(selected) {
  const writes = selected.hasDates
    ? 'typed `write_*()` per cell (`write_datetime_with_format()` for dates)'
    : '`write_number()`/`write_string()` per cell';
  return {
    default: `\`add_worksheet()\`, ${writes}`,
    constant: `\`add_worksheet_with_constant_memory()\`, ${writes}`,
    low: `\`add_worksheet_with_low_memory()\`, ${writes}`,
  };
}

class OutOfMemoryError extends Error {}

// A participant is an implementation, and for this package one of its write
// APIs: `rusc-xlsx/cell`, `rusc-xlsx/row`.
function participantKeys(config) {
  return config.implementations.flatMap((implementation) =>
    implementation === 'rusc-xlsx'
      ? config.ruscXlsxApis.map((api) => `rusc-xlsx/${api}`)
      : [implementation],
  );
}

function parseParticipant(key) {
  const [library, writeMode = null] = key.split('/');
  return { library, writeMode };
}

function label(key) {
  const { library, writeMode } = parseParticipant(key);
  return writeMode ? `${LIBRARY_LABELS[library]} (${writeMode})` : LIBRARY_LABELS[library];
}

function parseList(value, allowed, name) {
  const items = (value ?? '').split(',').filter(Boolean);
  const unknown = items.filter((item) => !allowed.includes(item));
  if (items.length === 0 || unknown.length > 0) {
    throw new Error(`${name} must be a comma-separated list of: ${allowed.join(', ')}`);
  }
  return allowed.filter((item) => items.includes(item));
}

function parsePositiveInteger(value, name) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed <= 0) {
    throw new Error(`${name} must be a positive integer`);
  }
  return parsed;
}

function parseArgs(argv) {
  const config = {
    workload: 'write-performance',
    rows: null,
    implementations: IMPLEMENTATIONS,
    modes: MODES,
    ruscXlsxApis: RUSC_XLSX_APIS,
    warmUpMs: DEFAULT_WARM_UP_MS,
    nodeHeapMib: null,
    memory: true,
    markdown: false,
    skipBuild: false,
  };

  for (let index = 0; index < argv.length; index++) {
    const flag = argv[index];
    switch (flag) {
      case '--workload': {
        const name = argv[++index];
        if (!WORKLOADS.includes(name)) {
          throw new Error(`workload must be one of: ${WORKLOADS.join(', ')}`);
        }
        config.workload = name;
        break;
      }
      case '--rows':
        config.rows = parsePositiveInteger(argv[++index], 'rows');
        break;
      case '--implementations':
        config.implementations = parseList(argv[++index], IMPLEMENTATIONS, 'implementations');
        if (!config.implementations.includes('pure-rust')) {
          throw new Error('implementations must include pure-rust, the baseline of every ratio');
        }
        break;
      case '--modes':
        config.modes = parseList(argv[++index], MODES, 'modes');
        break;
      case '--rusc-xlsx-apis':
        config.ruscXlsxApis = parseList(argv[++index], RUSC_XLSX_APIS, 'rusc-xlsx-apis');
        break;
      case '--warm-up-ms': {
        const warmUpMs = Number(argv[++index]);
        if (!Number.isSafeInteger(warmUpMs) || warmUpMs < 0) {
          throw new Error('warm-up-ms must be a non-negative integer');
        }
        config.warmUpMs = warmUpMs;
        break;
      }
      case '--node-heap-mib':
        config.nodeHeapMib = parsePositiveInteger(argv[++index], 'node-heap-mib');
        break;
      case '--skip-memory':
        config.memory = false;
        break;
      case '--markdown':
        config.markdown = true;
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

  if (config.rows !== null && config.workload !== 'export') {
    throw new Error('--rows applies to the export workload only');
  }
  return config;
}

function printHelp() {
  process.stdout.write(
    [
      'Write the same workbooks through rust_xlsxwriter, this package, ExcelJS and SheetJS.',
      '',
      'Usage: node benchmarks/compare-write-performance.js [options]',
      '',
      'Options:',
      `  --workload NAME          ${WORKLOADS.join(' or ')} (default: write-performance)`,
      `  --rows N                 Rows of the export workload (default: ${DEFAULT_EXPORT_ROWS})`,
      `  --implementations LIST   ${IMPLEMENTATIONS.join(',')} (default: all; pure-rust is required)`,
      `  --modes LIST             ${MODES.join(',')} (default: all)`,
      `  --rusc-xlsx-apis LIST    ${RUSC_XLSX_APIS.join(',')}: per-cell writes, writeRow() (default: both)`,
      `  --warm-up-ms N           Warm-up time per benchmark (default: ${DEFAULT_WARM_UP_MS})`,
      "  --node-heap-mib N        Node.js heap limit of the runners (default: Node's own)",
      '  --skip-memory            Do not measure peak memory',
      '  --markdown               Print GitHub-flavored Markdown tables',
      '  --skip-build             Reuse existing release builds',
      '  --help                   Show this help',
      '',
    ].join('\n'),
  );
}

function workloadArgs(config) {
  return [
    '--workload',
    config.workload,
    ...(config.workload === 'export' ? ['--rows', String(context.selected.configs[0][0])] : []),
  ];
}

function runnerCommand(key, args) {
  const { library, writeMode } = parseParticipant(key);
  const fullArgs = [...workloadArgs(context.config), ...args];
  if (library === 'pure-rust') {
    return [context.pureRustBinary, fullArgs];
  }
  const heapArgs = context.config.nodeHeapMib
    ? [`--max-old-space-size=${context.config.nodeHeapMib}`]
    : [];
  const writeModeArgs = writeMode ? ['--write-mode', writeMode] : [];
  return [
    process.execPath,
    [...heapArgs, '--expose-gc', NODE_RUNNER, '--library', library, ...writeModeArgs, ...fullArgs],
  ];
}

function runRunner(key, args) {
  const [command, commandArgs] = runnerCommand(key, args);
  const result = spawnSync(command, commandArgs, {
    cwd: ROOT,
    encoding: 'utf8',
    maxBuffer: 16 * 1024 * 1024,
    // SheetJS converts Dates in local time and every other library in UTC.
    env: { ...process.env, TZ: 'UTC' },
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    const output = result.stderr || result.stdout;
    if (output.includes(HEAP_OUT_OF_MEMORY)) {
      throw new OutOfMemoryError(`${label(key)} ${args.join(' ')}: ${HEAP_OUT_OF_MEMORY}`);
    }
    throw new Error(
      `${label(key)} ${args.join(' ')} failed with status ${result.status}\n${output}`,
    );
  }

  return result.stdout
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        throw new Error(`${label(key)} ${args.join(' ')} printed invalid JSON:\n${result.stdout}`);
      }
    });
}

/** Runs a runner, or records that it ran out of memory for this mode. */
function runOrOutOfMemory(key, mode, args) {
  try {
    return runRunner(key, args);
  } catch (error) {
    if (!(error instanceof OutOfMemoryError)) {
      throw error;
    }
    process.stderr.write(`[${mode}] ${error.message}\n`);
    context.outOfMemory[mode].add(key);
    return null;
  }
}

function describeParticipants() {
  const descriptions = {
    'pure-rust': {
      version: context.rustXlsxwriter.version,
      modes: MODES,
      apis: pureRustApis(context.selected),
    },
  };
  for (const key of participantKeys(context.config)) {
    if (key !== 'pure-rust') {
      [descriptions[key]] = runRunner(key, ['--describe']);
    }
  }
  return descriptions;
}

/** CRC and size of every part except the creation timestamp. */
function workbookFingerprint(filename) {
  const parts = new AdmZip(filename)
    .getEntries()
    .filter((entry) => !TIMESTAMPED_PARTS.has(entry.entryName))
    .map((entry) => [entry.entryName, `${entry.header.crc >>> 0}:${entry.header.size}`])
    .toSorted(([left], [right]) => left.localeCompare(right));
  return JSON.stringify(parts);
}

/** Why a cell read back by SheetJS is not `expected`, or null if it is. */
function cellMismatch(cell, expected) {
  if (expected === null) {
    return cell === undefined || cell.t === 'z' ? null : 'expected an empty cell';
  }
  if (expected instanceof Date) {
    const serial = expected.getTime() / 86_400_000 + UNIX_EPOCH_SERIAL;
    if (cell?.t !== 'n' || Math.abs(cell.v - serial) > DATE_SERIAL_TOLERANCE) {
      return `expected the date serial ${serial}`;
    }
    return cell.z !== undefined && XLSX.SSF.is_date(cell.z)
      ? null
      : `expected a date number format, found ${JSON.stringify(cell.z)}`;
  }
  const type = { number: 'n', string: 's', boolean: 'b' }[typeof expected];
  return cell?.t === type && cell.v === expected ? null : `expected ${JSON.stringify(expected)}`;
}

/** Reads a workbook back and checks the type and value of every cell. */
function checkCellValues(filename, group, rows) {
  const { columns } = context.selected;
  const workbook = XLSX.readFile(filename, { dense: true, cellNF: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const expectedRange = `A1:${XLSX.utils.encode_cell({ r: rows - 1, c: columns - 1 })}`;
  if (sheet['!ref'] !== expectedRange) {
    throw new Error(`${filename}: range ${sheet['!ref']}, expected ${expectedRange}`);
  }

  const valueOf = valueAt(context.selected, group);
  const data = sheet['!data'];
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < columns; col++) {
      const cell = data[row]?.[col];
      const mismatch = cellMismatch(cell, valueOf(row, col));
      if (mismatch) {
        const found = cell ? JSON.stringify({ t: cell.t, v: cell.v, z: cell.z }) : 'empty';
        throw new Error(
          `${filename}: ${XLSX.utils.encode_cell({ r: row, c: col })} is ${found}, ${mismatch}`,
        );
      }
    }
  }
}

function validateMode(mode, participants) {
  const { columns, configs, groups } = context.selected;
  const expectedFiles = groups.flatMap(([group]) =>
    configs.map(([rows]) => ({ name: workbookName(group, rows * columns), group, rows })),
  );
  const directories = {};

  for (const key of participants) {
    const directory = fs.mkdtempSync(
      path.join(context.workDirectory, `${mode}-${key.replace('/', '-')}-`),
    );
    if (!runOrOutOfMemory(key, mode, ['--mode', mode, '--validate-dir', directory])) {
      continue;
    }
    const written = fs.readdirSync(directory).toSorted();
    const expected = expectedFiles.map((file) => file.name).toSorted();
    if (written.join() !== expected.join()) {
      throw new Error(`${label(key)} (${mode}) wrote [${written.join(', ')}]`);
    }
    for (const file of expectedFiles) {
      checkCellValues(path.join(directory, file.name), file.group, file.rows);
    }
    directories[key] = directory;
  }

  // Every write API of this package must write exactly what rust_xlsxwriter writes.
  const ruscXlsx = participants.filter(
    (key) => parseParticipant(key).library === 'rusc-xlsx' && directories[key],
  );
  for (const key of ruscXlsx) {
    for (const { name } of expectedFiles) {
      const pureRust = workbookFingerprint(path.join(directories['pure-rust'], name));
      const written = workbookFingerprint(path.join(directories[key], name));
      if (pureRust !== written) {
        throw new Error(
          `${mode}/${name}: ${label(key)} and rust_xlsxwriter workbooks differ\n` +
            `  rust_xlsxwriter: ${pureRust}\n  ${label(key)}: ${written}`,
        );
      }
    }
  }
}

function rotate(items, offset) {
  const shift = offset % items.length;
  return [...items.slice(shift), ...items.slice(0, shift)];
}

function median(values) {
  const sorted = values.toSorted((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

function measureMode(mode, participants, modeIndex) {
  const running = () => participants.filter((key) => !context.outOfMemory[mode].has(key));

  const timings = {};
  for (const key of rotate(running(), modeIndex)) {
    process.stderr.write(`[${mode}] timing ${label(key)}...\n`);
    timings[key] = runOrOutOfMemory(key, mode, [
      '--mode',
      mode,
      '--warm-up-ms',
      String(context.config.warmUpMs),
    ]);
  }

  const memory = {};
  if (context.config.memory) {
    for (const [groupIndex, [group]] of context.selected.groups.entries()) {
      for (const key of rotate(running(), modeIndex + groupIndex)) {
        process.stderr.write(`[${mode}] memory of ${group} with ${label(key)}...\n`);
        const result = runOrOutOfMemory(key, mode, ['--mode', mode, '--memory', group]);
        if (result) {
          memory[key] ??= {};
          memory[key][group] = result[0];
        }
      }
    }
  }

  return { participants, timings, memory };
}

function formatTime(milliseconds) {
  if (milliseconds < 1) {
    return `${(milliseconds * 1_000).toFixed(1)} µs`;
  }
  if (milliseconds < 1_000) {
    return `${milliseconds.toPrecision(4)} ms`;
  }
  return `${(milliseconds / 1_000).toPrecision(4)} s`;
}

function formatMiB(bytes) {
  return `${(bytes / 1024 / 1024).toFixed(1)} MiB`;
}

function sizeName(rows) {
  return context.selected.configs.find(([configRows]) => configRows === rows)[1];
}

/** Median of each configuration of a participant, keyed by group and rows. */
function medianTimes(timings, key) {
  return new Map(
    (timings[key] ?? []).map((result) => [
      `${result.group}/${result.rows}`,
      median(result.samplesMs),
    ]),
  );
}

function configurationKeys() {
  const { configs, groups } = context.selected;
  return groups.flatMap(([group]) => configs.map(([rows]) => ({ group, rows })));
}

function timingRows({ participants, timings }) {
  const medians = Object.fromEntries(participants.map((key) => [key, medianTimes(timings, key)]));
  return configurationKeys().map(({ group, rows }) => {
    const id = `${group}/${rows}`;
    const rustMs = medians['pure-rust'].get(id);
    const row = { Group: group, Size: sizeName(rows) };
    for (const key of participants) {
      const milliseconds = medians[key].get(id);
      if (milliseconds === undefined) {
        row[label(key)] = OUT_OF_MEMORY;
      } else {
        row[label(key)] =
          key === 'pure-rust'
            ? formatTime(milliseconds)
            : `${formatTime(milliseconds)} (${(milliseconds / rustMs).toFixed(2)}x)`;
      }
    }
    return row;
  });
}

/** Extra time per cell each write API of this package spends over rust_xlsxwriter. */
function overheadRows({ participants, timings }) {
  const ruscXlsx = participants.filter(
    (key) => parseParticipant(key).library === 'rusc-xlsx' && timings[key],
  );
  if (ruscXlsx.length === 0) {
    return null;
  }
  const rust = medianTimes(timings, 'pure-rust');
  const medians = Object.fromEntries(ruscXlsx.map((key) => [key, medianTimes(timings, key)]));
  const { columns } = context.selected;
  return configurationKeys().map(({ group, rows }) => {
    const id = `${group}/${rows}`;
    const row = { Group: group, Size: sizeName(rows) };
    for (const key of ruscXlsx) {
      const extraMs = medians[key].get(id) - rust.get(id);
      row[label(key)] = `${((extraMs * 1e6) / (rows * columns)).toFixed(0)} ns`;
    }
    return row;
  });
}

function memoryRows({ participants, memory }, valueOf) {
  return context.selected.groups.map(([group]) => {
    const row = { Group: group };
    for (const key of participants) {
      const measurement = memory[key]?.[group];
      row[label(key)] = measurement ? valueOf(measurement) : OUT_OF_MEMORY;
    }
    return row;
  });
}

function idleRss({ participants, memory }) {
  return participants
    .filter((key) => memory[key])
    .map((key) => {
      const baselines = Object.values(memory[key]).map(
        (measurement) => measurement.baselineRssBytes,
      );
      return `${label(key)} ${formatMiB(median(baselines))}`;
    })
    .join(', ');
}

function printTable(rows) {
  if (!context.config.markdown) {
    console.table(rows);
    return;
  }
  const columns = Object.keys(rows[0]);
  const textColumns = new Set(['Group', 'Size']);
  process.stdout.write(`| ${columns.join(' | ')} |\n`);
  process.stdout.write(
    `|${columns.map((column) => (textColumns.has(column) ? '---' : '---:')).join('|')}|\n`,
  );
  for (const row of rows) {
    process.stdout.write(`| ${columns.map((column) => row[column]).join(' | ')} |\n`);
  }
}

function printHeading(text) {
  process.stdout.write(context.config.markdown ? `\n#### ${text}\n\n` : `\n${text}:\n`);
}

function commandVersion(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8' });
  return result.status === 0 ? result.stdout.trim() : 'unknown';
}

function nodeHeapLimit() {
  return context.config.nodeHeapMib
    ? `${context.config.nodeHeapMib} MiB (--max-old-space-size)`
    : `${Math.round(v8.getHeapStatistics().heap_size_limit / 1024 / 1024)} MiB (Node's default)`;
}

function printEnvironment(descriptions) {
  const { config, rustXlsxwriter, selected } = context;
  // One entry per library: this package's write APIs share its version.
  const versions = config.implementations.map((implementation) => {
    if (implementation === 'pure-rust') {
      return `rust_xlsxwriter ${rustXlsxwriter.version} (features: ${rustXlsxwriter.features.join(', ')})`;
    }
    const key = participantKeys(config).find(
      (candidate) => parseParticipant(candidate).library === implementation,
    );
    return `${LIBRARY_LABELS[implementation]} ${descriptions[key].version}`;
  });
  const memoryRowsName = sizeName(selected.memoryRows);
  const lines = [
    `Workload: ${WORKLOAD_TITLES[selected.name]}, ${selected.columns} columns`,
    versions.join(', '),
    `${commandVersion('rustc', ['--version'])}, Node.js ${process.version} (heap limit ${nodeHeapLimit()}), TZ=UTC`,
    `${os.cpus()[0]?.model || 'unknown CPU'}, ${os.type()} ${os.release()} (${os.arch()})`,
    `Time: median of 50/20/10 iterations (rust_xlsxwriter's sample sizes) after ${config.warmUpMs} ms ` +
      'of warm-up; an iteration creates a workbook, writes the cells, saves it to a temporary file ' +
      'and removes it. Ratios are against rust_xlsxwriter.',
    ...(config.memory
      ? [
          `Memory: peak RSS increase while writing the ${memoryRowsName} workbook once, in a fresh ` +
            'process, over the RSS before the workbook was created (after the library loaded).',
        ]
      : []),
    'Every implementation wrote the expected value and type in every cell of every ' +
      'configuration (dates with a date number format); rusc-xlsx wrote byte-identical ' +
      'workbooks to rust_xlsxwriter (all parts but docProps/core.xml).',
  ];

  if (config.markdown) {
    process.stdout.write(`\n${lines.map((line) => `- ${line}`).join('\n')}\n`);
  } else {
    process.stdout.write(`\n${lines.join('\n')}\n`);
  }
}

function printApis(mode, participants, descriptions) {
  const lines = participants.map((key) => `${label(key)}: ${descriptions[key].apis[mode]}`);
  if (context.config.markdown) {
    process.stdout.write(`${lines.map((line) => `- ${line}`).join('\n')}\n\n`);
  } else {
    process.stdout.write(`${lines.map((line) => line.replaceAll('`', '')).join('\n')}\n`);
  }
}

function printMode(mode, result, descriptions) {
  const title = MODE_TITLES[mode];
  const memoryRowsName = sizeName(context.selected.memoryRows);
  process.stdout.write(context.config.markdown ? `\n### ${title}\n\n` : `\n=== ${title} ===\n`);
  printApis(mode, result.participants, descriptions);
  if (context.outOfMemory[mode].size > 0) {
    const failed = [...context.outOfMemory[mode]].map(label).join(', ');
    process.stdout.write(`Out of memory at the heap limit of ${nodeHeapLimit()}: ${failed}\n`);
  }

  printHeading(`${title}: median time per workbook`);
  printTable(timingRows(result));

  const overhead = overheadRows(result);
  if (overhead) {
    printHeading(`${title}: rusc-xlsx extra time per cell over rust_xlsxwriter`);
    printTable(overhead);
  }

  if (context.config.memory) {
    printHeading(`${title}: peak RSS increase at ${memoryRowsName}`);
    printTable(
      memoryRows(result, (measurement) =>
        formatMiB(measurement.peakRssBytes - measurement.baselineRssBytes),
      ),
    );
    process.stdout.write(`\nRSS before the workbook: ${idleRss(result)}\n`);

    printHeading(`${title}: XLSX size at ${memoryRowsName}`);
    printTable(memoryRows(result, (measurement) => formatMiB(measurement.outputBytes)));
  }
}

// The run's configuration, release binaries, workload and out-of-memory results.
let context;

function main() {
  const config = parseArgs(process.argv.slice(2));
  if (config.help) {
    printHelp();
    return;
  }

  if (!config.skipBuild) {
    buildReleaseBinaries();
  }
  const metadata = cargoMetadata();
  context = {
    config,
    selected: workload(config.workload, config.rows ?? DEFAULT_EXPORT_ROWS),
    pureRustBinary: pureRustBinary(metadata, 'write_performance'),
    rustXlsxwriter: rustXlsxwriterBuild(metadata),
    workDirectory: fs.mkdtempSync(path.join(os.tmpdir(), 'rusc-xlsx-write-performance-')),
    outOfMemory: Object.fromEntries(MODES.map((mode) => [mode, new Set()])),
  };
  if (!fs.existsSync(context.pureRustBinary)) {
    throw new Error(`Pure Rust release binary not found: ${context.pureRustBinary}`);
  }

  const results = {};
  let descriptions;
  try {
    descriptions = describeParticipants();
    for (const [modeIndex, mode] of config.modes.entries()) {
      const participants = participantKeys(config).filter((key) =>
        descriptions[key].modes.includes(mode),
      );
      process.stderr.write(`[${mode}] validating ${participants.map(label).join(', ')}...\n`);
      validateMode(mode, participants);
      results[mode] = measureMode(mode, participants, modeIndex);
    }
  } finally {
    fs.rmSync(context.workDirectory, { recursive: true, force: true });
  }

  printEnvironment(descriptions);
  for (const mode of config.modes) {
    printMode(mode, results[mode], descriptions);
  }
}

try {
  main();
} catch (error) {
  process.stderr.write(`Benchmark failed: ${error.stack || error.message}\n`);
  process.exitCode = 1;
}
