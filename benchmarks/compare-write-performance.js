'use strict';

// Runs rust_xlsxwriter's write_performance benchmark (benches/write_performance.rs,
// v0.102.0) through direct rust_xlsxwriter calls, this package, ExcelJS and
// SheetJS, with the same configurations, sample sizes and timing, then
// measures the peak memory of each 1M-cell workbook in a fresh process.

const fs = require('fs');
const os = require('os');
const path = require('path');
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
  COLS,
  CONFIGS,
  GROUPS,
  MEMORY_ROWS,
  MODES,
  valueAt,
  workbookName,
} = require('./write-performance-workloads.js');

const NODE_RUNNER = path.join(__dirname, 'node-write-performance.js');
const IMPLEMENTATIONS = ['pure-rust', 'rusc-xlsx', 'exceljs', 'sheetjs'];
const LIBRARY_LABELS = {
  'pure-rust': 'rust_xlsxwriter',
  'rusc-xlsx': 'rusc-xlsx',
  exceljs: 'ExcelJS',
  sheetjs: 'SheetJS',
};
// This package's write APIs: per-cell writeNumber()/writeString() and writeRow().
const RUSC_XLSX_APIS = ['cell', 'row'];
const MODE_TITLES = {
  default: 'Default (in-memory) worksheet',
  constant: 'Constant memory',
  low: 'Low memory',
};
// Upstream's Rust API in each mode; the JavaScript runners describe their own.
const PURE_RUST_APIS = {
  default: '`add_worksheet()`, `write_number()`/`write_string()` per cell',
  constant: '`add_worksheet_with_constant_memory()`, `write_number()`/`write_string()` per cell',
  low: '`add_worksheet_with_low_memory()`, `write_number()`/`write_string()` per cell',
};
const DEFAULT_WARM_UP_MS = 1_000;
// Workbook parts that record when the file was created.
const TIMESTAMPED_PARTS = new Set(['docProps/core.xml']);

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

function parseArgs(argv) {
  const config = {
    implementations: IMPLEMENTATIONS,
    modes: MODES,
    ruscXlsxApis: RUSC_XLSX_APIS,
    warmUpMs: DEFAULT_WARM_UP_MS,
    memory: true,
    markdown: false,
    skipBuild: false,
  };

  for (let index = 0; index < argv.length; index++) {
    const flag = argv[index];
    switch (flag) {
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

  return config;
}

function printHelp() {
  process.stdout.write(
    [
      "Run rust_xlsxwriter's write_performance benchmark against rust_xlsxwriter,",
      'this package, ExcelJS and SheetJS.',
      '',
      'Usage: node benchmarks/compare-write-performance.js [options]',
      '',
      'Options:',
      `  --implementations LIST   ${IMPLEMENTATIONS.join(',')} (default: all; pure-rust is required)`,
      `  --modes LIST             ${MODES.join(',')} (default: all)`,
      `  --rusc-xlsx-apis LIST    ${RUSC_XLSX_APIS.join(',')}: per-cell writes, writeRow() (default: both)`,
      `  --warm-up-ms N           Warm-up time per benchmark (default: ${DEFAULT_WARM_UP_MS})`,
      '  --skip-memory            Do not measure peak memory',
      '  --markdown               Print GitHub-flavored Markdown tables',
      '  --skip-build             Reuse existing release builds',
      '  --help                   Show this help',
      '',
    ].join('\n'),
  );
}

function runnerCommand(key, context, args) {
  const { library, writeMode } = parseParticipant(key);
  if (library === 'pure-rust') {
    return [context.pureRustBinary, args];
  }
  const writeModeArgs = writeMode ? ['--write-mode', writeMode] : [];
  return [
    process.execPath,
    ['--expose-gc', NODE_RUNNER, '--library', library, ...writeModeArgs, ...args],
  ];
}

function runRunner(implementation, context, args) {
  const [command, commandArgs] = runnerCommand(implementation, context, args);
  const result = spawnSync(command, commandArgs, {
    cwd: ROOT,
    encoding: 'utf8',
    maxBuffer: 16 * 1024 * 1024,
    env: process.env,
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(
      `${implementation} ${args.join(' ')} failed with status ${result.status}\n` +
        (result.stderr || result.stdout),
    );
  }

  return result.stdout
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        throw new Error(
          `${implementation} ${args.join(' ')} printed invalid JSON:\n${result.stdout}`,
        );
      }
    });
}

function describeImplementations(context) {
  const descriptions = {
    'pure-rust': {
      version: context.rustXlsxwriter.version,
      modes: MODES,
      apis: PURE_RUST_APIS,
    },
  };
  for (const key of participantKeys(context.config)) {
    if (key !== 'pure-rust') {
      [descriptions[key]] = runRunner(key, context, ['--describe']);
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

/** Reads a workbook back and checks the type and value of every cell. */
function checkCellValues(filename, group, rows) {
  const workbook = XLSX.readFile(filename, { dense: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const expectedRange = `A1:${XLSX.utils.encode_cell({ r: rows - 1, c: COLS - 1 })}`;
  if (sheet['!ref'] !== expectedRange) {
    throw new Error(`${filename}: range ${sheet['!ref']}, expected ${expectedRange}`);
  }

  const valueOf = valueAt(group);
  const data = sheet['!data'];
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < COLS; col++) {
      const expected = valueOf(row, col);
      const cell = data[row]?.[col];
      if (cell?.t !== (typeof expected === 'number' ? 'n' : 's') || cell.v !== expected) {
        const found = cell ? JSON.stringify({ t: cell.t, v: cell.v }) : 'empty';
        throw new Error(
          `${filename}: ${XLSX.utils.encode_cell({ r: row, c: col })} is ${found}, ` +
            `expected ${JSON.stringify(expected)}`,
        );
      }
    }
  }
}

function validateMode(context, mode, participants) {
  const expectedFiles = GROUPS.flatMap(([group]) =>
    CONFIGS.map(([rows]) => ({ name: workbookName(group, rows * COLS), group, rows })),
  );
  const directories = {};

  for (const key of participants) {
    const directory = fs.mkdtempSync(
      path.join(context.workDirectory, `${mode}-${key.replace('/', '-')}-`),
    );
    runRunner(key, context, ['--mode', mode, '--validate-dir', directory]);
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
  for (const key of participants.filter((id) => parseParticipant(id).library === 'rusc-xlsx')) {
    for (const { name } of expectedFiles) {
      const pureRust = workbookFingerprint(path.join(directories['pure-rust'], name));
      const ruscXlsx = workbookFingerprint(path.join(directories[key], name));
      if (pureRust !== ruscXlsx) {
        throw new Error(
          `${mode}/${name}: ${label(key)} and rust_xlsxwriter workbooks differ\n` +
            `  rust_xlsxwriter: ${pureRust}\n  ${label(key)}: ${ruscXlsx}`,
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

function measureMode(context, mode, participants, modeIndex) {
  const timings = {};
  for (const key of rotate(participants, modeIndex)) {
    process.stderr.write(`[${mode}] timing ${label(key)}...\n`);
    timings[key] = runRunner(key, context, [
      '--mode',
      mode,
      '--warm-up-ms',
      String(context.config.warmUpMs),
    ]);
  }

  const memory = {};
  if (context.config.memory) {
    for (const [groupIndex, [group]] of GROUPS.entries()) {
      for (const key of rotate(participants, modeIndex + groupIndex)) {
        process.stderr.write(`[${mode}] memory of ${group} with ${label(key)}...\n`);
        const [result] = runRunner(key, context, ['--mode', mode, '--memory', group]);
        memory[key] ??= {};
        memory[key][group] = result;
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

function medianTimes(timings, key) {
  return timings[key].map((result) => ({ ...result, medianMs: median(result.samplesMs) }));
}

function timingRows({ participants, timings }) {
  const medians = Object.fromEntries(participants.map((key) => [key, medianTimes(timings, key)]));
  return medians['pure-rust'].map(({ group, size, medianMs: rustMs }, index) => {
    const row = { Group: group, Size: size };
    for (const key of participants) {
      const milliseconds = medians[key][index].medianMs;
      row[label(key)] =
        key === 'pure-rust'
          ? formatTime(milliseconds)
          : `${formatTime(milliseconds)} (${(milliseconds / rustMs).toFixed(2)}x)`;
    }
    return row;
  });
}

/** Extra time per cell each write API of this package spends over rust_xlsxwriter. */
function overheadRows({ participants, timings }) {
  const ruscXlsx = participants.filter((key) => parseParticipant(key).library === 'rusc-xlsx');
  if (ruscXlsx.length === 0) {
    return null;
  }
  const rust = medianTimes(timings, 'pure-rust');
  const medians = Object.fromEntries(ruscXlsx.map((key) => [key, medianTimes(timings, key)]));
  return rust.map(({ group, size, cells, medianMs: rustMs }, index) => {
    const row = { Group: group, Size: size };
    for (const key of ruscXlsx) {
      row[label(key)] =
        `${(((medians[key][index].medianMs - rustMs) * 1e6) / cells).toFixed(0)} ns`;
    }
    return row;
  });
}

function memoryRows({ participants, memory }, valueOf) {
  return GROUPS.map(([group]) => {
    const row = { Group: group };
    for (const key of participants) {
      row[label(key)] = valueOf(memory[key][group]);
    }
    return row;
  });
}

function idleRss({ participants, memory }) {
  return participants
    .map((key) => {
      const baselines = GROUPS.map(([group]) => memory[key][group].baselineRssBytes);
      return `${label(key)} ${formatMiB(median(baselines))}`;
    })
    .join(', ');
}

function printTable(rows, config) {
  if (!config.markdown) {
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

function printHeading(text, config) {
  process.stdout.write(config.markdown ? `\n#### ${text}\n\n` : `\n${text}:\n`);
}

function commandVersion(command, args) {
  const result = spawnSync(command, args, { encoding: 'utf8' });
  return result.status === 0 ? result.stdout.trim() : 'unknown';
}

function printEnvironment(context, descriptions) {
  const { config, rustXlsxwriter } = context;
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
  const lines = [
    versions.join(', '),
    `${commandVersion('rustc', ['--version'])}, Node.js ${process.version}`,
    `${os.cpus()[0]?.model || 'unknown CPU'}, ${os.type()} ${os.release()} (${os.arch()})`,
    `Time: median of 50/20/10 iterations (upstream's sample sizes) after ${config.warmUpMs} ms of ` +
      'warm-up; an iteration creates a workbook, writes the cells, saves it to a temporary file ' +
      'and removes it. Ratios are against rust_xlsxwriter.',
    ...(config.memory
      ? [
          `Memory: peak RSS increase while writing the ${(MEMORY_ROWS * COLS).toLocaleString('en-US')}-cell ` +
            'workbook once, in a fresh process, over the RSS before the workbook was created ' +
            '(after the library loaded).',
        ]
      : []),
    'Every implementation wrote the expected value and type in every cell of every ' +
      'configuration; rusc-xlsx wrote byte-identical workbooks to rust_xlsxwriter ' +
      '(all parts but docProps/core.xml).',
  ];

  if (config.markdown) {
    process.stdout.write(`\n${lines.map((line) => `- ${line}`).join('\n')}\n`);
  } else {
    process.stdout.write(`\n${lines.join('\n')}\n`);
  }
}

function printApis(mode, participants, descriptions, config) {
  const lines = participants.map((key) => `${label(key)}: ${descriptions[key].apis[mode]}`);
  if (config.markdown) {
    process.stdout.write(`${lines.map((line) => `- ${line}`).join('\n')}\n\n`);
  } else {
    process.stdout.write(`${lines.map((line) => line.replaceAll('`', '')).join('\n')}\n`);
  }
}

function printMode(mode, result, descriptions, config) {
  const title = MODE_TITLES[mode];
  process.stdout.write(config.markdown ? `\n### ${title}\n\n` : `\n=== ${title} ===\n`);
  printApis(mode, result.participants, descriptions, config);

  printHeading(`${title}: median time per workbook`, config);
  printTable(timingRows(result), config);

  const overhead = overheadRows(result);
  if (overhead) {
    printHeading(`${title}: rusc-xlsx extra time per cell over rust_xlsxwriter`, config);
    printTable(overhead, config);
  }

  if (config.memory) {
    printHeading(`${title}: peak RSS increase at 1M cells`, config);
    printTable(
      memoryRows(result, (measurement) =>
        formatMiB(measurement.peakRssBytes - measurement.baselineRssBytes),
      ),
      config,
    );
    process.stdout.write(`\nRSS before the workbook: ${idleRss(result)}\n`);

    printHeading(`${title}: XLSX size at 1M cells`, config);
    printTable(
      memoryRows(result, (measurement) => formatMiB(measurement.outputBytes)),
      config,
    );
  }
}

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
  const context = {
    config,
    pureRustBinary: pureRustBinary(metadata, 'write_performance'),
    rustXlsxwriter: rustXlsxwriterBuild(metadata),
    workDirectory: fs.mkdtempSync(path.join(os.tmpdir(), 'rusc-xlsx-write-performance-')),
  };
  if (!fs.existsSync(context.pureRustBinary)) {
    throw new Error(`Pure Rust release binary not found: ${context.pureRustBinary}`);
  }

  const results = {};
  let descriptions;
  try {
    descriptions = describeImplementations(context);
    for (const [modeIndex, mode] of config.modes.entries()) {
      const participants = participantKeys(config).filter((key) =>
        descriptions[key].modes.includes(mode),
      );
      process.stderr.write(`[${mode}] validating ${participants.map(label).join(', ')}...\n`);
      validateMode(context, mode, participants);
      results[mode] = measureMode(context, mode, participants, modeIndex);
    }
  } finally {
    fs.rmSync(context.workDirectory, { recursive: true, force: true });
  }

  printEnvironment(context, descriptions);
  for (const mode of config.modes) {
    printMode(mode, results[mode], descriptions, config);
  }
}

try {
  main();
} catch (error) {
  process.stderr.write(`Benchmark failed: ${error.stack || error.message}\n`);
  process.exitCode = 1;
}
