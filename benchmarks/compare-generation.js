'use strict'

const fs = require('fs')
const os = require('os')
const path = require('path')
const { performance } = require('perf_hooks')
const { spawnSync } = require('child_process')
const AdmZip = require('adm-zip')

const ROOT = path.resolve(__dirname, '..')
const PURE_RUST_MANIFEST = path.join(__dirname, 'pure-rust', 'Cargo.toml')
const PURE_RUST_TARGET = path.join(__dirname, 'pure-rust', 'target')
const PURE_RUST_BINARY = path.join(
  PURE_RUST_TARGET,
  'release',
  `rusc-xlsx-pure-rust-benchmark${process.platform === 'win32' ? '.exe' : ''}`,
)
const NODE_RUNNER = path.join(__dirname, 'node-package.js')
const DEFAULT_ROWS = 200_000
const DEFAULT_COLUMNS = 20
const DEFAULT_RUNS = 3
const DEFAULT_BATCH_ROWS = 1_000

function parsePositiveInteger(value, name, maximum = Number.MAX_SAFE_INTEGER) {
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed) || parsed <= 0 || parsed > maximum) {
    throw new Error(`${name} must be an integer from 1 to ${maximum}`)
  }
  return parsed
}

function parseArgs(argv) {
  const config = {
    rows: DEFAULT_ROWS,
    columns: DEFAULT_COLUMNS,
    runs: DEFAULT_RUNS,
    nodeWriteMode: 'cell',
    batchRows: DEFAULT_BATCH_ROWS,
    keepFiles: false,
    skipBuild: false,
  }

  for (let index = 0; index < argv.length; index++) {
    const flag = argv[index]
    switch (flag) {
      case '--rows':
        config.rows = parsePositiveInteger(argv[++index], 'rows', 1_048_576)
        break
      case '--columns':
        config.columns = parsePositiveInteger(argv[++index], 'columns', 16_384)
        break
      case '--runs':
        config.runs = parsePositiveInteger(argv[++index], 'runs', 100)
        break
      case '--node-write-mode': {
        const writeMode = argv[++index]
        if (!['cell', 'row', 'batch'].includes(writeMode)) {
          throw new Error('node-write-mode must be one of: cell, row, batch')
        }
        config.nodeWriteMode = writeMode
        break
      }
      case '--batch-rows':
        config.batchRows = parsePositiveInteger(argv[++index], 'batch-rows', 1_048_576)
        break
      case '--keep-files':
        config.keepFiles = true
        break
      case '--skip-build':
        config.skipBuild = true
        break
      case '--help':
      case '-h':
        config.help = true
        break
      default:
        throw new Error(`Unknown argument: ${flag}`)
    }
  }

  return config
}

function printHelp() {
  process.stdout.write([
    'Compare pure rust_xlsxwriter with @nazarkalytiuk/rusc-xlsx.',
    '',
    'Usage: node benchmarks/compare-generation.js [options]',
    '',
    'Options:',
    '  --rows N         Rows per workbook (default: 200000)',
    '  --columns N      Columns per workbook (default: 20)',
    '  --runs N         Measured runs per implementation (default: 3)',
    '  --node-write-mode MODE  Node API: cell, row, or batch (default: cell)',
    '  --batch-rows N          Rows per writeRows() call (default: 1000)',
    '  --skip-build     Reuse existing release binaries',
    '  --keep-files     Keep generated XLSX files',
    '  --help           Show this help',
    '',
  ].join('\n'))
}

function runInherited(command, args) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    stdio: 'inherit',
    env: process.env,
  })

  if (result.error) {
    throw result.error
  }
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} exited with status ${result.status}`)
  }
}

function buildReleaseBinaries() {
  const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm'
  process.stdout.write('Building the N-API package in release mode...\n')
  runInherited(npm, ['run', 'build'])

  process.stdout.write('Building the pure Rust generator in release mode...\n')
  runInherited('cargo', [
    'build',
    '--release',
    '--manifest-path',
    PURE_RUST_MANIFEST,
    '--target-dir',
    PURE_RUST_TARGET,
  ])
}

function executeRunner(command, args) {
  const processStart = performance.now()
  const result = spawnSync(command, args, {
    cwd: ROOT,
    encoding: 'utf8',
    maxBuffer: 1024 * 1024,
    env: process.env,
  })
  const processWallMs = performance.now() - processStart

  if (result.error) {
    throw result.error
  }
  if (result.status !== 0) {
    throw new Error(
      `${command} failed with status ${result.status}\n${result.stderr || result.stdout}`,
    )
  }

  const lines = result.stdout.trim().split(/\r?\n/).filter(Boolean)
  const resultLine = lines.at(-1)
  let parsed
  try {
    parsed = JSON.parse(resultLine)
  } catch {
    throw new Error(`Runner returned invalid JSON:\n${result.stdout}`)
  }

  return { ...parsed, processWallMs }
}

function runImplementation(implementation, config, output) {
  const commonArgs = [
    '--rows',
    String(config.rows),
    '--columns',
    String(config.columns),
    '--output',
    output,
  ]

  if (implementation === 'pure-rust') {
    return executeRunner(PURE_RUST_BINARY, commonArgs)
  }
  return executeRunner(process.execPath, [
    NODE_RUNNER,
    ...commonArgs,
    '--write-mode',
    config.nodeWriteMode,
    '--batch-rows',
    String(config.batchRows),
  ])
}

function worksheetFingerprint(filename) {
  const zip = new AdmZip(filename)
  const entry = zip.getEntry('xl/worksheets/sheet1.xml')
  if (!entry) {
    throw new Error(`${filename} does not contain xl/worksheets/sheet1.xml`)
  }
  return {
    crc: entry.header.crc >>> 0,
    size: entry.header.size,
  }
}

function validatePair(pureRust, nodePackage, pureRustFile, nodePackageFile, config) {
  const expectedCells = config.rows * config.columns
  for (const result of [pureRust, nodePackage]) {
    if (
      result.rows !== config.rows ||
      result.columns !== config.columns ||
      result.cells !== expectedCells
    ) {
      throw new Error(`${result.implementation} reported an unexpected workload`)
    }
  }

  const pureRustSheet = worksheetFingerprint(pureRustFile)
  const nodePackageSheet = worksheetFingerprint(nodePackageFile)
  if (
    pureRustSheet.crc !== nodePackageSheet.crc ||
    pureRustSheet.size !== nodePackageSheet.size
  ) {
    throw new Error(
      `Generated worksheet data differs: pure Rust ${JSON.stringify(pureRustSheet)}, ` +
        `Node package ${JSON.stringify(nodePackageSheet)}`,
    )
  }

  return pureRustSheet
}

function median(values) {
  if (values.length === 0) {
    return null
  }
  const sorted = [...values].sort((left, right) => left - right)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0
    ? (sorted[middle - 1] + sorted[middle]) / 2
    : sorted[middle]
}

function formatMilliseconds(value) {
  return `${value.toFixed(1)} ms`
}

function formatMiB(bytes) {
  return bytes == null ? 'n/a' : `${(bytes / 1024 / 1024).toFixed(1)} MiB`
}

function formatRate(cells, milliseconds) {
  return Math.round(cells / (milliseconds / 1_000)).toLocaleString('en-US')
}

function printResults(results, config) {
  const runRows = []
  for (let index = 0; index < config.runs; index++) {
    for (const implementation of ['pure-rust', 'node-package']) {
      const result = results[implementation][index]
      runRows.push({
        Run: index + 1,
        Implementation: result.implementation,
        Total: formatMilliseconds(result.totalMs),
        Write: formatMilliseconds(result.writeMs),
        Save: formatMilliseconds(result.saveMs),
        'Cells/s': formatRate(result.cells, result.totalMs),
        'Max RSS': formatMiB(result.maxRssBytes),
        XLSX: formatMiB(result.outputBytes),
      })
    }
  }

  process.stdout.write('\nMeasured runs:\n')
  console.table(runRows)

  const summaries = {}
  const summaryRows = []
  for (const implementation of ['pure-rust', 'node-package']) {
    const implementationResults = results[implementation]
    const summary = {
      totalMs: median(implementationResults.map((result) => result.totalMs)),
      writeMs: median(implementationResults.map((result) => result.writeMs)),
      saveMs: median(implementationResults.map((result) => result.saveMs)),
      maxRssBytes: median(
        implementationResults
          .map((result) => result.maxRssBytes)
          .filter((value) => value != null),
      ),
      outputBytes: median(implementationResults.map((result) => result.outputBytes)),
    }
    summaries[implementation] = summary
    summaryRows.push({
      Implementation: implementationResults[0].implementation,
      'Median total': formatMilliseconds(summary.totalMs),
      'Median write': formatMilliseconds(summary.writeMs),
      'Median save': formatMilliseconds(summary.saveMs),
      'Median cells/s': formatRate(config.rows * config.columns, summary.totalMs),
      'Median max RSS': formatMiB(summary.maxRssBytes),
      'Median XLSX': formatMiB(summary.outputBytes),
    })
  }

  process.stdout.write('\nMedian summary:\n')
  console.table(summaryRows)

  const ratio = summaries['node-package'].totalMs / summaries['pure-rust'].totalMs
  const percentDifference = (ratio - 1) * 100
  process.stdout.write(
    `Node package / pure Rust total-time ratio: ${ratio.toFixed(2)}x ` +
      `(${percentDifference >= 0 ? '+' : ''}${percentDifference.toFixed(1)}%).\n`,
  )
}

function main() {
  const config = parseArgs(process.argv.slice(2))
  if (config.help) {
    printHelp()
    return
  }

  if (!config.skipBuild) {
    buildReleaseBinaries()
  }
  if (!fs.existsSync(PURE_RUST_BINARY)) {
    throw new Error(`Pure Rust release binary not found: ${PURE_RUST_BINARY}`)
  }

  const workDirectory = fs.mkdtempSync(
    path.join(os.tmpdir(), 'rusc-xlsx-generation-benchmark-'),
  )
  const results = { 'pure-rust': [], 'node-package': [] }
  let validatedFingerprint

  process.stdout.write('\nGeneration benchmark\n')
  process.stdout.write(`Rows: ${config.rows.toLocaleString('en-US')}\n`)
  process.stdout.write(`Columns: ${config.columns.toLocaleString('en-US')}\n`)
  process.stdout.write(`Cells: ${(config.rows * config.columns).toLocaleString('en-US')}\n`)
  process.stdout.write(`Mode: constant memory\n`)
  const nodeWriteApi = config.nodeWriteMode === 'cell'
    ? 'write()'
    : config.nodeWriteMode === 'row'
      ? 'writeRow()'
      : `writeRows() in batches of ${config.batchRows.toLocaleString('en-US')} rows`
  process.stdout.write(`Node write API: ${nodeWriteApi}\n`)
  process.stdout.write(`Runs: ${config.runs}\n`)
  process.stdout.write(`Node: ${process.version}\n`)
  process.stdout.write(`CPU: ${os.cpus()[0]?.model || 'unknown'}\n`)
  process.stdout.write('Build time, process startup, and output validation are excluded.\n\n')

  try {
    for (let run = 0; run < config.runs; run++) {
      const pureRustFile = path.join(workDirectory, `run-${run + 1}-pure-rust.xlsx`)
      const nodePackageFile = path.join(workDirectory, `run-${run + 1}-node-package.xlsx`)
      const files = {
        'pure-rust': pureRustFile,
        'node-package': nodePackageFile,
      }
      const order = run % 2 === 0
        ? ['pure-rust', 'node-package']
        : ['node-package', 'pure-rust']
      const pair = {}

      for (const implementation of order) {
        process.stdout.write(
          `[run ${run + 1}/${config.runs}] ${implementation}... `,
        )
        pair[implementation] = runImplementation(
          implementation,
          config,
          files[implementation],
        )
        process.stdout.write(`${formatMilliseconds(pair[implementation].totalMs)}\n`)
      }

      validatedFingerprint = validatePair(
        pair['pure-rust'],
        pair['node-package'],
        pureRustFile,
        nodePackageFile,
        config,
      )
      results['pure-rust'].push(pair['pure-rust'])
      results['node-package'].push(pair['node-package'])
      if (!config.keepFiles) {
        fs.rmSync(pureRustFile, { force: true })
        fs.rmSync(nodePackageFile, { force: true })
      }
    }

    process.stdout.write(
      `Worksheet equivalence verified for every run ` +
        `(CRC ${validatedFingerprint.crc}, ${validatedFingerprint.size.toLocaleString('en-US')} XML bytes).\n`,
    )
    printResults(results, config)

    if (config.keepFiles) {
      process.stdout.write(`Generated files kept in ${workDirectory}\n`)
    }
  } finally {
    if (!config.keepFiles) {
      fs.rmSync(workDirectory, { recursive: true, force: true })
    }
  }
}

try {
  main()
} catch (error) {
  process.stderr.write(`Benchmark failed: ${error.stack || error.message}\n`)
  process.exitCode = 1
}
