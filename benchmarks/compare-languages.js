'use strict'

const fs = require('fs')
const os = require('os')
const path = require('path')
const { performance } = require('perf_hooks')
const { spawnSync } = require('child_process')

const ROOT = path.resolve(__dirname, '..')
const NODE_RUNNER = path.join(__dirname, 'node-package.js')
const GO_DIRECTORY = path.join(__dirname, 'go-excelize')
const GO_BINARY = path.join(
  GO_DIRECTORY,
  'bin',
  `go-excelize-benchmark${process.platform === 'win32' ? '.exe' : ''}`,
)
const CSHARP_DIRECTORY = path.join(__dirname, 'csharp-openxml')
const CSHARP_PROJECT = path.join(CSHARP_DIRECTORY, 'CSharpBenchmark.csproj')
const CSHARP_BINARY = path.join(
  CSHARP_DIRECTORY,
  'bin',
  'Release',
  'net8.0',
  'rusc-xlsx-csharp-benchmark.dll',
)
const PYTHON_DIRECTORY = path.join(__dirname, 'python-xlsxwriter')
const PYTHON_RUNNER = path.join(PYTHON_DIRECTORY, 'benchmark.py')
const PYTHON_REQUIREMENTS = path.join(PYTHON_DIRECTORY, 'requirements.txt')
const PYTHON_VENV = path.join(PYTHON_DIRECTORY, '.venv')
const PYTHON_VENV_BINARY = process.platform === 'win32'
  ? path.join(PYTHON_VENV, 'Scripts', 'python.exe')
  : path.join(PYTHON_VENV, 'bin', 'python')
const WORKBOOK_VALIDATOR = path.join(__dirname, 'validate-workbook.py')
const NPM = process.platform === 'win32' ? 'npm.cmd' : 'npm'
const DEFAULT_ROWS = 200_000
const DEFAULT_COLUMNS = 20
const DEFAULT_RUNS = 3
const DEFAULT_BATCH_ROWS = 1_000
const DEFAULT_IMPLEMENTATIONS = ['node', 'go', 'csharp', 'python']
const IMPLEMENTATION_LABELS = {
  node: 'Node.js / rusc-xlsx',
  go: 'Go / Excelize',
  csharp: 'C# / Open XML SDK',
  python: 'Python / XlsxWriter',
}

function parsePositiveInteger(value, name, maximum = Number.MAX_SAFE_INTEGER) {
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed) || parsed <= 0 || parsed > maximum) {
    throw new Error(`${name} must be an integer from 1 to ${maximum}`)
  }
  return parsed
}

function parseImplementations(value) {
  const implementations = value.split(',').map((item) => item.trim()).filter(Boolean)
  if (implementations.length === 0) {
    throw new Error('implementations must contain at least one value')
  }
  const unknown = implementations.filter(
    (implementation) => !DEFAULT_IMPLEMENTATIONS.includes(implementation),
  )
  if (unknown.length > 0) {
    throw new Error(
      `Unknown implementations: ${unknown.join(', ')}. ` +
        `Expected a comma-separated subset of ${DEFAULT_IMPLEMENTATIONS.join(', ')}`,
    )
  }
  if (new Set(implementations).size !== implementations.length) {
    throw new Error('implementations must not contain duplicates')
  }
  return implementations
}

function parseArgs(argv) {
  const config = {
    rows: DEFAULT_ROWS,
    columns: DEFAULT_COLUMNS,
    runs: DEFAULT_RUNS,
    implementations: [...DEFAULT_IMPLEMENTATIONS],
    nodeWriteMode: 'row',
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
      case '--implementations':
        config.implementations = parseImplementations(argv[++index] || '')
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
    'Compare streaming XLSX generation in Node.js, Go, C#, and Python.',
    '',
    'Usage: node benchmarks/compare-languages.js [options]',
    '',
    'Options:',
    '  --rows N                    Rows per workbook (default: 200000)',
    '  --columns N                 Columns per workbook (default: 20)',
    '  --runs N                    Measured runs per implementation (default: 3)',
    '  --implementations LIST      Comma-separated subset of node,go,csharp,python',
    '  --node-write-mode MODE      Node API: cell, row, or batch (default: row)',
    '  --batch-rows N              Rows per Node writeRows() call (default: 1000)',
    '  --skip-build                Reuse existing release artifacts and Python venv',
    '  --keep-files                Keep measured XLSX files',
    '  --help                      Show this help',
    '',
    'Prerequisites for the default set: Rust, Go 1.25+, .NET SDK 8, and Python 3.10+.',
    'Set PYTHON to override the Python interpreter used to create the virtualenv.',
    '',
  ].join('\n'))
}

function runInherited(command, args, cwd = ROOT) {
  const result = spawnSync(command, args, {
    cwd,
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

function commandResult(command, args) {
  return spawnSync(command, args, {
    cwd: ROOT,
    encoding: 'utf8',
    maxBuffer: 1024 * 1024,
    env: process.env,
  })
}

function requireCommand(command, args, prerequisite) {
  const result = commandResult(command, args)
  if (result.error) {
    if (result.error.code === 'ENOENT') {
      throw new Error(`${prerequisite} is required but '${command}' was not found`)
    }
    throw result.error
  }
  if (result.status !== 0) {
    throw new Error(
      `${prerequisite} check failed with status ${result.status}: ` +
        `${result.stderr || result.stdout}`,
    )
  }
  return (result.stdout || result.stderr).trim()
}

function resolveSystemPython() {
  const candidates = [process.env.PYTHON, 'python3', 'python'].filter(Boolean)
  for (const candidate of candidates) {
    const result = commandResult(candidate, ['--version'])
    if (!result.error && result.status === 0) {
      return candidate
    }
  }
  throw new Error(
    `Python 3.10+ is required for workbook validation, but none of ` +
      `${candidates.join(', ')} was runnable`,
  )
}

function checkPrerequisites(config, systemPython) {
  requireCommand(systemPython, ['--version'], 'Python 3.10+')
  if (config.implementations.includes('node') && !config.skipBuild) {
    requireCommand(NPM, ['--version'], 'npm')
  }
  if (config.implementations.includes('go')) {
    requireCommand('go', ['version'], 'Go 1.25+')
  }
  if (config.implementations.includes('csharp')) {
    requireCommand('dotnet', ['--version'], '.NET SDK 8')
  }
}

function buildImplementations(config, systemPython) {
  if (config.implementations.includes('node')) {
    process.stdout.write('Building the N-API package in release mode...\n')
    runInherited(NPM, ['run', 'build'])
  }
  if (config.implementations.includes('go')) {
    process.stdout.write('Building the Go Excelize generator...\n')
    fs.mkdirSync(path.dirname(GO_BINARY), { recursive: true })
    runInherited(
      'go',
      ['build', '-trimpath', '-ldflags=-s -w', '-o', GO_BINARY, '.'],
      GO_DIRECTORY,
    )
  }
  if (config.implementations.includes('csharp')) {
    process.stdout.write('Building the C# Open XML SDK generator...\n')
    runInherited('dotnet', [
      'build',
      CSHARP_PROJECT,
      '--configuration',
      'Release',
      '--nologo',
    ])
  }
  if (config.implementations.includes('python')) {
    process.stdout.write('Preparing the Python XlsxWriter virtual environment...\n')
    runInherited(systemPython, ['-m', 'venv', PYTHON_VENV])
    runInherited(PYTHON_VENV_BINARY, [
      '-m',
      'pip',
      'install',
      '--disable-pip-version-check',
      '--requirement',
      PYTHON_REQUIREMENTS,
    ])
  }
}

function assertBuildArtifacts(config) {
  const expectedArtifacts = []
  if (config.implementations.includes('go')) {
    expectedArtifacts.push(['Go release binary', GO_BINARY])
  }
  if (config.implementations.includes('csharp')) {
    expectedArtifacts.push(['C# release assembly', CSHARP_BINARY])
  }
  if (config.implementations.includes('python')) {
    expectedArtifacts.push(['Python virtualenv interpreter', PYTHON_VENV_BINARY])
  }
  for (const [name, filename] of expectedArtifacts) {
    if (!fs.existsSync(filename)) {
      throw new Error(`${name} not found: ${filename}`)
    }
  }
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
  try {
    return { ...JSON.parse(resultLine), processWallMs }
  } catch {
    throw new Error(`Runner returned invalid JSON:\n${result.stdout}`)
  }
}

function implementationCommand(implementation, commonArgs, config) {
  switch (implementation) {
    case 'node':
      return [process.execPath, [
        NODE_RUNNER,
        ...commonArgs,
        '--write-mode',
        config.nodeWriteMode,
        '--batch-rows',
        String(config.batchRows),
      ]]
    case 'go':
      return [GO_BINARY, commonArgs]
    case 'csharp':
      return ['dotnet', [CSHARP_BINARY, ...commonArgs]]
    case 'python':
      return [PYTHON_VENV_BINARY, [PYTHON_RUNNER, ...commonArgs]]
    default:
      throw new Error(`Unsupported implementation: ${implementation}`)
  }
}

function validateRunnerResult(result, implementation, config, output) {
  const expectedCells = config.rows * config.columns
  if (
    result.rows !== config.rows ||
    result.columns !== config.columns ||
    result.cells !== expectedCells
  ) {
    throw new Error(`${implementation} reported an unexpected workload`)
  }
  for (const metric of ['writeMs', 'saveMs', 'totalMs', 'processWallMs']) {
    if (!Number.isFinite(result[metric]) || result[metric] < 0) {
      throw new Error(`${implementation} reported an invalid ${metric}`)
    }
  }
  if (
    result.maxRssBytes != null &&
    (!Number.isFinite(result.maxRssBytes) || result.maxRssBytes <= 0)
  ) {
    throw new Error(`${implementation} reported an invalid maxRssBytes`)
  }
  const outputBytes = fs.statSync(output).size
  if (result.outputBytes !== outputBytes || outputBytes === 0) {
    throw new Error(`${implementation} reported an invalid output size`)
  }
  return { ...result, id: implementation }
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
  const [command, args] = implementationCommand(implementation, commonArgs, config)
  return validateRunnerResult(
    executeRunner(command, args),
    implementation,
    config,
    output,
  )
}

function validateWorkbook(systemPython, filename, rows, columns) {
  executeRunner(systemPython, [
    WORKBOOK_VALIDATOR,
    '--file',
    filename,
    '--rows',
    String(rows),
    '--columns',
    String(columns),
  ])
}

function validateImplementations(config, systemPython, workDirectory) {
  const validationConfig = { ...config, rows: Math.min(config.rows, 17) }
  const metadata = {}
  process.stdout.write(
    `Validating ${validationConfig.rows} rows × ${validationConfig.columns} columns ` +
      'semantically before measurement...\n',
  )

  for (const implementation of config.implementations) {
    const output = path.join(workDirectory, `validation-${implementation}.xlsx`)
    process.stdout.write(`  ${implementation}... `)
    metadata[implementation] = runImplementation(
      implementation,
      validationConfig,
      output,
    )
    validateWorkbook(
      systemPython,
      output,
      validationConfig.rows,
      validationConfig.columns,
    )
    fs.rmSync(output, { force: true })
    process.stdout.write('valid\n')
  }
  return metadata
}

function balancedOrder(implementations, run) {
  const indexes = [0]
  for (let offset = 1; indexes.length < implementations.length; offset++) {
    indexes.push(offset)
    if (indexes.length < implementations.length) {
      indexes.push(implementations.length - offset)
    }
  }
  return indexes.map(
    (index) => implementations[(index + run) % implementations.length],
  )
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

function displayName(result) {
  const version = result.libraryVersion ? ` ${result.libraryVersion}` : ''
  return `${result.implementation}${version}`
}

function printResults(results, config) {
  const measuredRows = []
  for (let run = 0; run < config.runs; run++) {
    for (const implementation of config.implementations) {
      const result = results[implementation][run]
      measuredRows.push({
        Run: run + 1,
        Implementation: displayName(result),
        Generation: formatMilliseconds(result.totalMs),
        'Process wall': formatMilliseconds(result.processWallMs),
        'Cells/s': formatRate(result.cells, result.totalMs),
        'Max RSS': formatMiB(result.maxRssBytes),
        XLSX: formatMiB(result.outputBytes),
      })
    }
  }

  process.stdout.write('\nMeasured runs:\n')
  console.table(measuredRows)

  const summaries = {}
  for (const implementation of config.implementations) {
    const implementationResults = results[implementation]
    summaries[implementation] = {
      implementation: implementationResults[0].implementation,
      libraryVersion: implementationResults[0].libraryVersion,
      totalMs: median(implementationResults.map((result) => result.totalMs)),
      processWallMs: median(
        implementationResults.map((result) => result.processWallMs),
      ),
      maxRssBytes: median(
        implementationResults
          .map((result) => result.maxRssBytes)
          .filter((value) => value != null),
      ),
      outputBytes: median(
        implementationResults.map((result) => result.outputBytes),
      ),
    }
  }

  const nodeTotalMs = summaries.node?.totalMs
  const summaryRows = config.implementations.map((implementation) => {
    const summary = summaries[implementation]
    return {
      Implementation: displayName(summary),
      'Median generation': formatMilliseconds(summary.totalMs),
      'Median process wall': formatMilliseconds(summary.processWallMs),
      'Median cells/s': formatRate(
        config.rows * config.columns,
        summary.totalMs,
      ),
      'Median max RSS': formatMiB(summary.maxRssBytes),
      'Median XLSX': formatMiB(summary.outputBytes),
      'Generation / Node': nodeTotalMs == null
        ? 'n/a'
        : `${(summary.totalMs / nodeTotalMs).toFixed(2)}x`,
    }
  })

  process.stdout.write('\nMedian summary:\n')
  console.table(summaryRows)
}

function printEnvironment(config, metadata) {
  process.stdout.write('\nCross-language XLSX generation benchmark\n')
  process.stdout.write(`Rows: ${config.rows.toLocaleString('en-US')}\n`)
  process.stdout.write(`Columns: ${config.columns.toLocaleString('en-US')}\n`)
  process.stdout.write(
    `Cells: ${(config.rows * config.columns).toLocaleString('en-US')}\n`,
  )
  process.stdout.write('Dataset: 50% strings, 50% numbers\n')
  process.stdout.write('Mode: bounded-memory, sequential-row generation\n')
  process.stdout.write(`Runs: ${config.runs}\n`)
  process.stdout.write(`CPU: ${os.cpus()[0]?.model || 'unknown'}\n`)
  process.stdout.write('Implementations:\n')
  for (const implementation of config.implementations) {
    const result = metadata[implementation]
    process.stdout.write(
      `  ${IMPLEMENTATION_LABELS[implementation]}: ` +
        `${result.runtimeVersion || 'unknown runtime'}, ` +
        `${result.libraryVersion || 'unknown library version'}\n`,
    )
  }
  process.stdout.write(
    'Generation excludes process startup; process wall includes it. ' +
      'Build and semantic validation are excluded.\n\n',
  )
}

function main() {
  const config = parseArgs(process.argv.slice(2))
  if (config.help) {
    printHelp()
    return
  }

  const systemPython = resolveSystemPython()
  checkPrerequisites(config, systemPython)
  if (!config.skipBuild) {
    buildImplementations(config, systemPython)
  }
  assertBuildArtifacts(config)

  const workDirectory = fs.mkdtempSync(
    path.join(os.tmpdir(), 'rusc-xlsx-language-benchmark-'),
  )
  const results = Object.fromEntries(
    config.implementations.map((implementation) => [implementation, []]),
  )

  try {
    const metadata = validateImplementations(
      config,
      systemPython,
      workDirectory,
    )
    printEnvironment(config, metadata)

    for (let run = 0; run < config.runs; run++) {
      const order = balancedOrder(config.implementations, run)
      for (const implementation of order) {
        const output = path.join(
          workDirectory,
          `run-${run + 1}-${implementation}.xlsx`,
        )
        process.stdout.write(
          `[run ${run + 1}/${config.runs}] ${implementation}... `,
        )
        const result = runImplementation(implementation, config, output)
        results[implementation].push(result)
        process.stdout.write(`${formatMilliseconds(result.totalMs)}\n`)
        if (!config.keepFiles) {
          fs.rmSync(output, { force: true })
        }
      }
    }

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
  process.stderr.write(`Language benchmark failed: ${error.stack || error.message}\n`)
  process.exitCode = 1
}
