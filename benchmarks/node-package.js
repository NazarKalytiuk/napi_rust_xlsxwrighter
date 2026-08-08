'use strict'

const fs = require('fs')
const path = require('path')
const { performance } = require('perf_hooks')
const { Workbook } = require('../wrapper.js')
const packageJson = require('../package.json')

const DEFAULT_ROWS = 200_000
const DEFAULT_COLUMNS = 20
const MAX_EXCEL_ROWS = 1_048_576
const MAX_EXCEL_COLUMNS = 16_384
const DEFAULT_BATCH_ROWS = 1_000

function parsePositiveInteger(value, name, maximum) {
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed) || parsed <= 0) {
    throw new Error(`${name} must be a positive integer`)
  }
  if (parsed > maximum) {
    throw new Error(`${name} must not exceed Excel's limit of ${maximum}`)
  }
  return parsed
}

function parseArgs(argv) {
  const config = {
    rows: DEFAULT_ROWS,
    columns: DEFAULT_COLUMNS,
    output: path.resolve('node-package-generation.xlsx'),
    writeMode: 'cell',
    batchRows: DEFAULT_BATCH_ROWS,
  }

  for (let index = 0; index < argv.length; index += 2) {
    const flag = argv[index]
    const value = argv[index + 1]
    if (value === undefined) {
      throw new Error(`Missing value for ${flag}`)
    }

    switch (flag) {
      case '--rows':
        config.rows = parsePositiveInteger(value, 'rows', MAX_EXCEL_ROWS)
        break
      case '--columns':
        config.columns = parsePositiveInteger(value, 'columns', MAX_EXCEL_COLUMNS)
        break
      case '--output':
        config.output = path.resolve(value)
        break
      case '--write-mode':
        if (!['cell', 'row', 'batch'].includes(value)) {
          throw new Error('write-mode must be one of: cell, row, batch')
        }
        config.writeMode = value
        break
      case '--batch-rows':
        config.batchRows = parsePositiveInteger(value, 'batch-rows', MAX_EXCEL_ROWS)
        break
      default:
        throw new Error(`Unknown argument: ${flag}`)
    }
  }

  return config
}

function makeRow(row, columns) {
  const values = new Array(columns)
  for (let column = 0; column < columns; column++) {
    values[column] = column % 2 === 0
      ? `Row ${row} Col ${column}`
      : row * column * 0.123
  }
  return values
}

function writeData(worksheet, config) {
  if (config.writeMode === 'cell') {
    for (let row = 0; row < config.rows; row++) {
      for (let column = 0; column < config.columns; column++) {
        worksheet.write(
          row,
          column,
          column % 2 === 0
            ? `Row ${row} Col ${column}`
            : row * column * 0.123,
        )
      }
    }
    return
  }

  if (config.writeMode === 'row') {
    for (let row = 0; row < config.rows; row++) {
      worksheet.writeRow(row, 0, makeRow(row, config.columns))
    }
    return
  }

  for (let startRow = 0; startRow < config.rows; startRow += config.batchRows) {
    const rowCount = Math.min(config.batchRows, config.rows - startRow)
    const rows = new Array(rowCount)
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      rows[rowOffset] = makeRow(startRow + rowOffset, config.columns)
    }
    worksheet.writeRows(startRow, 0, rows)
  }
}

function run(config) {
  const outputDirectory = path.dirname(config.output)
  fs.mkdirSync(outputDirectory, { recursive: true })

  const totalStart = performance.now()
  const workbook = new Workbook()
  workbook.setTempdir(outputDirectory)
  const worksheet = workbook.addWorksheetWithConstantMemory('Data')

  const writeStart = performance.now()
  writeData(worksheet, config)
  const writeMs = performance.now() - writeStart

  const saveStart = performance.now()
  workbook.save(config.output)
  const saveMs = performance.now() - saveStart
  const totalMs = performance.now() - totalStart

  const cells = config.rows * config.columns
  const result = {
    implementation: config.writeMode === 'batch'
      ? `node-package/batch-${config.batchRows}`
      : `node-package/${config.writeMode}`,
    libraryVersion: packageJson.version,
    runtimeVersion: process.version,
    mode: 'constant-memory',
    writeMode: config.writeMode,
    batchRows: config.writeMode === 'batch' ? config.batchRows : null,
    rows: config.rows,
    columns: config.columns,
    cells,
    writeMs,
    saveMs,
    totalMs,
    maxRssBytes: process.resourceUsage().maxRSS * 1_024,
    outputBytes: fs.statSync(config.output).size,
  }

  process.stdout.write(`${JSON.stringify(result)}\n`)
}

try {
  if (process.argv.includes('--help') || process.argv.includes('-h')) {
    process.stdout.write(
      'Usage: node node-package.js [--rows N] [--columns N] [--output PATH] ' +
        '[--write-mode cell|row|batch] [--batch-rows N]\n',
    )
  } else {
    run(parseArgs(process.argv.slice(2)))
  }
} catch (error) {
  process.stderr.write(`Node package benchmark failed: ${error.stack || error.message}\n`)
  process.exitCode = 1
}
