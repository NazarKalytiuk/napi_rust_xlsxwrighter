'use strict';

// The workloads of the write benchmark for the JavaScript runners;
// benchmarks/pure-rust/src/bin/write_performance.rs is the Rust copy.
//
// - write-performance: rust_xlsxwriter's benches/write_performance.rs
//   (v0.102.0), four groups at 1K to 1M cells in 50 columns.
// - export: a business-report export (benchmarks/export-workload.js) of
//   30 columns of mixed data, 200,000 rows unless given.

const { COLUMNS: EXPORT_COLUMNS, exportValue } = require('./export-workload.js');

const WORKLOADS = ['write-performance', 'export'];
const MODES = ['default', 'constant', 'low'];
const DEFAULT_EXPORT_ROWS = 200_000;

// [group, value of the cell at (row, col)] in upstream's `criterion_group!` order.
const UPSTREAM_GROUPS = [
  ['numbers', () => 12345.0],
  ['mixed', (_row, col) => (col % 2 === 1 ? 'Foo' : 12345.0)],
  ['string_static', () => 'Foo'],
  ['string_dynamic', (row, col) => `Cell (${row}, ${col})`],
];

/**
 * A workload: its columns, its [rows, name] configurations, the rows whose
 * memory use is measured, its groups and whether it writes dates.
 */
function workload(name, rows = DEFAULT_EXPORT_ROWS) {
  switch (name) {
    case 'write-performance':
      return {
        name,
        columns: 50,
        configs: [
          [20, '1K cells'],
          [200, '10K cells'],
          [2_000, '100K cells'],
          [20_000, '1M cells'],
        ],
        memoryRows: 20_000,
        groups: UPSTREAM_GROUPS,
        hasDates: false,
      };
    case 'export':
      return {
        name,
        columns: EXPORT_COLUMNS,
        configs: [[rows, `${rows.toLocaleString('en-US')} rows`]],
        memoryRows: rows,
        groups: [['export', exportValue]],
        hasDates: true,
      };
    default:
      throw new Error(`Unknown workload: ${name}`);
  }
}

// Fewer samples for the larger, slower configs (upstream's sample sizes).
function sampleSize(cellCount) {
  if (cellCount >= 1_000_000) {
    return 10;
  }
  if (cellCount >= 100_000) {
    return 20;
  }
  return 50;
}

function valueAt(selected, group) {
  const entry = selected.groups.find(([name]) => name === group);
  if (!entry) {
    throw new Error(`Unknown group of ${selected.name}: ${group}`);
  }
  return entry[1];
}

function workbookName(group, cellCount) {
  return `${group}-${cellCount}.xlsx`;
}

module.exports = {
  DEFAULT_EXPORT_ROWS,
  MODES,
  WORKLOADS,
  sampleSize,
  valueAt,
  workbookName,
  workload,
};
