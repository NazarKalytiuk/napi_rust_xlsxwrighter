'use strict';

// The workloads of rust_xlsxwriter's benches/write_performance.rs (v0.102.0)
// for the JavaScript runners; benchmarks/pure-rust/src/bin/write_performance.rs
// is the Rust copy.

// All configurations use a fixed column count and vary the row count.
const COLS = 50;

// [rows, name] for 1K, 10K, 100K and 1M cells.
const CONFIGS = [
  [20, '1K cells'],
  [200, '10K cells'],
  [2_000, '100K cells'],
  [20_000, '1M cells'],
];

// The configuration whose memory use is measured.
const MEMORY_ROWS = 20_000;

// [group, value of the cell at (row, col)] in upstream's `criterion_group!` order.
const GROUPS = [
  ['numbers', () => 12345.0],
  ['mixed', (_row, col) => (col % 2 === 1 ? 'Foo' : 12345.0)],
  ['string_static', () => 'Foo'],
  ['string_dynamic', (row, col) => `Cell (${row}, ${col})`],
];

const MODES = ['default', 'constant', 'low'];

// Fewer samples for the larger, slower configs.
function sampleSize(cellCount) {
  if (cellCount >= 1_000_000) {
    return 10;
  }
  if (cellCount >= 100_000) {
    return 20;
  }
  return 50;
}

function valueAt(group) {
  const entry = GROUPS.find(([name]) => name === group);
  if (!entry) {
    throw new Error(`Unknown group: ${group}`);
  }
  return entry[1];
}

function workbookName(group, cellCount) {
  return `${group}-${cellCount}.xlsx`;
}

module.exports = { COLS, CONFIGS, GROUPS, MEMORY_ROWS, MODES, sampleSize, valueAt, workbookName };
