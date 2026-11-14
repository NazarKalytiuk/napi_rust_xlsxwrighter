import { expect, test } from 'vitest';

import type { Capability, Group } from './checklist.mts';
import { requireTestedFullRows } from './generate.mts';

function groupOf(...capabilities: Partial<Capability>[]): Group {
  return {
    module: 'table',
    title: 'Таблиці',
    capabilities: capabilities.map((capability, index) => ({
      id: `rust_xlsxwriter::table::Row${index}#struct`,
      status: 'FULL',
      mapping: [],
      evidence: 'static-forwarding',
      tests: [],
      remarks: [],
      nested: null,
      ...capability,
    })),
  };
}

test('requireTestedFullRows rejects FULL rows whose mapped APIs no test calls', () => {
  const group = groupOf(
    { mapping: ['Table.prototype.setStyle(0)'], tests: ['tests/tables.test.mjs:9'] },
    { mapping: ['Table.prototype.setStyle(61)'] },
    { status: 'PARTIAL', mapping: ['Table.prototype.setName'] },
    { mapping: ['Table class'] },
  );
  expect(() => requireTestedFullRows([group])).toThrow(
    'rust_xlsxwriter::table::Row1#struct → Table.prototype.setStyle(61)',
  );
});

test('requireTestedFullRows accepts FULL rows that have call sites or only prose mappings', () => {
  const group = groupOf(
    { mapping: ['Workbook.prototype.save'], tests: ['tests/workbook.test.mjs:4'] },
    { mapping: ['Workbook class'] },
  );
  expect(() => requireTestedFullRows([group])).not.toThrow();
});
