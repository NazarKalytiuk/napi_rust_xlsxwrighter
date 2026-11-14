import { expect, test } from 'vitest';

import { type RustdocCrate, buildInventory } from './inventory.mts';

// A minimal rustdoc JSON crate: a struct with a public field and method, an enum
// with a payload variant, and a trait with a concrete and a generic impl.
const item = (id: number, name: string | null, inner: object, visibility = 'public') => ({
  id,
  crate_id: 0,
  name,
  visibility,
  inner,
});
const path = (type: object) => ({ resolved_path: { path: '', id: 99, args: null }, ...type });
const sig = (output: object | null) => ({
  function: {
    sig: {
      inputs: [['self', { borrowed_ref: { is_mutable: true, type: { generic: 'Self' } } }]],
      output,
    },
  },
});

const krate: RustdocCrate = {
  root: 0,
  crate_version: '1.0.0',
  format_version: 61,
  paths: {
    1: { crate_id: 0, path: ['rust_xlsxwriter', 'note', 'Note'], kind: 'struct' },
    5: { crate_id: 0, path: ['rust_xlsxwriter', 'error', 'XlsxError'], kind: 'enum' },
    8: { crate_id: 0, path: ['rust_xlsxwriter', 'worksheet', 'IntoExcelData'], kind: 'trait' },
    99: { crate_id: 1, path: ['alloc', 'string', 'String'], kind: 'struct' },
  },
  index: {
    1: item(1, 'Note', { struct: { kind: { plain: { fields: [2, 3] } }, impls: [4] } }),
    2: item(2, 'width', { struct_field: { primitive: 'u32' } }),
    3: item(3, 'author', { struct_field: { primitive: 'u32' } }, 'crate'),
    4: item(4, null, { impl: { trait: null, items: [40, 41], generics: { params: [] } } }),
    40: item(40, 'set_text', sig(null)),
    41: item(41, 'private_helper', sig(null), 'crate'),
    5: item(5, 'XlsxError', { enum: { variants: [6, 7], impls: [] } }),
    6: item(6, 'ParameterError', { variant: { kind: { tuple: [60] } } }, 'default'),
    60: item(60, '0', { struct_field: path({}) }, 'default'),
    7: item(7, 'Unknown', { variant: { kind: 'plain' } }, 'default'),
    8: item(8, 'IntoExcelData', { trait: { items: [80], implementations: [81, 82] } }),
    80: item(80, 'write', sig(null), 'default'),
    81: item(81, null, {
      impl: {
        for: { borrowed_ref: { is_mutable: false, type: path({}) } },
        generics: { params: [] },
      },
    }),
    82: item(82, null, {
      impl: { for: { generic: 'T' }, generics: { params: [{ name: 'T', kind: { type: {} } }] } },
    }),
  },
} as unknown as RustdocCrate;

test('builds canonical rows and groups variant payloads under their variant', () => {
  const rows = buildInventory(krate);
  expect([...rows.keys()].toSorted()).toEqual([
    'rust_xlsxwriter::error::XlsxError#enum',
    'rust_xlsxwriter::error::XlsxError#enum::ParameterError#variant',
    'rust_xlsxwriter::error::XlsxError#enum::ParameterError#variant::0#field',
    'rust_xlsxwriter::error::XlsxError#enum::Unknown#variant',
    'rust_xlsxwriter::note::Note#struct',
    'rust_xlsxwriter::note::Note#struct::set_text#function',
    'rust_xlsxwriter::note::Note#struct::width#field',
    'rust_xlsxwriter::worksheet::IntoExcelData#trait',
    // The generic `impl<T> IntoExcelData for T` has no row of its own.
    'rust_xlsxwriter::worksheet::IntoExcelData#trait-impl-for:&alloc::string::String',
    'rust_xlsxwriter::worksheet::IntoExcelData#trait::write#function',
  ]);
  expect(
    rows.get('rust_xlsxwriter::error::XlsxError#enum::ParameterError#variant::0#field'),
  ).toMatchObject({
    capability: 'rust_xlsxwriter::error::XlsxError#enum::ParameterError#variant',
    module: 'error',
    section: 'XlsxError',
    signature: 'alloc::string::String',
  });
  expect(rows.get('rust_xlsxwriter::note::Note#struct::set_text#function')!.signature).toBe(
    '(&mut Self)->()',
  );
});
