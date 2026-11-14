import { expect, test } from 'vitest';

import { countCapabilities, countRows, parseGroup, renderGroup } from './checklist.mts';

const GROUP = `# Кольори (\`color\`)

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> Pinned baseline: \`rust_xlsxwriter 1.0.0\`; **3 grouped capabilities**, **4 canonical symbol rows**.

- Grouped: **1 full · 1 partial · 0 missing · 1 N/A**.
- Symbol rows: **2 full · 1 partial · 0 missing · 1 N/A**.
- \`[x] FULL\`; \`[ ] PARTIAL\`; \`[ ] MISSING\`; plain \`N/A\` is outside the compiled/runtime denominator.
- Before marking \`FULL\`, satisfy the [shared definition of done](../rust-xlsxwriter-coverage-checklist.md#definition-of-done).

## \`Color\`

- [x] **FULL** \`rust_xlsxwriter::color::Color#enum\` — → \`Format.prototype.setFontColor\` — evidence: \`runtime\` — tests: \`tests/a.test.mjs:3\` <!-- coverage-capability:rust_xlsxwriter::color::Color#enum --><!-- coverage-row:rust_xlsxwriter::color::Color#enum -->
- [ ] **PARTIAL** \`rust_xlsxwriter::color::Color#enum::RGB#variant\` — evidence: \`static-forwarding\` — gap: Hex only — no theme colors. <!-- coverage-capability:rust_xlsxwriter::color::Color#enum::RGB#variant -->
  - **PARTIAL** \`rust_xlsxwriter::color::Color#enum::RGB#variant\` <!-- coverage-row:rust_xlsxwriter::color::Color#enum::RGB#variant -->
  - **FULL** \`rust_xlsxwriter::color::Color#enum::RGB#variant::0#field\` <!-- coverage-row:rust_xlsxwriter::color::Color#enum::RGB#variant::0#field -->

## \`IntoColor\`

- **N/A** \`rust_xlsxwriter::color::IntoColor#trait-impl-for:&mut rust_xlsxwriter::color::Color\` — gap: Compile-time only. <!-- coverage-capability:rust_xlsxwriter::color::IntoColor#trait-impl-for:&mut rust_xlsxwriter::color::Color --><!-- coverage-row:rust_xlsxwriter::color::IntoColor#trait-impl-for:&mut rust_xlsxwriter::color::Color -->
`;

test('parses audited fields and renders the group back unchanged', () => {
  const group = parseGroup(GROUP, '03-color.md');
  expect(group).toMatchObject({ module: 'color', title: 'Кольори' });
  const [color, rgb, impl] = group.capabilities;
  expect(color).toMatchObject({
    status: 'FULL',
    mapping: ['Format.prototype.setFontColor'],
    evidence: 'runtime',
    tests: ['tests/a.test.mjs:3'],
    nested: null,
  });
  // " — " inside free text stays part of the remark.
  expect(rgb!.remarks).toEqual([{ label: 'gap', text: 'Hex only — no theme colors.' }]);
  expect(rgb!.nested).toHaveLength(2);
  expect(impl!.id).toContain('&mut rust_xlsxwriter::color::Color');
  expect(renderGroup(group, '1.0.0')).toBe(GROUP);
});

test('counts nested rows separately from grouped capabilities', () => {
  const { capabilities } = parseGroup(GROUP, '03-color.md');
  expect(countCapabilities(capabilities)).toEqual({ FULL: 1, PARTIAL: 1, MISSING: 0, 'N/A': 1 });
  expect(countRows(capabilities)).toEqual({ FULL: 2, PARTIAL: 1, MISSING: 0, 'N/A': 1 });
});

test('rejects rows whose markers disagree with their id', () => {
  const broken = GROUP.replace(
    'coverage-capability:rust_xlsxwriter::color::Color#enum -->',
    'coverage-capability:x -->',
  );
  expect(() => parseGroup(broken, '03-color.md')).toThrow('03-color.md:14: markers do not match');
});
