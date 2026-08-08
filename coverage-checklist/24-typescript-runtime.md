# TypeScript/runtime integrity backlog

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> **8 open declaration/runtime mismatches.** Це окремі blockers для чесного type-safe package contract.

- [ ] **signature-mismatch** `RuscButton::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:3`. <!-- coverage-typescript:RuscButton::constructor -->
- [ ] **signature-mismatch** `RuscDocProperties::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:77`. <!-- coverage-typescript:RuscDocProperties::constructor -->
- [ ] **signature-mismatch** `RuscImage::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:187`. <!-- coverage-typescript:RuscImage::constructor -->
- [ ] **signature-mismatch** `RuscShape::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:284`. <!-- coverage-typescript:RuscShape::constructor -->
- [ ] **signature-mismatch** `RuscSparkline::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:299`. <!-- coverage-typescript:RuscSparkline::constructor -->
- [ ] **signature-mismatch** `RuscTable::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:354`. <!-- coverage-typescript:RuscTable::constructor -->
- [ ] **signature-mismatch** `RuscTableColumn::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:377`. <!-- coverage-typescript:RuscTableColumn::constructor -->
- [ ] **signature-mismatch** `RuscWorksheet::constructor` — TypeScript permits construction but runtime rejects it: Class contains no `constructor`, can not new it! Declarations: `index.d.ts:436`. <!-- coverage-typescript:RuscWorksheet::constructor -->
