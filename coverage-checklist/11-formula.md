# Формули (`formula`)

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> Pinned baseline: `rust_xlsxwriter 0.104.0`; **3 grouped capabilities**, **3 canonical symbol rows**.

- Grouped: **3 full · 0 partial · 0 missing · 0 N/A**.
- Symbol rows: **3 full · 0 partial · 0 missing · 0 N/A**.
- `[x] FULL`; `[ ] PARTIAL`; `[ ] MISSING`; plain `N/A` is outside the compiled/runtime denominator.
- Before marking `FULL`, satisfy the [shared definition of done](../rust-xlsxwriter-coverage-checklist.md#definition-of-done).

## `Formula`

- [x] **FULL** `rust_xlsxwriter::formula::Formula#struct` — → `Formula class`, `RuscFormula class/factory` — evidence: `runtime` <!-- coverage-capability:rust_xlsxwriter::formula::Formula#struct --><!-- coverage-row:rust_xlsxwriter::formula::Formula#struct -->
- [x] **FULL** `rust_xlsxwriter::formula::Formula#struct::new#function` — → `Formula::constructor`, `RuscFormula::constructor` — evidence: `static-forwarding` — tests: `tests/formulas.test.mjs:27`, `tests/formulas.test.mjs:32`, `tests/formulas.test.mjs:38` <!-- coverage-capability:rust_xlsxwriter::formula::Formula#struct::new#function --><!-- coverage-row:rust_xlsxwriter::formula::Formula#struct::new#function -->
- [x] **FULL** `rust_xlsxwriter::formula::Formula#struct::set_result#function` — → `Formula.prototype.setResult`, `RuscFormula.prototype.setResult` — evidence: `differential-runtime` — tests: `tests/formulas.test.mjs:33`, `tests/formulas.test.mjs:38`, `tests/formulas.test.mjs:116` <!-- coverage-capability:rust_xlsxwriter::formula::Formula#struct::set_result#function --><!-- coverage-row:rust_xlsxwriter::formula::Formula#struct::set_result#function -->
