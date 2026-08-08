# Гіперпосилання (`url`)

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> Pinned baseline: `rust_xlsxwriter 0.97.1`; **4 grouped capabilities**, **4 canonical symbol rows**.

- Grouped: **0 full · 3 partial · 1 missing · 0 N/A**.
- Symbol rows: **0 full · 3 partial · 1 missing · 0 N/A**.
- `[x] FULL`; `[ ] PARTIAL`; `[ ] MISSING`; plain `N/A` is outside the compiled/runtime denominator.
- Before marking `FULL`, satisfy the [shared definition of done](../rust-xlsxwriter-coverage-checklist.md#definition-of-done).

## `Url`

- [ ] **PARTIAL** `rust_xlsxwriter::url::Url#struct` — evidence: `static-forwarding` — gap: The Rust builder/value type is not exposed, but a narrowed facade helper reaches part of its XLSX output capability. untested <!-- coverage-capability:rust_xlsxwriter::url::Url#struct --><!-- coverage-row:rust_xlsxwriter::url::Url#struct -->
- [ ] **PARTIAL** `rust_xlsxwriter::url::Url#struct::new#function` — → `Image.prototype.setUrl`, `RuscImage.prototype.setUrl` — evidence: `static-forwarding` — tests: `tests/images.test.js:213`, `tests/hyperlinks-filters.test.js:26`, `tests/hyperlinks-filters.test.js:54` — gap: A local adapter path reaches the upstream operation, but option space, error behavior, builder access, or fluent/result semantics are narrowed. URL writing/text is available, but Url is not exposed as a JS value object and tip/object methods are incomplete. <!-- coverage-capability:rust_xlsxwriter::url::Url#struct::new#function --><!-- coverage-row:rust_xlsxwriter::url::Url#struct::new#function -->
- [ ] **PARTIAL** `rust_xlsxwriter::url::Url#struct::set_text#function` — → `Worksheet.prototype.writeUrlWithText`, `RuscWorksheet.prototype.writeUrlWithText` — evidence: `static-forwarding` — tests: `tests/hyperlinks-filters.test.js:38`, `tests/hyperlinks-filters.test.js:71`, `tests/hyperlinks-filters.test.js:84` — gap: A local adapter path reaches the upstream operation, but option space, error behavior, builder access, or fluent/result semantics are narrowed. URL writing/text is available, but Url is not exposed as a JS value object and tip/object methods are incomplete. <!-- coverage-capability:rust_xlsxwriter::url::Url#struct::set_text#function --><!-- coverage-row:rust_xlsxwriter::url::Url#struct::set_text#function -->
- [ ] **MISSING** `rust_xlsxwriter::url::Url#struct::set_tip#function` — evidence: `runtime` <!-- coverage-capability:rust_xlsxwriter::url::Url#struct::set_tip#function --><!-- coverage-row:rust_xlsxwriter::url::Url#struct::set_tip#function -->
