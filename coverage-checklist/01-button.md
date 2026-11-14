# Кнопки (`button`)

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> Pinned baseline: `rust_xlsxwriter 0.102.0`; **8 grouped capabilities**, **8 canonical symbol rows**.

- Grouped: **8 full · 0 partial · 0 missing · 0 N/A**.
- Symbol rows: **8 full · 0 partial · 0 missing · 0 N/A**.
- `[x] FULL`; `[ ] PARTIAL`; `[ ] MISSING`; plain `N/A` is outside the compiled/runtime denominator.
- Before marking `FULL`, satisfy the [shared definition of done](../rust-xlsxwriter-coverage-checklist.md#definition-of-done).

## `Button`

- [x] **FULL** `rust_xlsxwriter::button::Button#struct` — → `Button class`, `RuscButton class/factory` — evidence: `runtime` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct --><!-- coverage-row:rust_xlsxwriter::button::Button#struct -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::new#function` — → `Button::constructor`, `RuscButton.static.new` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:143`, `tests/vba-shapes-metadata.test.mjs:149`, `tests/vba-shapes-metadata.test.mjs:166` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::new#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::new#function -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::set_alt_text#function` — → `Button.prototype.setAltText`, `RuscButton.prototype.setAltText` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:151` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::set_alt_text#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::set_alt_text#function -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::set_caption#function` — → `Button.prototype.setCaption`, `RuscButton.prototype.setCaption` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:151`, `tests/vba-shapes-metadata.test.mjs:167`, `tests/vba-shapes-metadata.test.mjs:181` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::set_caption#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::set_caption#function -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::set_height#function` — → `Button.prototype.setHeight`, `RuscButton.prototype.setHeight` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:151`, `tests/vba-shapes-metadata.test.mjs:167`, `tests/vba-shapes-metadata.test.mjs:181` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::set_height#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::set_height#function -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::set_macro#function` — → `Button.prototype.setMacro`, `RuscButton.prototype.setMacro` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:151`, `tests/vba-shapes-metadata.test.mjs:167`, `tests/vba-shapes-metadata.test.mjs:226` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::set_macro#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::set_macro#function -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::set_object_movement#function` — → `Button.prototype.setObjectMovement`, `RuscButton.prototype.setObjectMovement` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:195`, `tests/vba-shapes-metadata.test.mjs:199`, `tests/vba-shapes-metadata.test.mjs:203` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::set_object_movement#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::set_object_movement#function -->
- [x] **FULL** `rust_xlsxwriter::button::Button#struct::set_width#function` — → `Button.prototype.setWidth`, `RuscButton.prototype.setWidth` — evidence: `static-forwarding` — tests: `tests/vba-shapes-metadata.test.mjs:151`, `tests/vba-shapes-metadata.test.mjs:167`, `tests/vba-shapes-metadata.test.mjs:181` <!-- coverage-capability:rust_xlsxwriter::button::Button#struct::set_width#function --><!-- coverage-row:rust_xlsxwriter::button::Button#struct::set_width#function -->
