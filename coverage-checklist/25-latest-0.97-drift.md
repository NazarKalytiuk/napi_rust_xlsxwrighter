# Migration record: `0.90.2` → `0.97.1`

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> Historical diff: **17 added · 11 changed · 3 removed** stable rows. The changes are incorporated into the pinned `0.97.1` dashboard; this file is not counted separately.

## Помилки (`error`)

- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::DefaultFormatError#variant` — Errors cross the N-API boundary as JavaScript Error text; variants are not preserved structurally. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::DefaultFormatError#variant -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::DefaultFormatError#variant::0#field` — Error payloads cross the N-API boundary as text. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::DefaultFormatError#variant::0#field -->
- [ ] **PARTIAL / variant-changed** `rust_xlsxwriter::error::XlsxError#enum::IoError#variant::0#field` — The `0.97.1` payload is compiled, but N-API flattens it into JavaScript Error text. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::IoError#variant::0#field -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::ThemeError#variant` — Errors cross the N-API boundary as JavaScript Error text; variants are not preserved structurally. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::ThemeError#variant -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::ThemeError#variant::0#field` — Error payloads cross the N-API boundary as text. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::ThemeError#variant::0#field -->

## Форматування (`format`)

- [x] **FULL / added** `rust_xlsxwriter::format::FontScheme#enum` — Exposed as the `FontScheme` union and validated by `Format.setFontScheme()`. <!-- coverage-drift:rust_xlsxwriter::format::FontScheme#enum -->
- [x] **FULL / added** `rust_xlsxwriter::format::FontScheme#enum::Body#variant` — Maps to OOXML minor/body theme fonts. <!-- coverage-drift:rust_xlsxwriter::format::FontScheme#enum::Body#variant -->
- [x] **FULL / added** `rust_xlsxwriter::format::FontScheme#enum::Headings#variant` — Maps to OOXML major/headings theme fonts. <!-- coverage-drift:rust_xlsxwriter::format::FontScheme#enum::Headings#variant -->
- [x] **FULL / added** `rust_xlsxwriter::format::FontScheme#enum::None#variant` — Removes the OOXML font scheme. <!-- coverage-drift:rust_xlsxwriter::format::FontScheme#enum::None#variant -->
- [x] **FULL / signature-changed** `rust_xlsxwriter::format::Format#struct::set_font_scheme#function` — The facade and N-API adapter use the typed `0.97.1` enum contract. <!-- coverage-drift:rust_xlsxwriter::format::Format#struct::set_font_scheme#function -->

## Серіалізація (`serializer`)

- **N/A / signature-changed** `rust_xlsxwriter::serializer::CustomSerializeField#struct::set_column_width_pixels#function` — The updated signature is outside the compiled baseline because the optional `serde` feature is disabled. <!-- coverage-drift:rust_xlsxwriter::serializer::CustomSerializeField#struct::set_column_width_pixels#function -->

## Таблиці (`table`)

- [x] **FULL / added** `rust_xlsxwriter::table::Table#struct::set_alt_text#function` — Exposed through native and facade table APIs with OOXML evidence. <!-- coverage-drift:rust_xlsxwriter::table::Table#struct::set_alt_text#function -->
- [x] **FULL / added** `rust_xlsxwriter::table::Table#struct::set_alt_text_title#function` — Exposed through native and facade table APIs with OOXML evidence. <!-- coverage-drift:rust_xlsxwriter::table::Table#struct::set_alt_text_title#function -->

## Допоміжні функції (`utility`)

- [ ] **MISSING / signature-changed** `rust_xlsxwriter::utility::cell_autofit_width#function` — The `0.97.1` return type is pinned, but no public JavaScript utility export exists. <!-- coverage-drift:rust_xlsxwriter::utility::cell_autofit_width#function -->

## Workbook (`workbook`)

- [ ] **PARTIAL / signature-changed** `rust_xlsxwriter::workbook::Workbook#struct::save_to_writer#function` — `saveToStream()` uses a temporary file instead of the native Rust writer contract. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::save_to_writer#function -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::Workbook#struct::set_default_format#function` — Exposed before worksheet creation with OOXML and ordering-error tests. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::set_default_format#function -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::Workbook#struct::use_custom_theme#function` — Exposed with custom theme round-trip and error tests. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::use_custom_theme#function -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::Workbook#struct::use_excel_2023_theme#function` — Exposed with Excel 2023 theme OOXML and ordering-error tests. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::use_excel_2023_theme#function -->

## Worksheet (`worksheet`)

- **N/A / feature-gate-changed** `rust_xlsxwriter::worksheet::IntoExcelData#trait-impl-for:rust_decimal::decimal::Decimal` — Compile-time Rust trait integration has no independent JavaScript runtime operation. <!-- coverage-drift:rust_xlsxwriter::worksheet::IntoExcelData#trait-impl-for:rust_decimal::decimal::Decimal -->
- [x] **FULL / added** `rust_xlsxwriter::worksheet::Worksheet#struct::insert_image_fit_to_cell_centered#function` — Exposed through native and facade worksheet APIs with drawing-offset evidence. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::insert_image_fit_to_cell_centered#function -->
- [x] **FULL / added** `rust_xlsxwriter::worksheet::Worksheet#struct::set_autofit_max_row#function` — Exposed through native and facade worksheet APIs with row-boundary evidence. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_autofit_max_row#function -->
- [x] **FULL / added** `rust_xlsxwriter::worksheet::Worksheet#struct::set_autofit_max_width#function` — Exposed through native and facade worksheet APIs with width-boundary evidence. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_autofit_max_width#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::worksheet::Worksheet#struct::set_column_autofit_width#function` — The `u32` signature is pinned, but no public JavaScript adapter exists. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_column_autofit_width#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::worksheet::Worksheet#struct::set_column_range_width_pixels#function` — The `u32` signature is pinned, but no public JavaScript adapter exists. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_column_range_width_pixels#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::worksheet::Worksheet#struct::set_column_width_pixels#function` — The `u32` signature is pinned, but no public JavaScript adapter exists. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_column_width_pixels#function -->
- [ ] **MISSING / added** `rust_xlsxwriter::worksheet::Worksheet#struct::set_default_format#function` — Workbook-backed public worksheets cannot safely expose the independent-worksheet default-format contract. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_default_format#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::worksheet::Worksheet#struct::set_default_row_height_pixels#function` — The `u32` signature is pinned, but no public JavaScript adapter exists. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_default_row_height_pixels#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::worksheet::Worksheet#struct::set_row_height_pixels#function` — The `u32` signature is pinned, but no public JavaScript adapter exists. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_row_height_pixels#function -->

## Removed upstream rows

- **REMOVED** `rust_xlsxwriter::utility::serialize_chrono_naive_to_excel#function` — Absent from the `0.97.1` stable baseline. <!-- coverage-removed:rust_xlsxwriter::utility::serialize_chrono_naive_to_excel#function -->
- **REMOVED** `rust_xlsxwriter::utility::serialize_chrono_option_naive_to_excel#function` — Absent from the `0.97.1` stable baseline. <!-- coverage-removed:rust_xlsxwriter::utility::serialize_chrono_option_naive_to_excel#function -->
- **REMOVED** `rust_xlsxwriter::worksheet::Worksheet#struct::autofit_to_max_width#function` — Replaced upstream by `set_autofit_max_width()` plus `autofit()`. <!-- coverage-removed:rust_xlsxwriter::worksheet::Worksheet#struct::autofit_to_max_width#function -->
