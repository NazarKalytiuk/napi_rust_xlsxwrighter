# Migration record: `0.97.1` → `0.102.0`

[← Coverage dashboard](../rust-xlsxwriter-coverage-checklist.md)

> Historical diff: **17 added · 4 changed · 2 removed** stable rows. The changes are incorporated into the pinned `0.102.0` dashboard; this file is not counted separately.

## Помилки (`error`)

- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::NameError#variant` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::NameError#variant -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::NameError#variant::0#field` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::NameError#variant::0#field -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::NameError#variant::1#field` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::NameError#variant::1#field -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::NameReused#variant` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::NameReused#variant -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::NameReused#variant::0#field` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::NameReused#variant::0#field -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::TempFileError#variant` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::TempFileError#variant -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::TempFileError#variant::0#field` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::TempFileError#variant::0#field -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::error::XlsxError#enum::TempFileError#variant::1#field` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::error::XlsxError#enum::TempFileError#variant::1#field -->

## Допоміжні функції (`utility`)

- [ ] **MISSING / added** `rust_xlsxwriter::utility::check_name#function` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::utility::check_name#function -->

## Workbook (`workbook`)

- [x] **FULL / added** `rust_xlsxwriter::workbook::CalculationMode#enum` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::CalculationMode#enum -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::CalculationMode#enum::Automatic#variant` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::CalculationMode#enum::Automatic#variant -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::CalculationMode#enum::Manual#variant` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::CalculationMode#enum::Manual#variant -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::CalculationMode#enum::Partial#variant` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::CalculationMode#enum::Partial#variant -->
- [x] **FULL / signature-changed** `rust_xlsxwriter::workbook::Workbook#struct::add_worksheet_with_constant_memory#function` — `(&mut Self)->&mut rust_xlsxwriter::worksheet::Worksheet` → `(&mut Self)->core::result::Result<&mut rust_xlsxwriter::worksheet::Worksheet,rust_xlsxwriter::error::XlsxError>`; the audited status was carried over and needs re-verification. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::add_worksheet_with_constant_memory#function -->
- [x] **FULL / signature-changed** `rust_xlsxwriter::workbook::Workbook#struct::add_worksheet_with_low_memory#function` — `(&mut Self)->&mut rust_xlsxwriter::worksheet::Worksheet` → `(&mut Self)->core::result::Result<&mut rust_xlsxwriter::worksheet::Worksheet,rust_xlsxwriter::error::XlsxError>`; the audited status was carried over and needs re-verification. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::add_worksheet_with_low_memory#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::workbook::Workbook#struct::new_worksheet_with_constant_memory#function` — `(&mut Self)->rust_xlsxwriter::worksheet::Worksheet` → `(&mut Self)->core::result::Result<rust_xlsxwriter::worksheet::Worksheet,rust_xlsxwriter::error::XlsxError>`; the audited status was carried over and needs re-verification. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::new_worksheet_with_constant_memory#function -->
- [ ] **MISSING / signature-changed** `rust_xlsxwriter::workbook::Workbook#struct::new_worksheet_with_low_memory#function` — `(&mut Self)->rust_xlsxwriter::worksheet::Worksheet` → `(&mut Self)->core::result::Result<rust_xlsxwriter::worksheet::Worksheet,rust_xlsxwriter::error::XlsxError>`; the audited status was carried over and needs re-verification. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::new_worksheet_with_low_memory#function -->
- [x] **FULL / added** `rust_xlsxwriter::workbook::Workbook#struct::set_calculation_mode#function` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::set_calculation_mode#function -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::workbook::Workbook#struct::set_tab_ratio#function` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::set_tab_ratio#function -->
- [ ] **PARTIAL / added** `rust_xlsxwriter::workbook::Workbook#struct::set_zip_compression_level#function` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::workbook::Workbook#struct::set_zip_compression_level#function -->

## Worksheet (`worksheet`)

- [x] **FULL / added** `rust_xlsxwriter::worksheet::Worksheet#struct::set_hidden_zeroes#function` — Added upstream in `0.102.0`. <!-- coverage-drift:rust_xlsxwriter::worksheet::Worksheet#struct::set_hidden_zeroes#function -->

## Removed upstream rows

- **REMOVED** `rust_xlsxwriter::error::XlsxError#enum::TableNameReused#variant` — Absent from the `0.102.0` stable baseline. <!-- coverage-removed:rust_xlsxwriter::error::XlsxError#enum::TableNameReused#variant -->
- **REMOVED** `rust_xlsxwriter::error::XlsxError#enum::TableNameReused#variant::0#field` — Absent from the `0.102.0` stable baseline. <!-- coverage-removed:rust_xlsxwriter::error::XlsxError#enum::TableNameReused#variant::0#field -->
