# `rust_xlsxwriter` coverage checklist

> Внутрішній трекер повноти `@nazarkalytiuk/rusc-xlsx`. Пінований baseline аудиту — `rust_xlsxwriter 0.97.1` з resolved features `default + constant_memory`, згенерований з rustdoc JSON.

- Аудит зафіксовано: **2026-08-08**.
- Пакет: **@nazarkalytiuk/rusc-xlsx 2.0.0** / Cargo **rusc-xlsx 2.0.0**.
- Pinned rustdoc SHA-256: `51152353cf9c3b2a6d92c3519b95da5126803c3656adc07ed2418213018d7d42`.
- Міграція з попереднього baseline: **17 added · 11 changed · 3 removed** stable rows.
- Це item-weighted API checklist, а не usage-weighted оцінка популярності функцій.
- Повний список розбитий на окремі файли за функціональними групами, щоб GitHub/editor не мусив рендерити один файл на 1500+ checkboxes.

## Як вести checklist

- `[x] FULL` — capability поведінково еквівалентна upstream у public JavaScript package.
- `[ ] PARTIAL` — шлях існує, але звужені options, return/error semantics, builder access або runtime behavior.
- `[ ] MISSING` — відповідного public N-API/facade шляху немає.
- `N/A` — optional feature не скомпільована або Rust-only compile-time construct не має окремої runtime-операції.
- Для enum variant з payload використовується один checkbox на capability; усі canonical symbol rows залишаються вкладеними.
- При роботі над пунктом змінюйте status послідовно `MISSING` → `PARTIAL` → `FULL`; не ставте `[x]`, не змінивши label та evidence.

<a id="definition-of-done"></a>
### Definition of done для `[x] FULL`

1. Операція доступна через shipped N-API та/або public facade без прихованого native-only обходу.
2. TypeScript declaration відповідає фактичному runtime.
3. Збережено релевантний option space, return/error semantics та observable XLSX behavior.
4. Є regression test; для складного mapping — direct `rust_xlsxwriter 0.97.1` parity probe або перевірка XLSX/XML.
5. Оновлено local mapping/evidence у capability line та dashboard counts у цьому файлі.

## Поточний прогрес

| Режим підрахунку | Усього rows/capabilities | N/A | Denominator | Full | Partial | Missing | Full % |
|---|---:|---:|---:|---:|---:|---:|---:|
| Symbol rows | 1658 | 202 | 1456 | 400 | 281 | 775 | 27.5% |
| Grouped input capabilities | 1587 | 200 | 1387 | 377 | 252 | 758 | 27.2% |

## Dashboard і group milestones

| Done | Група | Denominator | Full | Partial | Missing | N/A | Full % |
|---|---|---:|---:|---:|---:|---:|---:|
| [x] | [Кнопки](coverage-checklist/01-button.md) | 8 | 8 | 0 | 0 | 0 | 100.0% |
| [ ] | [Діаграми](coverage-checklist/02-chart.md) | 404 | 30 | 16 | 358 | 15 | 7.4% |
| [ ] | [Кольори](coverage-checklist/03-color.md) | 21 | 17 | 3 | 1 | 0 | 81.0% |
| [ ] | [Comment API](coverage-checklist/04-comment.md) | 1 | 0 | 0 | 1 | 0 | 0.0% |
| [ ] | [Умовне форматування](coverage-checklist/05-conditional-format.md) | 181 | 52 | 53 | 76 | 51 | 28.7% |
| [ ] | [Валідація даних](coverage-checklist/06-data-validation.md) | 39 | 26 | 3 | 10 | 17 | 66.7% |
| [ ] | [Дата й час](coverage-checklist/07-datetime.md) | 10 | 0 | 5 | 5 | 16 | 0.0% |
| [ ] | [Помилки](coverage-checklist/08-error.md) | 34 | 0 | 34 | 0 | 2 | 0.0% |
| [ ] | [Фільтри](coverage-checklist/09-filter.md) | 22 | 0 | 0 | 22 | 5 | 0.0% |
| [ ] | [Форматування](coverage-checklist/10-format.md) | 121 | 56 | 49 | 16 | 0 | 46.3% |
| [x] | [Формули](coverage-checklist/11-formula.md) | 3 | 3 | 0 | 0 | 0 | 100.0% |
| [ ] | [Зображення](coverage-checklist/12-image.md) | 25 | 16 | 1 | 8 | 0 | 64.0% |
| [ ] | [Нотатки / коментарі](coverage-checklist/13-note.md) | 13 | 7 | 1 | 5 | 0 | 53.8% |
| [ ] | [Властивості документа](coverage-checklist/14-properties.md) | 15 | 12 | 0 | 3 | 22 | 80.0% |
| [x] | [Захист аркуша](coverage-checklist/15-protection.md) | 18 | 18 | 0 | 0 | 0 | 100.0% |
| N/A | [Серіалізація](coverage-checklist/16-serializer.md) | 0 | 0 | 0 | 0 | 20 | 0.0% |
| [ ] | [Фігури](coverage-checklist/17-shape.md) | 142 | 7 | 0 | 135 | 7 | 4.9% |
| [ ] | [Спарклайни](coverage-checklist/18-sparkline.md) | 33 | 23 | 8 | 2 | 0 | 69.7% |
| [ ] | [Таблиці](coverage-checklist/19-table.md) | 95 | 83 | 9 | 3 | 0 | 87.4% |
| [ ] | [Гіперпосилання](coverage-checklist/20-url.md) | 4 | 0 | 3 | 1 | 0 | 0.0% |
| [ ] | [Допоміжні функції](coverage-checklist/21-utility.md) | 11 | 0 | 0 | 11 | 2 | 0.0% |
| [ ] | [Workbook](coverage-checklist/22-workbook.md) | 27 | 14 | 2 | 11 | 0 | 51.9% |
| [ ] | [Worksheet](coverage-checklist/23-worksheet.md) | 160 | 5 | 65 | 90 | 43 | 3.1% |

## Cross-cutting backlogs

- [TypeScript/runtime integrity — 8 open mismatches](coverage-checklist/24-typescript-runtime.md)
- [`0.90.2` → `0.97.1` migration record — 17 added, 11 changed, 3 removed](coverage-checklist/25-latest-0.97-drift.md)

## Local-only extras

Ці exports не мають прямого `rust_xlsxwriter` counterpart і не входять у coverage denominator.

- [x] `rustMemoryStats` — Package-specific Rust allocator diagnostics; no rust_xlsxwriter counterpart.
- [x] `rustMemoryResetPeak` — Package-specific peak allocator reset; no rust_xlsxwriter counterpart.

## Audit invariants

- Pinned canonical symbol rows: **1658**.
- Pinned grouped capabilities: **1587**.
- Pinned rustdoc artifact: **rust_xlsxwriter 0.97.1**, SHA-256 `51152353cf9c3b2a6d92c3519b95da5126803c3656adc07ed2418213018d7d42`.
- Historical drift from `0.90.2`: **17 added, 11 changed, 3 removed**.
- Однакові назви не вважаються доказом parity: статус визначається mapping, semantics і runtime/test evidence.
- Не позначати `FULL` лише тому, що метод існує або test перевіряє тільки `not.toThrow()`.
