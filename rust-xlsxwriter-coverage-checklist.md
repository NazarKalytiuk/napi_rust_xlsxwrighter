# `rust_xlsxwriter` coverage checklist

> Внутрішній трекер повноти `@nazarkalytiuk/rusc-xlsx`. Пінований baseline аудиту — `rust_xlsxwriter 0.102.0` з resolved features `default + constant_memory + zmij`, згенерований з rustdoc JSON (`npm run coverage:sync`).

- Аудит зафіксовано: **2026-11-14**.
- Пакет: **@nazarkalytiuk/rusc-xlsx 2.1.0** / Cargo **rusc-xlsx 2.1.0**.
- Inventory SHA-256 (sorted canonical row ids): `83018e65c3d539594dbcd6a1cb4c3257852b81def10a47f71a8592ffd2e1d512`.
- Міграція з попереднього baseline: **17 added · 4 changed · 2 removed** stable rows.
- Це item-weighted API checklist, а не usage-weighted оцінка популярності функцій.
- Повний список розбитий на окремі файли за функціональними групами, щоб GitHub/editor не мусив рендерити один файл на 1500+ checkboxes.
- Статуси, mapping, evidence і gap/note редагуються в group-файлах; inventory, `tests:` call sites, лічильники та цей dashboard генеруються. `npm run coverage:check` перевіряє, що вони актуальні.

## Як вести checklist

- `[x] FULL` — capability поведінково еквівалентна upstream у public JavaScript package.
- `[ ] PARTIAL` — шлях існує, але звужені options, return/error semantics, builder access або runtime behavior.
- `[ ] MISSING` — відповідного public N-API/facade шляху немає.
- `N/A` — optional feature не скомпільована або Rust-only compile-time construct не має окремої runtime-операції.
- Для enum variant з payload використовується один checkbox на capability; усі canonical symbol rows залишаються вкладеними.
- При роботі над пунктом змінюйте status послідовно `MISSING` → `PARTIAL` → `FULL`; не ставте `[x]`, не змінивши label та evidence.
- `tests:` виводяться з mapping: перші три call sites відповідних API у `tests/*.test.mjs` (для `JavaScript Error boundary` — перші три `toThrow` assertions).
- Mapping з літеральним аргументом, напр. `Format.prototype.setAlign('left')` чи `Table.prototype.setStyle(0)`, рахує лише виклики, що передають цей літерал — прямо, через константу або через колонку таблиці `test.each`/`it.each`/`describe.each`. Так рядок enum variant посилається на тести саме цього значення.
- `npm run coverage:check` падає, якщо `[x] FULL` рядок має API mapping, але жоден тест його не викликає.

<a id="definition-of-done"></a>
### Definition of done для `[x] FULL`

1. Операція доступна через shipped N-API та/або public facade без прихованого native-only обходу.
2. TypeScript declaration відповідає фактичному runtime.
3. Збережено релевантний option space, return/error semantics та observable XLSX behavior.
4. Є regression test; для складного mapping — direct `rust_xlsxwriter 0.102.0` parity probe або перевірка XLSX/XML.
5. Оновлено local mapping/evidence у capability line; dashboard counts оновлює `npm run coverage:sync`.

## Поточний прогрес

| Режим підрахунку | Усього rows/capabilities | N/A | Denominator | Full | Partial | Missing | Full % |
|---|---:|---:|---:|---:|---:|---:|---:|
| Symbol rows | 1673 | 202 | 1471 | 423 | 270 | 778 | 28.8% |
| Grouped input capabilities | 1598 | 200 | 1398 | 400 | 238 | 760 | 28.6% |

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
| [ ] | [Помилки](coverage-checklist/08-error.md) | 36 | 0 | 36 | 0 | 2 | 0.0% |
| [ ] | [Фільтри](coverage-checklist/09-filter.md) | 22 | 0 | 0 | 22 | 5 | 0.0% |
| [ ] | [Форматування](coverage-checklist/10-format.md) | 121 | 70 | 35 | 16 | 0 | 57.9% |
| [x] | [Формули](coverage-checklist/11-formula.md) | 3 | 3 | 0 | 0 | 0 | 100.0% |
| [ ] | [Зображення](coverage-checklist/12-image.md) | 25 | 16 | 1 | 8 | 0 | 64.0% |
| [ ] | [Нотатки / коментарі](coverage-checklist/13-note.md) | 13 | 7 | 1 | 5 | 0 | 53.8% |
| [ ] | [Властивості документа](coverage-checklist/14-properties.md) | 15 | 12 | 0 | 3 | 22 | 80.0% |
| [x] | [Захист аркуша](coverage-checklist/15-protection.md) | 18 | 18 | 0 | 0 | 0 | 100.0% |
| N/A | [Серіалізація](coverage-checklist/16-serializer.md) | 0 | 0 | 0 | 0 | 20 | 0.0% |
| [ ] | [Фігури](coverage-checklist/17-shape.md) | 142 | 7 | 0 | 135 | 7 | 4.9% |
| [ ] | [Спарклайни](coverage-checklist/18-sparkline.md) | 33 | 23 | 8 | 2 | 0 | 69.7% |
| [ ] | [Таблиці](coverage-checklist/19-table.md) | 95 | 86 | 5 | 4 | 0 | 90.5% |
| [ ] | [Гіперпосилання](coverage-checklist/20-url.md) | 4 | 0 | 3 | 1 | 0 | 0.0% |
| [ ] | [Допоміжні функції](coverage-checklist/21-utility.md) | 12 | 0 | 0 | 12 | 2 | 0.0% |
| [ ] | [Workbook](coverage-checklist/22-workbook.md) | 34 | 19 | 4 | 11 | 0 | 55.9% |
| [ ] | [Worksheet](coverage-checklist/23-worksheet.md) | 161 | 6 | 65 | 90 | 43 | 3.7% |

## Cross-cutting backlogs

- [TypeScript/runtime integrity — 8 open mismatches](coverage-checklist/24-typescript-runtime.md)
- [`0.97.1` → `0.102.0` migration record — 17 added, 4 changed, 2 removed](coverage-checklist/26-drift-0.97.1-0.102.0.md)
- [`0.90.2` → `0.97.1` migration record — 17 added, 11 changed, 3 removed](coverage-checklist/25-drift-0.90.2-0.97.1.md)

## Local-only extras

Ці exports не мають прямого `rust_xlsxwriter` counterpart і не входять у coverage denominator.

- [x] `rustMemoryStats` — Package-specific Rust allocator diagnostics; no rust_xlsxwriter counterpart.
- [x] `rustMemoryResetPeak` — Package-specific peak allocator reset; no rust_xlsxwriter counterpart.

## Audit invariants

- Pinned canonical symbol rows: **1673**.
- Pinned grouped capabilities: **1598**.
- Pinned rustdoc inventory: **rust_xlsxwriter 0.102.0**, SHA-256 `83018e65c3d539594dbcd6a1cb4c3257852b81def10a47f71a8592ffd2e1d512`.
- Historical drift from `0.97.1`: **17 added, 4 changed, 2 removed**.
- Однакові назви не вважаються доказом parity: статус визначається mapping, semantics і runtime/test evidence.
- Не позначати `FULL` лише тому, що метод існує або test перевіряє тільки `not.toThrow()`.
