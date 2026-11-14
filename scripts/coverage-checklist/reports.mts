// Renders the dashboard, the TypeScript/runtime backlog and the migration
// records that sit next to the per-module checklist files.

import {
  type Counts,
  type Group,
  type Status,
  DASHBOARD,
  countCapabilities,
  countRows,
  denominator,
  fullPercent,
  total,
} from './checklist.mts';

const BACK_LINK = `[← Coverage dashboard](../${DASHBOARD})`;

export interface DashboardFacts {
  version: string;
  features: string[];
  auditedOn: string;
  packageName: string;
  packageVersion: string;
  crateName: string;
  crateVersion: string;
  inventorySha256: string;
  extras: string[];
}

export interface GroupFile {
  fileName: string;
  group: Group;
}

export interface TypescriptMismatch {
  symbol: string;
  message: string;
  declaration: string;
}

export interface MigrationEntry {
  id: string;
  change: string;
  status: Status;
  explanation: string;
}

export interface Migration {
  fileName: string;
  from: string;
  to: string;
  groups: { module: string; title: string; entries: MigrationEntry[] }[];
  removed: { id: string; explanation: string }[];
}

const progressRow = (label: string, counts: Counts): string =>
  `| ${label} | ${total(counts)} | ${counts['N/A']} | ${denominator(counts)} | ${counts.FULL} | ${counts.PARTIAL} | ${counts.MISSING} | ${fullPercent(counts)} |`;

const box = (status: Status): string =>
  status === 'N/A' ? '' : status === 'FULL' ? '[x] ' : '[ ] ';

export function parseDashboard(markdown: string): Omit<DashboardFacts, 'inventorySha256'> {
  const match = (pattern: RegExp, what: string): RegExpExecArray => {
    const found = pattern.exec(markdown);
    if (!found) throw new Error(`${DASHBOARD}: cannot find ${what}`);
    return found;
  };
  const baseline = match(
    /`rust_xlsxwriter ([^`]+)` з resolved features `([^`]+)`/,
    'the pinned baseline',
  );
  const audited = match(/Аудит зафіксовано: \*\*([^*]+)\*\*/, 'the audit date');
  const pkg = match(
    /Пакет: \*\*(\S+) ([^*]+)\*\* \/ Cargo \*\*(\S+) ([^*]+)\*\*/,
    'the package versions',
  );
  const extrasSection = match(
    /## Local-only extras\n\n.*\n\n((?:- .*\n)+)/,
    'the local-only extras',
  );
  return {
    version: baseline[1]!,
    features: baseline[2]!.split(' + '),
    auditedOn: audited[1]!,
    packageName: pkg[1]!,
    packageVersion: pkg[2]!,
    crateName: pkg[3]!,
    crateVersion: pkg[4]!,
    extras: extrasSection[1]!.trimEnd().split('\n'),
  };
}

export function renderDashboard(
  facts: DashboardFacts,
  groups: GroupFile[],
  backlog: { fileName: string; mismatches: TypescriptMismatch[] },
  migrations: Migration[],
): string {
  const capabilities = groups.flatMap(({ group }) => group.capabilities);
  const rows = countRows(capabilities);
  const caps = countCapabilities(capabilities);
  const latest = migrations.at(-1);
  const latestCounts = latest ? migrationCounts(latest) : null;
  const groupRow = ({ fileName, group }: GroupFile): string => {
    const counts = countCapabilities(group.capabilities);
    const base = denominator(counts);
    const done = base === 0 ? 'N/A' : counts.FULL === base ? '[x]' : '[ ]';
    return `| ${done} | [${group.title}](coverage-checklist/${fileName}) | ${base} | ${counts.FULL} | ${counts.PARTIAL} | ${counts.MISSING} | ${counts['N/A']} | ${fullPercent(counts)} |`;
  };
  const migrationLink = (migration: Migration): string => {
    const counts = migrationCounts(migration);
    return `- [\`${migration.from}\` → \`${migration.to}\` migration record — ${counts.added} added, ${counts.changed} changed, ${counts.removed} removed](coverage-checklist/${migration.fileName})`;
  };

  return `${[
    '# `rust_xlsxwriter` coverage checklist',
    '',
    `> Внутрішній трекер повноти \`${facts.packageName}\`. Пінований baseline аудиту — \`rust_xlsxwriter ${facts.version}\` з resolved features \`${facts.features.join(' + ')}\`, згенерований з rustdoc JSON (\`npm run coverage:sync\`).`,
    '',
    `- Аудит зафіксовано: **${facts.auditedOn}**.`,
    `- Пакет: **${facts.packageName} ${facts.packageVersion}** / Cargo **${facts.crateName} ${facts.crateVersion}**.`,
    `- Inventory SHA-256 (sorted canonical row ids): \`${facts.inventorySha256}\`.`,
    latestCounts
      ? `- Міграція з попереднього baseline: **${latestCounts.added} added · ${latestCounts.changed} changed · ${latestCounts.removed} removed** stable rows.`
      : '- Міграцій з попереднього baseline немає.',
    '- Це item-weighted API checklist, а не usage-weighted оцінка популярності функцій.',
    '- Повний список розбитий на окремі файли за функціональними групами, щоб GitHub/editor не мусив рендерити один файл на 1500+ checkboxes.',
    '- Статуси, mapping, evidence і gap/note редагуються в group-файлах; inventory, `tests:` call sites, лічильники та цей dashboard генеруються. `npm run coverage:check` перевіряє, що вони актуальні.',
    '',
    '## Як вести checklist',
    '',
    '- `[x] FULL` — capability поведінково еквівалентна upstream у public JavaScript package.',
    '- `[ ] PARTIAL` — шлях існує, але звужені options, return/error semantics, builder access або runtime behavior.',
    '- `[ ] MISSING` — відповідного public N-API/facade шляху немає.',
    '- `N/A` — optional feature не скомпільована або Rust-only compile-time construct не має окремої runtime-операції.',
    '- Для enum variant з payload використовується один checkbox на capability; усі canonical symbol rows залишаються вкладеними.',
    '- При роботі над пунктом змінюйте status послідовно `MISSING` → `PARTIAL` → `FULL`; не ставте `[x]`, не змінивши label та evidence.',
    '- `tests:` виводяться з mapping: перші три call sites відповідних API у `tests/*.test.mjs` (для `JavaScript Error boundary` — перші три `toThrow` assertions).',
    "- Mapping з літеральним аргументом, напр. `Format.prototype.setAlign('left')` чи `Table.prototype.setStyle(0)`, рахує лише виклики, що передають цей літерал — прямо, через константу або через колонку таблиці `test.each`/`it.each`/`describe.each`. Так рядок enum variant посилається на тести саме цього значення.",
    '- `npm run coverage:check` падає, якщо `[x] FULL` рядок має API mapping, але жоден тест його не викликає.',
    '',
    '<a id="definition-of-done"></a>',
    '### Definition of done для `[x] FULL`',
    '',
    '1. Операція доступна через shipped N-API та/або public facade без прихованого native-only обходу.',
    '2. TypeScript declaration відповідає фактичному runtime.',
    '3. Збережено релевантний option space, return/error semantics та observable XLSX behavior.',
    `4. Є regression test; для складного mapping — direct \`rust_xlsxwriter ${facts.version}\` parity probe або перевірка XLSX/XML.`,
    '5. Оновлено local mapping/evidence у capability line; dashboard counts оновлює `npm run coverage:sync`.',
    '',
    '## Поточний прогрес',
    '',
    '| Режим підрахунку | Усього rows/capabilities | N/A | Denominator | Full | Partial | Missing | Full % |',
    '|---|---:|---:|---:|---:|---:|---:|---:|',
    progressRow('Symbol rows', rows),
    progressRow('Grouped input capabilities', caps),
    '',
    '## Dashboard і group milestones',
    '',
    '| Done | Група | Denominator | Full | Partial | Missing | N/A | Full % |',
    '|---|---|---:|---:|---:|---:|---:|---:|',
    ...groups.map(groupRow),
    '',
    '## Cross-cutting backlogs',
    '',
    `- [TypeScript/runtime integrity — ${backlog.mismatches.length} open mismatches](coverage-checklist/${backlog.fileName})`,
    ...[...migrations].toReversed().map(migrationLink),
    '',
    '## Local-only extras',
    '',
    'Ці exports не мають прямого `rust_xlsxwriter` counterpart і не входять у coverage denominator.',
    '',
    ...facts.extras,
    '',
    '## Audit invariants',
    '',
    `- Pinned canonical symbol rows: **${total(rows)}**.`,
    `- Pinned grouped capabilities: **${total(caps)}**.`,
    `- Pinned rustdoc inventory: **rust_xlsxwriter ${facts.version}**, SHA-256 \`${facts.inventorySha256}\`.`,
    ...(latest && latestCounts
      ? [
          `- Historical drift from \`${latest.from}\`: **${latestCounts.added} added, ${latestCounts.changed} changed, ${latestCounts.removed} removed**.`,
        ]
      : []),
    '- Однакові назви не вважаються доказом parity: статус визначається mapping, semantics і runtime/test evidence.',
    '- Не позначати `FULL` лише тому, що метод існує або test перевіряє тільки `not.toThrow()`.',
  ].join('\n')}\n`;
}

export function parseTypescriptBacklog(markdown: string): TypescriptMismatch[] {
  return [
    ...markdown.matchAll(
      /^- \[ \] \*\*signature-mismatch\*\* `([^`]+)` — TypeScript permits construction but runtime rejects it: (.*) Declarations: `([^`]+)`\. <!-- coverage-typescript:/gm,
    ),
  ].map(([, symbol, message, declaration]) => ({
    symbol: symbol!,
    message: message!,
    declaration: declaration!,
  }));
}

export function renderTypescriptBacklog(mismatches: TypescriptMismatch[]): string {
  return `${[
    '# TypeScript/runtime integrity backlog',
    '',
    BACK_LINK,
    '',
    `> **${mismatches.length} open declaration/runtime mismatches.** Це окремі blockers для чесного type-safe package contract.`,
    '',
    ...mismatches.map(
      ({ symbol, message, declaration }) =>
        `- [ ] **signature-mismatch** \`${symbol}\` — TypeScript permits construction but runtime rejects it: ${message} Declarations: \`${declaration}\`. <!-- coverage-typescript:${symbol} -->`,
    ),
  ].join('\n')}\n`;
}

export function parseMigration(markdown: string, fileName: string): Migration {
  const title = /^# Migration record: `([^`]+)` → `([^`]+)`$/m.exec(markdown);
  if (!title) throw new Error(`${fileName}: missing migration title`);
  const migration: Migration = {
    fileName,
    from: title[1]!,
    to: title[2]!,
    groups: [],
    removed: [],
  };
  for (const line of markdown.split('\n')) {
    const heading = /^## (.+) \(`(\w+)`\)$/.exec(line);
    if (heading) {
      migration.groups.push({ title: heading[1]!, module: heading[2]!, entries: [] });
      continue;
    }
    const entry =
      /^- (?:\[[ x]\] )?\*\*(FULL|PARTIAL|MISSING|N\/A) \/ ([\w-]+)\*\* `([^`]+)` — (.*) <!-- coverage-drift:/.exec(
        line,
      );
    if (entry) {
      const group = migration.groups.at(-1);
      if (!group) throw new Error(`${fileName}: entry outside a group: ${line}`);
      group.entries.push({
        status: entry[1] as Status,
        change: entry[2]!,
        id: entry[3]!,
        explanation: entry[4]!,
      });
      continue;
    }
    const removed = /^- \*\*REMOVED\*\* `([^`]+)` — (.*) <!-- coverage-removed:/.exec(line);
    if (removed) migration.removed.push({ id: removed[1]!, explanation: removed[2]! });
  }
  return migration;
}

export function migrationCounts(migration: Migration): {
  added: number;
  changed: number;
  removed: number;
} {
  const entries = migration.groups.flatMap((group) => group.entries);
  const added = entries.filter((entry) => entry.change === 'added').length;
  return { added, changed: entries.length - added, removed: migration.removed.length };
}

export function renderMigration(
  migration: Migration,
  statusOf: (id: string) => Status | undefined,
): string {
  const counts = migrationCounts(migration);
  const lines = [
    `# Migration record: \`${migration.from}\` → \`${migration.to}\``,
    '',
    BACK_LINK,
    '',
    `> Historical diff: **${counts.added} added · ${counts.changed} changed · ${counts.removed} removed** stable rows. The changes are incorporated into the pinned \`${migration.to}\` dashboard; this file is not counted separately.`,
  ];
  for (const group of migration.groups) {
    lines.push('', `## ${group.title} (\`${group.module}\`)`, '');
    for (const entry of group.entries) {
      // Rows keep their recorded change; the status follows the live checklist.
      const status = statusOf(entry.id) ?? entry.status;
      lines.push(
        `- ${box(status)}**${status} / ${entry.change}** \`${entry.id}\` — ${entry.explanation} <!-- coverage-drift:${entry.id} -->`,
      );
    }
  }
  if (migration.removed.length) {
    lines.push('', '## Removed upstream rows', '');
    for (const removed of migration.removed) {
      lines.push(
        `- **REMOVED** \`${removed.id}\` — ${removed.explanation} <!-- coverage-removed:${removed.id} -->`,
      );
    }
  }
  return `${lines.join('\n')}\n`;
}
