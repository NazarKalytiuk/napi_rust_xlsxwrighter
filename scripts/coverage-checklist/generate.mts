// Keeps the rust_xlsxwriter coverage checklist in sync with the code; see cli.mts.
//
// Sync documents the rust_xlsxwriter release pinned in Cargo.lock with rustdoc
// JSON, reconciles the canonical rows with the audited ones in the markdown,
// records a migration when the pinned release changes, and rewrites the
// generated parts. Check needs no nightly toolchain: it re-derives test call
// sites and counts from the markdown and fails if any file would change.

import crypto from 'node:crypto';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { parseSync } from 'oxc-parser';

import {
  type Capability,
  type Group,
  type Status,
  DASHBOARD,
  compareIds,
  parseGroup,
  renderGroup,
} from './checklist.mts';
import { type InventoryRow, buildInventory } from './inventory.mts';
import {
  type DashboardFacts,
  type GroupFile,
  type Migration,
  type MigrationEntry,
  type TypescriptMismatch,
  parseDashboard,
  parseMigration,
  parseTypescriptBacklog,
  renderDashboard,
  renderMigration,
  renderTypescriptBacklog,
} from './reports.mts';
import { declaredFeatures, projectBaseline, rustdocJson } from './rustdoc.mts';
import {
  ERROR_BOUNDARY,
  type CallSite,
  collectCallSites,
  declaredApis,
  readApiShapes,
} from './test-refs.mts';

const ROOT = path.resolve(import.meta.dirname, '..', '..');
const CHECKLIST_DIR = path.join(ROOT, 'coverage-checklist');
const WORK_DIR = path.join(ROOT, 'target', 'coverage-checklist');
const TESTS_PER_ROW = 3;
const COMPILE_TIME_GAP =
  'Compile-time Rust trait/type integration has no independent JavaScript runtime operation; concrete JS inputs/outputs are classified on their observable API rows.';

export interface Checklist {
  facts: Omit<DashboardFacts, 'inventorySha256'>;
  groups: Map<string, Group>;
  backlog: { fileName: string; mismatches: TypescriptMismatch[] };
  migrations: Migration[];
}

export function readChecklist(): Checklist {
  const facts = parseDashboard(fs.readFileSync(path.join(ROOT, DASHBOARD), 'utf8'));
  const groups = new Map<string, Group>();
  const migrations: Migration[] = [];
  let backlog: Checklist['backlog'] | null = null;
  for (const fileName of fs.readdirSync(CHECKLIST_DIR).toSorted()) {
    if (!fileName.endsWith('.md')) continue;
    const markdown = fs.readFileSync(path.join(CHECKLIST_DIR, fileName), 'utf8');
    if (markdown.startsWith('# Migration record:'))
      migrations.push(parseMigration(markdown, fileName));
    else if (markdown.startsWith('# TypeScript/runtime integrity backlog')) {
      backlog = { fileName, mismatches: parseTypescriptBacklog(markdown) };
    } else {
      const group = parseGroup(markdown, fileName);
      groups.set(group.module, group);
    }
  }
  if (!backlog) throw new Error('coverage-checklist has no TypeScript/runtime backlog file');
  return { facts, groups, backlog, migrations };
}

function rowStatuses(groups: Iterable<Group>): Map<string, Status> {
  const statuses = new Map<string, Status>();
  for (const group of groups) {
    for (const capability of group.capabilities) {
      if (capability.nested) for (const row of capability.nested) statuses.set(row.id, row.status);
      else statuses.set(capability.id, capability.status);
    }
  }
  return statuses;
}

function difference<T>(left: Iterable<T>, right: Set<T> | Map<T, unknown>): T[] {
  return [...left].filter((value) => !right.has(value));
}

/** Documents the pinned release and reconciles the checklist rows with it. */
export function sync(checklist: Checklist): void {
  const baseline = projectBaseline(ROOT);
  const features = declaredFeatures(WORK_DIR, baseline.version);
  const document = (version: string, enabled: string[]): Map<string, InventoryRow> =>
    buildInventory(
      rustdocJson(WORK_DIR, { version, features: enabled.filter((f) => f !== 'default') }),
    );
  const everything = document(baseline.version, features);
  const compiled = document(baseline.version, baseline.features);

  // Rows that only exist with optional features are N/A; find which features they need.
  const gated = difference(everything.keys(), compiled);
  const requiredFeatures = new Map<string, string[]>();
  for (const feature of features.filter((f) => !baseline.features.includes(f))) {
    if (!gated.length) break;
    const without = document(
      baseline.version,
      features.filter((f) => f !== feature),
    );
    for (const id of gated) {
      if (!without.has(id))
        requiredFeatures.set(id, [...(requiredFeatures.get(id) ?? []), feature]);
    }
  }

  const statuses = rowStatuses(checklist.groups.values());
  const pinned = checklist.facts.version;
  if (pinned !== baseline.version) {
    const previous = document(pinned, declaredFeatures(WORK_DIR, pinned));
    assertSameRows(statuses, previous, `the checklist pinned to ${pinned}`);
    const previousCompiled = document(pinned, checklist.facts.features);
    checklist.migrations.push(
      migration(checklist, previous, everything, previousCompiled, compiled, baseline.version),
    );
    checklist.facts.version = baseline.version;
    checklist.facts.auditedOn = new Date().toISOString().slice(0, 10);
  } else {
    assertSameRows(statuses, everything, `rust_xlsxwriter ${pinned}`);
  }
  checklist.facts.features = renderFeatures(baseline.features);

  const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  const cargoToml = fs.readFileSync(path.join(ROOT, 'Cargo.toml'), 'utf8');
  checklist.facts.packageName = packageJson.name;
  checklist.facts.packageVersion = packageJson.version;
  checklist.facts.crateName = /^name = "([^"]+)"/m.exec(cargoToml)![1]!;
  checklist.facts.crateVersion = /^version = "([^"]+)"/m.exec(cargoToml)![1]!;

  reconcile(checklist, everything, compiled, requiredFeatures, baseline.version);
}

function assertSameRows(
  statuses: Map<string, Status>,
  inventory: Map<string, InventoryRow>,
  what: string,
): void {
  const missing = difference(statuses.keys(), inventory);
  const extra = difference(inventory.keys(), statuses);
  if (missing.length || extra.length) {
    throw new Error(
      `The checklist rows do not match ${what}:\n` +
        [
          ...missing.map((id) => `  only in the checklist: ${id}`),
          ...extra.map((id) => `  only in rustdoc: ${id}`),
        ].join('\n'),
    );
  }
}

function renderFeatures(resolved: string[]): string[] {
  return ['default', ...resolved.filter((feature) => feature !== 'default').toSorted()];
}

function migration(
  checklist: Checklist,
  previous: Map<string, InventoryRow>,
  current: Map<string, InventoryRow>,
  previousCompiled: Map<string, InventoryRow>,
  compiled: Map<string, InventoryRow>,
  version: string,
): Migration {
  const entries = new Map<string, MigrationEntry[]>();
  const add = (row: InventoryRow, change: string, explanation: string): void => {
    entries.set(row.module, [
      ...(entries.get(row.module) ?? []),
      { id: row.id, change, status: 'MISSING', explanation },
    ]);
  };
  for (const row of current.values()) {
    const before = previous.get(row.id);
    if (!before) {
      add(row, 'added', `Added upstream in \`${version}\`.`);
    } else if (before.signature !== row.signature) {
      add(
        row,
        row.id.includes('#variant') ? 'variant-changed' : 'signature-changed',
        `\`${before.signature || '()'}\` → \`${row.signature || '()'}\`; the audited status was carried over and needs re-verification.`,
      );
    } else if (previousCompiled.has(row.id) !== compiled.has(row.id)) {
      add(
        row,
        'feature-gate-changed',
        compiled.has(row.id)
          ? 'Now compiled into the package baseline; the audited status needs re-verification.'
          : 'No longer compiled into the package baseline.',
      );
    }
  }
  const lastNumber = Math.max(
    ...[checklist.backlog.fileName, ...checklist.migrations.map((m) => m.fileName)].map((name) =>
      Number(name.slice(0, 2)),
    ),
  );
  const titleOf = (module: string): string => checklist.groups.get(module)?.title ?? module;
  return {
    fileName: `${String(lastNumber + 1).padStart(2, '0')}-drift-${checklist.facts.version}-${version}.md`,
    from: checklist.facts.version,
    to: version,
    groups: [...entries.keys()].toSorted().map((module) => ({
      module,
      title: titleOf(module),
      entries: entries.get(module)!.toSorted((a, b) => compareIds(a.id, b.id)),
    })),
    removed: difference(previous.keys(), current)
      .toSorted(compareIds)
      .map((id) => ({ id, explanation: `Absent from the \`${version}\` stable baseline.` })),
  };
}

/** Rebuilds every group from the inventory, keeping the audited data of known rows. */
function reconcile(
  checklist: Checklist,
  everything: Map<string, InventoryRow>,
  compiled: Map<string, InventoryRow>,
  requiredFeatures: Map<string, string[]>,
  version: string,
): void {
  const known = new Map<string, Capability>();
  for (const group of checklist.groups.values())
    for (const c of group.capabilities) known.set(c.id, c);
  const nestedStatus = rowStatuses(checklist.groups.values());

  const byCapability = new Map<string, InventoryRow[]>();
  for (const row of everything.values()) {
    byCapability.set(row.capability, [...(byCapability.get(row.capability) ?? []), row]);
  }
  const groups = new Map<string, Group>();
  for (const [id, rows] of byCapability) {
    const { module } = rows[0]!;
    const capability: Capability =
      known.get(id) ?? newCapability(id, compiled, requiredFeatures, version);
    capability.nested =
      rows.length > 1
        ? rows
            .map((row) => ({ id: row.id, status: nestedStatus.get(row.id) ?? capability.status }))
            .toSorted((a, b) => compareIds(a.id, b.id))
        : null;
    const group = groups.get(module) ?? {
      module,
      title: checklist.groups.get(module)?.title ?? module,
      capabilities: [],
    };
    group.capabilities.push(capability);
    groups.set(module, group);
  }
  checklist.groups = groups;
}

function newCapability(
  id: string,
  compiled: Map<string, InventoryRow>,
  requiredFeatures: Map<string, string[]>,
  version: string,
): Capability {
  const base = { id, mapping: [], evidence: null, tests: [], nested: null };
  if (!compiled.has(id)) {
    const required = requiredFeatures.get(id) ?? [];
    return {
      ...base,
      status: 'N/A',
      remarks: [
        {
          label: 'gap',
          text: `Out of the compiled dependency baseline; requires optional feature(s): ${required.join(', ') || 'a combination of optional features'}.`,
        },
      ],
    };
  }
  if (/#(?:trait|type_alias)(?:$|::|-impl-for:)/.test(id)) {
    return { ...base, status: 'N/A', remarks: [{ label: 'gap', text: COMPILE_TIME_GAP }] };
  }
  return {
    ...base,
    status: 'MISSING',
    remarks: [{ label: 'note', text: `Added upstream in \`${version}\`; not audited yet.` }],
  };
}

/** Fails on mappings to APIs that wrapper.ts and index.d.ts no longer declare. */
export function validateMappings(groups: Iterable<Group>): void {
  const apis = new Set([
    ...declaredApis(fs.readFileSync(path.join(ROOT, 'wrapper.ts'), 'utf8'), 'wrapper.ts'),
    ...declaredApis(fs.readFileSync(path.join(ROOT, 'index.d.ts'), 'utf8'), 'index.d.ts'),
  ]);
  const stale: string[] = [];
  for (const group of groups) {
    for (const capability of group.capabilities) {
      for (const target of capability.mapping) {
        if (
          /^\w+(?:\.(?:prototype|static)\.\w+|::constructor)$/.test(target) &&
          !apis.has(target)
        ) {
          stale.push(`  ${capability.id} → ${target}`);
        }
      }
    }
  }
  if (stale.length) {
    throw new Error(`Checklist mappings reference APIs that do not exist:\n${stale.join('\n')}`);
  }
}

/** Points every row at the first call sites of the APIs it maps to. */
export function refreshTests(groups: Iterable<Group>): void {
  const shapes = readApiShapes(
    fs.readFileSync(path.join(ROOT, 'wrapper.ts'), 'utf8'),
    'wrapper.ts',
  );
  const testFiles = fs
    .readdirSync(path.join(ROOT, 'tests'))
    .filter((file) => file.endsWith('.test.mjs'))
    .map((file) => `tests/${file}`);
  const sites = collectCallSites(ROOT, testFiles, shapes);
  for (const group of groups) {
    for (const capability of group.capabilities) {
      const keys = capability.mapping.flatMap((target) =>
        /^\w+\.(?:prototype|static)\.\w+$/.test(target)
          ? [target]
          : target === 'JavaScript Error boundary'
            ? [ERROR_BOUNDARY]
            : [],
      );
      const unique = new Map<string, CallSite>();
      for (const site of keys.flatMap((key) => sites.get(key) ?? []))
        unique.set(`${site.file}:${site.line}`, site);
      capability.tests = [...unique.entries()]
        .toSorted(([, a], [, b]) => compareIds(a.file, b.file) || a.line - b.line)
        .slice(0, TESTS_PER_ROW)
        .map(([key]) => key);
    }
  }
}

/** Native classes whose declarations allow `new` although the runtime refuses it. */
export function probeTypescriptRuntime(): TypescriptMismatch[] | null {
  let native: Record<string, new () => unknown>;
  try {
    native = createRequire(path.join(ROOT, 'index.js'))('./index.js');
  } catch {
    return null;
  }
  const source = fs.readFileSync(path.join(ROOT, 'index.d.ts'), 'utf8');
  const { program } = parseSync('index.d.ts', source);
  const mismatches: TypescriptMismatch[] = [];
  for (const statement of program.body as any[]) {
    const declaration = statement.type === 'ExportNamedDeclaration' ? statement.declaration : null;
    if (declaration?.type !== 'ClassDeclaration') continue;
    const declaresConstructor = declaration.body.body.some(
      (member: any) => member.type === 'MethodDefinition' && member.kind === 'constructor',
    );
    if (declaresConstructor) continue;
    const name: string = declaration.id.name;
    try {
      Reflect.construct(native[name]!, []);
    } catch (error) {
      const line = source.slice(0, declaration.start).split('\n').length;
      mismatches.push({
        symbol: `${name}::constructor`,
        message: (error as Error).message,
        declaration: `index.d.ts:${line}`,
      });
    }
  }
  return mismatches.toSorted((a, b) => compareIds(a.symbol, b.symbol));
}

/** Renders every generated file, keyed by its path relative to the project root. */
export function renderOutputs(checklist: Checklist): Map<string, string> {
  const modules = [...checklist.groups.keys()].toSorted();
  const groupFiles: GroupFile[] = modules.map((module, index) => ({
    fileName: `${String(index + 1).padStart(2, '0')}-${module.replaceAll('_', '-')}.md`,
    group: checklist.groups.get(module)!,
  }));
  checklist.backlog.fileName = `${String(modules.length + 1).padStart(2, '0')}-typescript-runtime.md`;
  const rowIds = [...rowStatuses(checklist.groups.values()).keys()].toSorted(compareIds);
  const facts: DashboardFacts = {
    ...checklist.facts,
    inventorySha256: crypto.createHash('sha256').update(rowIds.join('\n')).digest('hex'),
  };
  const statuses = rowStatuses(checklist.groups.values());
  return new Map<string, string>([
    [DASHBOARD, renderDashboard(facts, groupFiles, checklist.backlog, checklist.migrations)],
    ...groupFiles.map(({ fileName, group }): [string, string] => [
      `coverage-checklist/${fileName}`,
      renderGroup(group, facts.version),
    ]),
    [
      `coverage-checklist/${checklist.backlog.fileName}`,
      renderTypescriptBacklog(checklist.backlog.mismatches),
    ],
    ...checklist.migrations.map((m): [string, string] => [
      `coverage-checklist/${m.fileName}`,
      renderMigration(m, (id) => statuses.get(id)),
    ]),
  ]);
}

export interface RunOptions {
  /** Compare instead of writing; fails when a file would change. */
  check: boolean;
  /** Re-document the release with rustdoc even if it is already pinned. */
  rustdoc: boolean;
}

export function run({ check, rustdoc }: RunOptions): void {
  const checklist = readChecklist();
  const baseline = projectBaseline(ROOT);
  const features = renderFeatures(baseline.features);
  // The inventory only depends on the release and its features, so rustdoc is
  // needed only when Cargo.lock moved away from the pinned baseline.
  const pinned =
    baseline.version === checklist.facts.version &&
    features.join() === checklist.facts.features.join();
  if (check && !pinned) {
    throw new Error(
      `Cargo.lock resolves rust_xlsxwriter ${baseline.version} (${features.join(' + ')}), the checklist pins ${checklist.facts.version} (${checklist.facts.features.join(' + ')}); run \`npm run coverage:sync\`.`,
    );
  }
  if (!check && (!pinned || rustdoc)) sync(checklist);
  validateMappings(checklist.groups.values());
  refreshTests(checklist.groups.values());

  const probed = probeTypescriptRuntime();
  if (probed) checklist.backlog.mismatches = probed;
  else if (!check)
    throw new Error('Build the native addon first (`npm run build`): index.js cannot be loaded.');
  else console.warn('Skipping the TypeScript/runtime probe: the native addon is not built.');

  const outputs = renderOutputs(checklist);
  const stale = fs
    .readdirSync(CHECKLIST_DIR)
    .filter((fileName) => fileName.endsWith('.md'))
    .map((fileName) => `coverage-checklist/${fileName}`)
    .filter((file) => !outputs.has(file));

  const changed = [...outputs].filter(([file, text]) => {
    const target = path.join(ROOT, file);
    return !fs.existsSync(target) || fs.readFileSync(target, 'utf8') !== text;
  });
  if (check) {
    if (changed.length || stale.length) {
      console.error(
        [
          'The coverage checklist is out of date; run `npm run coverage:sync`.',
          ...changed.map(([file]) => `  would update ${file}`),
          ...stale.map((file) => `  would remove ${file}`),
        ].join('\n'),
      );
      process.exitCode = 1;
    }
    return;
  }
  for (const [file, text] of changed) fs.writeFileSync(path.join(ROOT, file), text);
  for (const file of stale) fs.rmSync(path.join(ROOT, file));
  console.log(
    `Coverage checklist synced to rust_xlsxwriter ${checklist.facts.version}: ${changed.length} file(s) updated, ${stale.length} removed.`,
  );
}
