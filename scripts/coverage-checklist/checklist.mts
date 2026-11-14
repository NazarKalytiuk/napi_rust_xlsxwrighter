// Reads and writes the checklist markdown. Audited facts (status, mapping,
// evidence, gap/note) live in the markdown itself, where maintainers edit
// them; this module parses them back so the generator can refresh everything
// else (inventory, test references, counts) around them.

export type Status = 'FULL' | 'PARTIAL' | 'MISSING' | 'N/A';
export const STATUSES: readonly Status[] = ['FULL', 'PARTIAL', 'MISSING', 'N/A'];

export interface Remark {
  label: 'gap' | 'note';
  text: string;
}

export interface Capability {
  id: string;
  status: Status;
  /** Local APIs the capability maps to, e.g. `Format.prototype.setBold`. */
  mapping: string[];
  evidence: string | null;
  /** `tests/<file>:<line>` call sites. */
  tests: string[];
  remarks: Remark[];
  /** Canonical rows grouped under the capability (enum variants with payload). */
  nested: { id: string; status: Status }[] | null;
}

export interface Group {
  module: string;
  title: string;
  capabilities: Capability[];
}

export interface Counts {
  FULL: number;
  PARTIAL: number;
  MISSING: number;
  'N/A': number;
}

export const DASHBOARD = 'rust-xlsxwriter-coverage-checklist.md';
const BACK_LINK = `[← Coverage dashboard](../${DASHBOARD})`;
const STATUS_LEGEND =
  '- `[x] FULL`; `[ ] PARTIAL`; `[ ] MISSING`; plain `N/A` is outside the compiled/runtime denominator.';
const DONE_LINK = `- Before marking \`FULL\`, satisfy the [shared definition of done](../${DASHBOARD}#definition-of-done).`;

const CAPABILITY_LINE =
  /^- (?:\[[ x]\] )?\*\*(FULL|PARTIAL|MISSING|N\/A)\*\* `([^`]+)`(.*?) <!-- coverage-capability:(.+?) -->(?:<!-- coverage-row:(.+?) -->)?$/;
const NESTED_LINE =
  /^ {2}- \*\*(FULL|PARTIAL|MISSING|N\/A)\*\* `([^`]+)` <!-- coverage-row:(.+?) -->$/;
const FIELD_PREFIXES = ['→ ', 'evidence: ', 'tests: ', 'gap: ', 'note: '];

const unticked = (value: string): string[] =>
  [...value.matchAll(/`([^`]+)`/g)].map((match) => match[1]!);
const ticked = (values: string[]): string => values.map((value) => `\`${value}\``).join(', ');

export function parseGroup(markdown: string, fileName: string): Group {
  const lines = markdown.split('\n');
  const heading = /^# (.+) \(`(\w+)`\)$/.exec(lines[0] ?? '');
  if (!heading) throw new Error(`${fileName}: missing "# Title (\`module\`)" heading`);
  const capabilities: Capability[] = [];
  lines.forEach((line, index) => {
    const where = `${fileName}:${index + 1}`;
    const capability = CAPABILITY_LINE.exec(line);
    if (capability) {
      const [, status, id, fields, markerId, rowId] = capability;
      if (markerId !== id || (rowId !== undefined && rowId !== id)) {
        throw new Error(`${where}: markers do not match row id ${id}`);
      }
      capabilities.push({
        id: id!,
        status: status as Status,
        ...parseFields(fields!, where),
        nested: rowId === undefined ? [] : null,
      });
      return;
    }
    const nested = NESTED_LINE.exec(line);
    if (nested) {
      const [, status, id, markerId] = nested;
      const owner = capabilities.at(-1);
      if (markerId !== id || !owner?.nested) throw new Error(`${where}: stray nested row ${id}`);
      owner.nested.push({ id: id!, status: status as Status });
      return;
    }
    if (line.includes('coverage-capability:') || line.includes('coverage-row:')) {
      throw new Error(`${where}: unrecognized checklist row`);
    }
  });
  for (const capability of capabilities) {
    if (capability.nested?.length === 0) {
      throw new Error(`${fileName}: capability ${capability.id} has no rows`);
    }
  }
  return { module: heading[2]!, title: heading[1]!, capabilities };
}

function parseFields(text: string, where: string): Omit<Capability, 'id' | 'status' | 'nested'> {
  const segments: string[] = [];
  if (text) {
    if (!text.startsWith(' — ')) throw new Error(`${where}: fields must start with " — "`);
    // Free text may itself contain " — ", so only known prefixes start a field.
    for (const part of text.slice(3).split(' — ')) {
      if (FIELD_PREFIXES.some((prefix) => part.startsWith(prefix)) || segments.length === 0) {
        segments.push(part);
      } else {
        segments[segments.length - 1] += ` — ${part}`;
      }
    }
  }
  const result: Omit<Capability, 'id' | 'status' | 'nested'> = {
    mapping: [],
    evidence: null,
    tests: [],
    remarks: [],
  };
  for (const segment of segments) {
    if (segment.startsWith('→ ')) result.mapping = unticked(segment.slice(2));
    else if (segment.startsWith('evidence: ')) result.evidence = unticked(segment)[0] ?? null;
    else if (segment.startsWith('tests: ')) result.tests = unticked(segment);
    else if (segment.startsWith('gap: '))
      result.remarks.push({ label: 'gap', text: segment.slice(5) });
    else if (segment.startsWith('note: '))
      result.remarks.push({ label: 'note', text: segment.slice(6) });
    else throw new Error(`${where}: unknown field "${segment}"`);
  }
  return result;
}

export function renderCapability(capability: Capability): string[] {
  const fields: string[] = [];
  if (capability.mapping.length) fields.push(`→ ${ticked(capability.mapping)}`);
  if (capability.evidence) fields.push(`evidence: \`${capability.evidence}\``);
  if (capability.tests.length) fields.push(`tests: ${ticked(capability.tests)}`);
  for (const remark of capability.remarks) fields.push(`${remark.label}: ${remark.text}`);
  const box = capability.status === 'N/A' ? '' : capability.status === 'FULL' ? '[x] ' : '[ ] ';
  const marker = `<!-- coverage-capability:${capability.id} -->`;
  const head = `- ${box}**${capability.status}** \`${capability.id}\`${fields.map((field) => ` — ${field}`).join('')}`;
  if (!capability.nested) return [`${head} ${marker}<!-- coverage-row:${capability.id} -->`];
  return [
    `${head} ${marker}`,
    ...capability.nested.map(
      (row) => `  - **${row.status}** \`${row.id}\` <!-- coverage-row:${row.id} -->`,
    ),
  ];
}

export function countCapabilities(capabilities: Capability[]): Counts {
  const counts = emptyCounts();
  for (const capability of capabilities) counts[capability.status]++;
  return counts;
}

export function countRows(capabilities: Capability[]): Counts {
  const counts = emptyCounts();
  for (const capability of capabilities) {
    if (capability.nested) for (const row of capability.nested) counts[row.status]++;
    else counts[capability.status]++;
  }
  return counts;
}

export function emptyCounts(): Counts {
  return { FULL: 0, PARTIAL: 0, MISSING: 0, 'N/A': 0 };
}

export const total = (counts: Counts): number =>
  counts.FULL + counts.PARTIAL + counts.MISSING + counts['N/A'];
export const denominator = (counts: Counts): number => total(counts) - counts['N/A'];
export const fullPercent = (counts: Counts): string => {
  const base = denominator(counts);
  return `${base ? ((counts.FULL / base) * 100).toFixed(1) : '0.0'}%`;
};

const summary = (counts: Counts): string =>
  `**${counts.FULL} full · ${counts.PARTIAL} partial · ${counts.MISSING} missing · ${counts['N/A']} N/A**`;

/** The section a row is listed under: the top-level item after the module. */
export function sectionOf(id: string): string {
  const path = id.split('#')[0]!.split('::');
  return path[2]!;
}

export function renderGroup(group: Group, version: string): string {
  const capabilities = [...group.capabilities].toSorted((a, b) => compareIds(a.id, b.id));
  const caps = countCapabilities(capabilities);
  const rows = countRows(capabilities);
  const lines = [
    `# ${group.title} (\`${group.module}\`)`,
    '',
    BACK_LINK,
    '',
    `> Pinned baseline: \`rust_xlsxwriter ${version}\`; **${total(caps)} grouped capabilities**, **${total(rows)} canonical symbol rows**.`,
    '',
    `- Grouped: ${summary(caps)}.`,
    `- Symbol rows: ${summary(rows)}.`,
    STATUS_LEGEND,
    DONE_LINK,
  ];
  const sections = new Map<string, Capability[]>();
  for (const capability of capabilities) {
    const section = sectionOf(capability.id);
    sections.set(section, [...(sections.get(section) ?? []), capability]);
  }
  for (const section of [...sections.keys()].toSorted(compareIds)) {
    lines.push('', `## \`${section}\``, '');
    for (const capability of sections.get(section)!) lines.push(...renderCapability(capability));
  }
  return `${lines.join('\n')}\n`;
}

/** Ids are ordered by UTF-16 code units, which keeps `#trait` < `#trait-impl` < `#trait::`. */
export function compareIds(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}
