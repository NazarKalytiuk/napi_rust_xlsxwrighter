import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, describe, expect, test } from 'vitest';

import {
  API_KEY,
  ERROR_BOUNDARY,
  collectCallSites,
  declaredApis,
  literalArgument,
  readApiShapes,
} from './test-refs.mts';

const WRAPPER = `
export class Workbook {
  private constructor() {}
  static create(): Workbook { return new Workbook(); }
  addWorksheet(): Worksheet { return new Worksheet(); }
  save(): void {}
}
export class Worksheet {
  setZoom(): this { return this; }
  setName(): this { return this; }
}
export class Note {
  constructor() {}
  setWidth(): this { return this; }
}
`;

const root = fs.mkdtempSync(path.join(os.tmpdir(), 'coverage-refs-'));
afterAll(() => fs.rmSync(root, { recursive: true, force: true }));

function sitesFor(source: string): Map<string, number[]> {
  fs.mkdirSync(path.join(root, 'tests'), { recursive: true });
  fs.writeFileSync(path.join(root, 'tests', 'sample.test.mjs'), source);
  const sites = collectCallSites(
    root,
    ['tests/sample.test.mjs'],
    readApiShapes(WRAPPER, 'wrapper.ts'),
  );
  return new Map([...sites].map(([key, list]) => [key, list.map((site) => site.line)]));
}

describe('collectCallSites', () => {
  test('types receivers through constructors, factories and `this` chains', () => {
    const sites = sitesFor(
      [
        "import { Workbook, Note } from '../wrapper.ts';",
        'const workbook = Workbook.create();',
        'const sheet = workbook.addWorksheet()',
        '  .setZoom()',
        '  .setName();',
        'new Note().setWidth();',
      ].join('\n'),
    );
    expect(sites.get('Workbook.static.create')).toEqual([2]);
    expect(sites.get('Workbook.prototype.addWorksheet')).toEqual([3]);
    // Chained calls start where their receiver expression starts.
    expect(sites.get('Worksheet.prototype.setZoom')).toEqual([3]);
    expect(sites.get('Worksheet.prototype.setName')).toEqual([3]);
    expect(sites.get('Note.prototype.setWidth')).toEqual([6]);
  });

  test('does not guess the receiver of same-named methods', () => {
    const sites = sitesFor(
      [
        "import { Note } from '../wrapper.ts';",
        'function f(note) { note.setWidth(); }',
        'unknown.setWidth();',
      ].join('\n'),
    );
    expect(sites.get('Note.prototype.setWidth')).toBeUndefined();
  });

  test('follows hoisted helpers, destructuring, reassignment and Object.create', () => {
    const sites = sitesFor(
      [
        "import { Workbook } from '../wrapper.ts';",
        'const { workbook: wb, worksheet } = makeBook();',
        'wb.save();',
        'worksheet.setZoom();',
        'let later;',
        'later = Workbook.create();',
        'later.save();',
        'Object.create(Workbook.prototype).save();',
        'function makeBook() {',
        '  const workbook = Workbook.create();',
        '  return { workbook, worksheet: workbook.addWorksheet() };',
        '}',
      ].join('\n'),
    );
    expect(sites.get('Workbook.prototype.save')).toEqual([3, 7, 8]);
    expect(sites.get('Worksheet.prototype.setZoom')).toEqual([4]);
  });

  test('records constructors and the literal arguments of calls', () => {
    const sites = sitesFor(
      [
        "import { Note, Workbook } from '../wrapper.ts';",
        "const label = 'kept';",
        "new Note('memo').setWidth(10, 'wide', label);",
        'Workbook.create().addWorksheet().setZoom(-5);',
        'new Note().setWidth(width);',
      ].join('\n'),
    );
    expect(sites.get('Note::constructor')).toEqual([3, 5]);
    expect(sites.get("Note::constructor('memo')")).toEqual([3]);
    expect(sites.get('Note.prototype.setWidth')).toEqual([3, 5]);
    expect(sites.get('Note.prototype.setWidth(10)')).toEqual([3]);
    expect(sites.get("Note.prototype.setWidth('wide')")).toEqual([3]);
    expect(sites.get("Note.prototype.setWidth('kept')")).toEqual([3]);
    expect(sites.get('Worksheet.prototype.setZoom(-5)')).toEqual([4]);
  });

  test('binds .each callback parameters to the literals of their table column', () => {
    const sites = sitesFor(
      [
        "import { Note } from '../wrapper.ts';",
        "const WIDTHS = ['narrow', 'wide'];",
        'test.each(WIDTHS)("%s", (width) => new Note().setWidth(width));',
        "it.each([[1, 'one'], [2, 'two']])('%i', (count, name) => {",
        '  new Note().setWidth(name);',
        '});',
        "describe.each([['x', { not: 'literal' }]])('%s', (_label, options) => {",
        '  new Note().setWidth(options);',
        '});',
      ].join('\n'),
    );
    expect(sites.get("Note.prototype.setWidth('narrow')")).toEqual([3]);
    expect(sites.get("Note.prototype.setWidth('wide')")).toEqual([3]);
    expect(sites.get("Note.prototype.setWidth('one')")).toEqual([5]);
    expect(sites.get("Note.prototype.setWidth('two')")).toEqual([5]);
    expect(sites.get('Note.prototype.setWidth(1)')).toBeUndefined();
    expect(sites.get('Note.prototype.setWidth')).toEqual([3, 5, 8]);
    expect([...sites.keys()].filter((key) => key.includes("('x')"))).toEqual([]);
  });

  test('records throwing assertions but not negated ones', () => {
    const sites = sitesFor(
      [
        'expect(() => f()).toThrow();',
        'expect(() => g()).not.toThrow();',
        'expect(h).toThrow(/x/);',
      ].join('\n'),
    );
    expect(sites.get(ERROR_BOUNDARY)).toEqual([1, 3]);
  });
});

test('declaredApis lists public members and skips private constructors', () => {
  expect([...declaredApis(WRAPPER, 'wrapper.ts')].toSorted()).toEqual([
    'Note.prototype.setWidth',
    'Note::constructor',
    'Workbook.prototype.addWorksheet',
    'Workbook.prototype.save',
    'Workbook.static.create',
    'Worksheet.prototype.setName',
    'Worksheet.prototype.setZoom',
  ]);
});

test('API keys name an API, optionally with one literal argument', () => {
  const keys = [
    'Format.prototype.setBold',
    'Image.static.fromBuffer',
    'Chart::constructor',
    "Chart::constructor('radar')",
    'Table.prototype.setStyle(30)',
    "Note.prototype.setAuthor('O\\'Brien')",
  ];
  expect(keys.filter((key) => !API_KEY.test(key))).toEqual([]);
  const notKeys = ['Button class', 'Format.setBold', "Table.prototype.setStyle('a', 'b')"];
  expect(notKeys.filter((key) => API_KEY.test(key))).toEqual([]);
  expect(API_KEY.exec("Chart::constructor('radar')")?.[1]).toBe('Chart::constructor');
});

test('literalArgument renders numbers bare and quotes strings', () => {
  expect(literalArgument(30)).toBe('(30)');
  expect(literalArgument('left')).toBe("('left')");
  expect(literalArgument("O'Brien")).toBe("('O\\'Brien')");
});
