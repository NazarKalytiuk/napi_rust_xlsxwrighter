// Finds where the tests exercise each public wrapper API, so checklist rows can
// point at their call sites instead of hand-maintained line numbers.
//
// Receivers are typed with a small source-order inference over the test files:
// `new Format()` is a Format, `format.setBold()` returns what wrapper.ts
// declares (`this` keeps the receiver type), bindings carry the type of their
// latest assignment, local helpers return the type of what they return, and
// object literals keep their property types through destructuring.
// Calls on receivers of unknown type are skipped rather than guessed from the
// method name, because names such as `setWidth` exist on several classes.
//
// A call that passes a string or number literal is also recorded under the
// API key with that argument, e.g. `Format.prototype.setAlign('left')`, so an
// enum variant row can point at the calls that select it. Parameters of
// `test.each`/`it.each`/`describe.each` callbacks stand for every literal in
// their column of the table.

import fs from 'node:fs';
import path from 'node:path';
import { parseSync } from 'oxc-parser';

/** Call-site key for assertions that an operation throws. */
export const ERROR_BOUNDARY = 'error-boundary';

/**
 * A public API key, optionally narrowed to calls passing one literal argument:
 * `Format.prototype.setBold`, `Chart::constructor('radar')`, `Table.prototype.setStyle(30)`.
 * The first group is the API without the argument.
 */
export const API_KEY: RegExp =
  /^(\w+(?:\.(?:prototype|static)\.\w+|::constructor))(?:\((?:'(?:[^'\\]|\\.)*'|-?\d+(?:\.\d+)?)\))?$/;

type Literal = string | number;

/** The key suffix for calls passing `value`, e.g. `('left')` or `(30)`. */
export function literalArgument(value: Literal): string {
  return typeof value === 'number'
    ? `(${value})`
    : `('${value.replaceAll('\\', '\\\\').replaceAll("'", "\\'")}')`;
}

const EACH_RUNNERS = new Set(['test', 'it', 'describe']);

export interface CallSite {
  file: string;
  line: number;
}

interface ClassShape {
  methods: Map<string, string | null>;
  statics: Map<string, string | null>;
}

type ValueType =
  | { instance: string }
  | { klass: string }
  | { returns: ValueType | null }
  | { fields: Map<string, ValueType | null> }
  | { literals: Literal[] }
  | { elements: (ValueType | null)[] };

/** Reads method return types of the exported classes in wrapper.ts. */
export function readApiShapes(wrapperSource: string, fileName: string): Map<string, ClassShape> {
  const { program, errors } = parseSync(fileName, wrapperSource);
  if (errors.length) throw new Error(`Cannot parse ${fileName}: ${errors[0]!.message}`);
  const shapes = new Map<string, ClassShape>();
  for (const statement of program.body as any[]) {
    const declaration =
      statement.type === 'ExportNamedDeclaration' ? statement.declaration : statement;
    if (declaration?.type !== 'ClassDeclaration') continue;
    const className: string = declaration.id.name;
    const shape: ClassShape = { methods: new Map(), statics: new Map() };
    for (const member of declaration.body.body) {
      if (member.type !== 'MethodDefinition' || member.kind !== 'method') continue;
      const returned = returnTypeName(member.value.returnType?.typeAnnotation, className);
      (member.static ? shape.statics : shape.methods).set(member.key.name, returned);
    }
    shapes.set(className, shape);
  }
  return shapes;
}

/**
 * Public API keys (`Format.prototype.setBold`, `Image.static.fromBuffer`,
 * `Format::constructor`) declared by the classes in a TypeScript source.
 */
export function declaredApis(source: string, fileName: string): Set<string> {
  const { program, errors } = parseSync(fileName, source);
  if (errors.length) throw new Error(`Cannot parse ${fileName}: ${errors[0]!.message}`);
  const apis = new Set<string>();
  for (const statement of program.body as any[]) {
    const declaration =
      statement.type === 'ExportNamedDeclaration' ? statement.declaration : statement;
    if (declaration?.type !== 'ClassDeclaration') continue;
    const className: string = declaration.id.name;
    for (const member of declaration.body.body) {
      if (member.type !== 'MethodDefinition' || member.accessibility === 'private') continue;
      if (member.kind === 'constructor') apis.add(`${className}::constructor`);
      else apis.add(`${className}.${member.static ? 'static' : 'prototype'}.${member.key.name}`);
    }
  }
  return apis;
}

function returnTypeName(annotation: any, className: string): string | null {
  if (annotation?.type === 'TSThisType') return className;
  if (annotation?.type === 'TSTypeReference' && annotation.typeName.type === 'Identifier') {
    return annotation.typeName.name;
  }
  return null;
}

const FUNCTION_TYPES = new Set([
  'FunctionDeclaration',
  'FunctionExpression',
  'ArrowFunctionExpression',
]);

/**
 * Collects call sites per API key (`Format.prototype.setBold`,
 * `Image.static.fromBuffer`, {@link ERROR_BOUNDARY}), ordered by file and line.
 */
export function collectCallSites(
  root: string,
  testFiles: string[],
  shapes: Map<string, ClassShape>,
): Map<string, CallSite[]> {
  const sites = new Map<string, CallSite[]>();
  for (const file of [...testFiles].toSorted()) {
    const source = fs.readFileSync(path.join(root, file), 'utf8');
    const { program, errors } = parseSync(file, source);
    if (errors.length) throw new Error(`Cannot parse ${file}: ${errors[0]!.message}`);
    const lineOf = lineLocator(source);
    // The first pass only infers what local helpers return, so that calls to a
    // helper declared further down are typed in the second, recording pass.
    const helperReturns = new Map<number, ValueType | null>();
    walkFile(program, shapes, helperReturns, null);
    walkFile(program, shapes, helperReturns, (key, offset) => {
      const list = sites.get(key) ?? [];
      const line = lineOf(offset);
      const last = list.at(-1);
      if (!last || last.file !== file || last.line !== line) list.push({ file, line });
      sites.set(key, list);
    });
  }
  return sites;
}

function walkFile(
  program: any,
  shapes: Map<string, ClassShape>,
  helperReturns: Map<number, ValueType | null>,
  record: ((key: string, offset: number) => void) | null,
): void {
  const scopes: Map<string, ValueType | null>[] = [];
  const functions: { start: number; returns: ValueType | null }[] = [];
  // Parameter types of `.each` callbacks, keyed by the callback's start offset.
  const eachParams = new Map<number, (ValueType | null)[]>();
  const lookup = (name: string): ValueType | null => {
    for (let i = scopes.length - 1; i >= 0; i--) {
      if (scopes[i]!.has(name)) return scopes[i]!.get(name) ?? null;
    }
    return null;
  };
  const assign = (name: string, type: ValueType | null): void => {
    for (let i = scopes.length - 1; i >= 0; i--) {
      if (scopes[i]!.has(name)) {
        scopes[i]!.set(name, type);
        return;
      }
    }
  };
  const helperType = (fn: any): ValueType => ({ returns: helperReturns.get(fn.start) ?? null });

  const typeOf = (node: any): ValueType | null => {
    switch (node?.type) {
      case 'Identifier':
        return lookup(node.name);
      case 'Literal':
        return typeof node.value === 'string' || typeof node.value === 'number'
          ? { literals: [node.value] }
          : null;
      case 'UnaryExpression':
        return node.operator === '-' &&
          node.argument.type === 'Literal' &&
          typeof node.argument.value === 'number'
          ? { literals: [-node.argument.value] }
          : null;
      case 'ArrayExpression':
        return { elements: node.elements.map(typeOf) };
      case 'ParenthesizedExpression':
        return typeOf(node.expression);
      case 'AwaitExpression':
        return typeOf(node.argument);
      case 'NewExpression': {
        const callee = typeOf(node.callee);
        return callee && 'klass' in callee ? { instance: callee.klass } : null;
      }
      case 'CallExpression': {
        const callee = node.callee;
        if (callee.type === 'Identifier') {
          const fn = lookup(callee.name);
          return fn && 'returns' in fn ? fn.returns : null;
        }
        if (callee.type !== 'MemberExpression' || callee.computed) return null;
        const prototypeOwner = objectCreatePrototype(node);
        if (prototypeOwner) return prototypeOwner;
        const receiver = typeOf(callee.object);
        if (!receiver || !('klass' in receiver || 'instance' in receiver)) return null;
        const isStatic = 'klass' in receiver;
        const shape = shapes.get(isStatic ? receiver.klass : receiver.instance);
        const returned = (isStatic ? shape?.statics : shape?.methods)?.get(callee.property.name);
        return returned && shapes.has(returned) ? { instance: returned } : null;
      }
      case 'ObjectExpression': {
        const fields = new Map<string, ValueType | null>();
        for (const property of node.properties) {
          if (
            property.type === 'Property' &&
            !property.computed &&
            property.key.type === 'Identifier'
          ) {
            fields.set(property.key.name, typeOf(property.value));
          }
        }
        return { fields };
      }
      case 'MemberExpression': {
        if (node.computed) return null;
        const owner = typeOf(node.object);
        return owner && 'fields' in owner ? (owner.fields.get(node.property.name) ?? null) : null;
      }
      default:
        return null;
    }
  };

  /** Binds `const { workbook, worksheet: sheet } = …` to the source's property types. */
  const bindPattern = (pattern: any, source: ValueType | null): void => {
    for (const property of pattern.properties) {
      if (property.type !== 'Property') continue;
      const target =
        property.value.type === 'AssignmentPattern' ? property.value.left : property.value;
      if (target.type !== 'Identifier') continue;
      const key =
        !property.computed && property.key.type === 'Identifier' ? property.key.name : null;
      const type = key && source && 'fields' in source ? (source.fields.get(key) ?? null) : null;
      scopes.at(-1)!.set(target.name, type);
    }
  };

  // `Object.create(Workbook.prototype)` builds a Workbook without its constructor.
  const objectCreatePrototype = (call: any): ValueType | null => {
    const { callee } = call;
    const [argument] = call.arguments;
    if (callee.object.type !== 'Identifier' || callee.object.name !== 'Object') return null;
    if (callee.property.name !== 'create' || argument?.type !== 'MemberExpression') return null;
    if (argument.computed || argument.property.name !== 'prototype') return null;
    const owner = typeOf(argument.object);
    return owner && 'klass' in owner ? { instance: owner.klass } : null;
  };

  /** Records `key` and, for each literal the call passes, `key('literal')`. */
  const recordCall = (key: string, call: any): void => {
    if (!record) return;
    record(key, call.start);
    const literals = new Set<Literal>();
    for (const argument of call.arguments) {
      const type = typeOf(argument);
      if (type && 'literals' in type) for (const value of type.literals) literals.add(value);
    }
    for (const value of literals) record(`${key}${literalArgument(value)}`, call.start);
  };

  /**
   * Types the callback parameters of `test.each(table)(name, callback)`: rows
   * that are arrays spread into the parameters, other rows are the only one.
   */
  const bindEachParams = (call: any): void => {
    const runner = call.callee;
    if (runner.type !== 'CallExpression' || runner.callee.type !== 'MemberExpression') return;
    const { object, property } = runner.callee;
    if (object.type !== 'Identifier' || !EACH_RUNNERS.has(object.name)) return;
    if (property.name !== 'each') return;
    const table = typeOf(runner.arguments[0]);
    const callback = call.arguments.find((argument: any) => FUNCTION_TYPES.has(argument.type));
    if (!table || !('elements' in table) || !callback) return;
    const columns = callback.params.map((_param: any, index: number): ValueType | null => {
      const values = table.elements.flatMap((row) => {
        const cell = row && 'elements' in row ? row.elements[index] : index === 0 ? row : null;
        return cell && 'literals' in cell ? cell.literals : [];
      });
      return values.length ? { literals: values } : null;
    });
    eachParams.set(callback.start, columns);
  };

  const hoist = (statements: any[]): void => {
    for (const statement of statements) {
      if (statement?.type === 'FunctionDeclaration' && statement.id) {
        scopes.at(-1)!.set(statement.id.name, helperType(statement));
      }
    }
  };

  const visit = (node: any): void => {
    if (!node || typeof node.type !== 'string') return;
    const isFunction = FUNCTION_TYPES.has(node.type);
    const opensScope = isFunction || node.type === 'BlockStatement' || node.type === 'Program';
    if (opensScope) {
      // Parameters shadow outer bindings; only `.each` callbacks know their types.
      const scope = new Map<string, ValueType | null>();
      const preset = eachParams.get(node.start);
      (node.params ?? []).forEach((param: any, index: number) => {
        if (param.type === 'Identifier') scope.set(param.name, preset?.[index] ?? null);
      });
      scopes.push(scope);
      if (!isFunction) hoist(node.body);
    }
    if (isFunction) functions.push({ start: node.start, returns: null });

    switch (node.type) {
      case 'ImportDeclaration':
        for (const specifier of node.specifiers) {
          const imported = specifier.imported?.name;
          if (imported && shapes.has(imported))
            scopes[0]!.set(specifier.local.name, { klass: imported });
        }
        break;
      case 'VariableDeclarator':
        visitChildren(node);
        if (node.id.type === 'Identifier') {
          const init = node.init;
          const type = FUNCTION_TYPES.has(init?.type) ? helperType(init) : typeOf(init);
          scopes.at(-1)!.set(node.id.name, type);
        } else if (node.id.type === 'ObjectPattern') {
          bindPattern(node.id, typeOf(node.init));
        }
        return;
      case 'AssignmentExpression':
        visitChildren(node);
        if (node.left.type === 'Identifier') assign(node.left.name, typeOf(node.right));
        return;
      case 'ReturnStatement': {
        const returned = typeOf(node.argument);
        const current = functions.at(-1);
        if (current && current.returns === null) current.returns = returned;
        break;
      }
      case 'CallExpression': {
        const callee = node.callee;
        bindEachParams(node);
        if (record && callee.type === 'MemberExpression' && !callee.computed) {
          const member: string = callee.property.name;
          const negated =
            callee.object.type === 'MemberExpression' && callee.object.property.name === 'not';
          if (member === 'toThrow' && !negated) record(ERROR_BOUNDARY, node.start);
          const receiver = typeOf(callee.object);
          if (receiver && 'klass' in receiver)
            recordCall(`${receiver.klass}.static.${member}`, node);
          if (receiver && 'instance' in receiver) {
            recordCall(`${receiver.instance}.prototype.${member}`, node);
          }
        }
        break;
      }
      case 'NewExpression': {
        const constructed = typeOf(node.callee);
        if (constructed && 'klass' in constructed) {
          recordCall(`${constructed.klass}::constructor`, node);
        }
        break;
      }
    }

    visitChildren(node);
    if (isFunction) {
      const fn = functions.pop()!;
      // Concise arrow bodies return their expression.
      const returned = node.expression ? typeOf(node.body) : null;
      if (record === null) helperReturns.set(node.start, fn.returns ?? returned);
    }
    if (opensScope) scopes.pop();
  };
  const visitChildren = (node: any): void => {
    for (const [key, value] of Object.entries(node)) {
      if (key === 'type' || key === 'start' || key === 'end') continue;
      if (Array.isArray(value)) value.forEach(visit);
      else if (value && typeof value === 'object') visit(value);
    }
  };
  visit(program);
}

function lineLocator(source: string): (offset: number) => number {
  const lineStarts = [0];
  for (let i = 0; i < source.length; i++) if (source[i] === '\n') lineStarts.push(i + 1);
  return (offset) => {
    let low = 0;
    let high = lineStarts.length - 1;
    while (low < high) {
      const mid = (low + high + 1) >> 1;
      if (lineStarts[mid]! <= offset) low = mid;
      else high = mid - 1;
    }
    return low + 1;
  };
}
