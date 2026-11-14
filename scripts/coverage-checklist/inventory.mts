// Builds the canonical symbol inventory of rust_xlsxwriter from rustdoc JSON.
//
// Row ids follow `<item path>#<kind>[::<member>#<kind>...]`, e.g.
// `rust_xlsxwriter::format::FormatPattern#enum::Solid#variant`. Enum variants
// with payload fields form one capability whose fields are nested rows.

// The subset of the rustdoc JSON format (format_version 61) used here.
type Id = number;
type Visibility = 'public' | 'default' | 'crate' | { restricted: unknown };
interface Item {
  id: Id;
  crate_id: number;
  name: string | null;
  visibility: Visibility;
  inner: Record<string, any>;
}
export interface RustdocCrate {
  root: Id;
  crate_version: string | null;
  format_version: number;
  index: Record<string, Item>;
  paths: Record<string, { crate_id: number; path: string[]; kind: string }>;
}

export interface InventoryRow {
  /** Canonical row id. */
  id: string;
  /** Capability the row belongs to (its own id unless it is a variant field). */
  capability: string;
  /** Module below the crate root, e.g. `format`. */
  module: string;
  /** Top-level item the row is listed under, e.g. `FormatPattern`. */
  section: string;
  /** Normalized signature, used to detect changed rows between versions. */
  signature: string;
}

const TOP_LEVEL_KINDS = new Set(['struct', 'enum', 'trait', 'type_alias', 'function']);

export function buildInventory(krate: RustdocCrate): Map<string, InventoryRow> {
  const rows = new Map<string, InventoryRow>();
  const item = (id: Id): Item => {
    const found = krate.index[String(id)];
    if (!found) throw new Error(`rustdoc item ${id} is missing from the index`);
    return found;
  };
  const typeName = (type: unknown): string => renderType(type, krate, true);

  for (const [id, summary] of Object.entries(krate.paths)) {
    if (summary.crate_id !== 0 || !TOP_LEVEL_KINDS.has(summary.kind)) continue;
    const top = krate.index[id];
    // Paths can reference items that are not part of the documented public API.
    if (!top) continue;
    const [, module, ...rest] = summary.path;
    if (!module || rest.length !== 1) continue;
    const section = rest[0]!;
    const base = `${summary.path.join('::')}#${summary.kind}`;
    const add = (rowId: string, signature: string, capability = rowId): void => {
      rows.set(rowId, { id: rowId, capability, module, section, signature });
    };

    const inner = top.inner[summary.kind];
    switch (summary.kind) {
      case 'struct': {
        add(base, structSignature(inner, item, typeName));
        const fields: Id[] = inner.kind?.plain?.fields ?? [];
        for (const fieldId of fields) {
          const field = item(fieldId);
          if (field.visibility !== 'public') continue;
          add(`${base}::${field.name}#field`, typeName(field.inner.struct_field));
        }
        addInherentMethods(inner.impls, base, add, item, typeName);
        break;
      }
      case 'enum': {
        add(base, '');
        for (const variantId of inner.variants as Id[]) {
          const variant = item(variantId);
          const variantRow = `${base}::${variant.name}#variant`;
          const fields = variantFields(variant, item);
          add(variantRow, fields.map(([, field]) => typeName(field.inner.struct_field)).join(','));
          for (const [name, field] of fields) {
            add(`${variantRow}::${name}#field`, typeName(field.inner.struct_field), variantRow);
          }
        }
        addInherentMethods(inner.impls, base, add, item, typeName);
        break;
      }
      case 'trait': {
        add(base, '');
        for (const memberId of inner.items as Id[]) {
          const member = item(memberId);
          if (!member.inner.function) continue;
          add(`${base}::${member.name}#function`, functionSignature(member, typeName));
        }
        for (const implId of inner.implementations as Id[]) {
          const impl = item(implId).inner.impl;
          if (impl.is_synthetic || impl.blanket_impl) continue;
          // Generic impls such as `impl<T: IntoExcelData> IntoExcelData for Option<T>`
          // forward to the concrete impls, which have rows of their own.
          if (impl.generics.params.some((param: any) => param.kind.type)) continue;
          add(`${base}-impl-for:${renderType(impl.for, krate, false)}`, '');
        }
        break;
      }
      case 'type_alias':
        add(base, typeName(inner.type));
        break;
      case 'function':
        add(base, functionSignature(top, typeName));
        break;
    }
  }
  return rows;
}

function addInherentMethods(
  implIds: Id[],
  base: string,
  add: (rowId: string, signature: string) => void,
  item: (id: Id) => Item,
  typeName: (type: unknown) => string,
): void {
  for (const implId of implIds) {
    const impl = item(implId).inner.impl;
    if (impl.trait || impl.is_synthetic || impl.blanket_impl) continue;
    for (const memberId of impl.items as Id[]) {
      const member = item(memberId);
      if (member.visibility !== 'public' || !member.inner.function) continue;
      add(`${base}::${member.name}#function`, functionSignature(member, typeName));
    }
  }
}

function variantFields(variant: Item, item: (id: Id) => Item): [string, Item][] {
  const kind = variant.inner.variant.kind;
  if (kind.tuple) {
    return (kind.tuple as (Id | null)[]).flatMap((fieldId, index): [string, Item][] =>
      fieldId === null ? [] : [[String(index), item(fieldId)]],
    );
  }
  if (kind.struct) {
    return (kind.struct.fields as Id[]).map((fieldId): [string, Item] => {
      const field = item(fieldId);
      return [field.name ?? '', field];
    });
  }
  return [];
}

function structSignature(
  inner: Record<string, any>,
  item: (id: Id) => Item,
  typeName: (type: unknown) => string,
): string {
  const fields: Id[] = inner.kind?.plain?.fields ?? [];
  return fields
    .map(item)
    .filter((field) => field.visibility === 'public')
    .map((field) => `${field.name}:${typeName(field.inner.struct_field)}`)
    .join(',');
}

function functionSignature(fn: Item, typeName: (type: unknown) => string): string {
  const sig = fn.inner.function.sig;
  const inputs = (sig.inputs as [string, unknown][]).map(([, type]) => typeName(type)).join(',');
  return `(${inputs})->${sig.output === null ? '()' : typeName(sig.output)}`;
}

/**
 * Renders a rustdoc type with full paths. Row ids spell types without generic
 * arguments (`withArgs = false`); signatures keep them so changes are visible.
 */
export function renderType(type: any, krate: RustdocCrate, withArgs: boolean): string {
  const render = (inner: unknown): string => renderType(inner, krate, withArgs);
  if (type === null || type === undefined) return '()';
  if (type.primitive) return type.primitive;
  if (type.generic) return type.generic;
  if (type.resolved_path) {
    const { id, path, args } = type.resolved_path;
    const name = krate.paths[String(id)]?.path.join('::') ?? path;
    const typeArgs = withArgs
      ? (args?.angle_bracketed?.args ?? [])
          .filter((arg: any) => arg.type)
          .map((arg: any) => render(arg.type))
      : [];
    return typeArgs.length ? `${name}<${typeArgs.join(',')}>` : name;
  }
  if (type.borrowed_ref) {
    const { is_mutable, type: target } = type.borrowed_ref;
    return `&${is_mutable ? 'mut ' : ''}${render(target)}`;
  }
  if (type.tuple) return `(${type.tuple.map(render).join(',')})`;
  if (type.slice) return `[${render(type.slice)}]`;
  if (type.array) return `[${render(type.array.type)};${type.array.len}]`;
  if (type.raw_pointer) {
    return `*${type.raw_pointer.is_mutable ? 'mut' : 'const'} ${render(type.raw_pointer.type)}`;
  }
  if (type.impl_trait) return 'impl';
  if (type.dyn_trait) return 'dyn';
  if (type.qualified_path) return `<${type.qualified_path.name}>`;
  if (type.function_pointer) return 'fn';
  return JSON.stringify(type);
}
