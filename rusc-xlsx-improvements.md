# @nazarkalytiuk/rusc-xlsx — Issues & Improvements

## Bug: Reserved keyword in generated type definitions

**File:** `index.d.ts:383`

**Problem:** NAPI-RS auto-generated `function` as a parameter name in `RuscTableColumn.setTotalFunction()`. `function` is a reserved keyword in TypeScript — this causes a syntax error that `skipLibCheck: true` cannot suppress (it's a parse error, not a type error).

```typescript
// Current (broken)
setTotalFunction(function: string): this

// Fix
setTotalFunction(fn: string): this
```

**Where to fix:** The Rust source that generates the NAPI binding. The parameter name in the `#[napi]` macro probably uses `function` — rename it to `fn` or `total_function`. NAPI-RS will then generate the correct TypeScript declaration.

---

## Bug: RichText rejects blank strings

**Problem:** `writeRichString()` throws when a `RichText` segment has an empty string `""`:

```
Error: Failed to write rich string: Parameter error:
'Strings in rich string (&Format, &str) tuples cannot be blank.'
```

This is a rust_xlsxwriter constraint, but it's unexpected for JS consumers. ExcelJS silently handles empty rich text cells.

**Options:**
1. Skip empty segments silently in the wrapper (recommended)
2. Write a blank cell with `writeBlank()` as fallback when all segments are empty
3. Document the limitation

**Suggested wrapper fix in `wrapper.js`:**
```javascript
writeRichString(row, col, richText) {
  // If all segments are blank, write a blank cell instead
  if (!richText._hasContent()) {
    return; // or this._native.writeBlank(row, col, ...)
  }
  this._native.writeRichString(row, col, richText._native);
}
```

---

## Feature: `saveToStream(writable)`

**Problem:** Currently only `save(filename)` and `saveToBuffer()` are available. For server-side use cases (streaming XLSX to cloud storage, HTTP responses), consumers must either:
- `saveToBuffer()` — holds entire file in memory (bad for large exports)
- `save(tempFile)` + `createReadStream()` — works but adds disk I/O latency

**Proposed API:**
```typescript
saveToStream(writable: import('stream').Writable): void;
```

**Implementation approach:** rust_xlsxwriter's `Workbook::save_to_writer()` accepts any `impl std::io::Write`. Create a NAPI wrapper that:
1. Accepts a Node.js `Writable` stream via `napi::JsObject`
2. Implements `std::io::Write` by calling `.write()` on the JS stream
3. Handles backpressure via `drain` events

**Benchmark motivation:**

| Metric | exceljs (streaming) | rusc-xlsx `saveToBuffer()` | rusc-xlsx `save(tempFile)` |
|---|---|---|---|
| Duration | 9,006ms | 3,961ms | 6,793ms |
| Heap delta | 42.03MB | 360.23MB | -11.75MB |
| RSS delta | 178.19MB | 1,012.31MB | 70.11MB |

`saveToStream()` would combine the speed of `saveToBuffer()` with the memory profile of `save(tempFile)` — no disk I/O, no buffer allocation.

---

## Feature: `addWorksheetWithConstantMemory()` documentation

The constant memory mode works well and keeps WASM memory near-zero, but it's not obvious from the README that it should be combined with `save(filename)` rather than `saveToBuffer()` for maximum benefit. `saveToBuffer()` defeats the purpose since it materializes the entire file in memory at the end.

Add a note to the README:
> For large exports, use `addWorksheetWithConstantMemory()` with `save(filename)` — not `saveToBuffer()`. The constant memory mode flushes rows to temp files as they're written, but `saveToBuffer()` would load the entire result back into memory.
