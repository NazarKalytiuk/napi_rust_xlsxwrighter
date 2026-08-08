# I Got Tired of Node.js Running Out of Memory Generating Excel Files — So I Rewrote It in Rust

**How a Rust-powered npm package exports 10 million rows to Excel using just 21 KB of RAM**

---

> **Title Options (pick your favorite):**
>
> 1. **I Got Tired of Node.js Running Out of Memory Generating Excel Files — So I Rewrote It in Rust** — *How a Rust-powered npm package exports 10 million rows to Excel using just 21 KB of RAM*
> 2. **Exporting 10 Million Excel Rows With 0.02 MB of RAM: A Rust + Node.js Story** — *Meet rusc-xlsx — the write-only Excel library that laughs at your heap limits*
> 3. **Your Node.js Excel Exports Are Eating All Your RAM. Here's the Rust Fix.** — *Introducing constant-memory XLSX generation for Node.js via native Rust bindings*

---

If you've ever built a data-heavy Node.js application — a reporting dashboard, an admin panel, an ETL pipeline — you've probably hit **that** wall. The one where your perfectly working Excel export suddenly crashes in production with:

```
FATAL ERROR: CALL_AND_RETRY_LAST Allocation failed - JavaScript heap out of memory
```

It's 2 AM. Your Slack is on fire. A client needs "just a simple export" of 500,000 rows. And your Node.js process just folded like a lawn chair.

I've been there. Many of us have.

The problem isn't your code. The problem is that pure JavaScript Excel libraries need to hold your entire dataset in memory as JavaScript objects — arrays of arrays of objects, each one tracked by V8's garbage collector. At scale, this is a death sentence. You either bump `--max-old-space-size` to absurd levels, stream to CSV instead (and lose all formatting), or tell your PM that "Excel doesn't support that many rows."

None of those answers felt good enough. So I built something better.

---

## Introducing `@nazarkalytiuk/rusc-xlsx` 🦀

[Insert Hero Banner Image Here]

**[`@nazarkalytiuk/rusc-xlsx`](https://www.npmjs.com/package/@nazarkalytiuk/rusc-xlsx)** is a fast, feature-complete Excel (.xlsx) writer library for Node.js. Under the hood, it wraps the battle-tested Rust crate [`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter) via [NAPI-RS](https://napi.rs/), giving you native-level performance without leaving the Node.js ecosystem.

What does that mean in practice? It means your data goes from JavaScript → Rust → a proper `.xlsx` file on disk, and **the heavy lifting never touches V8's heap**. Rust manages its own memory, its own allocations, its own lifecycle. Node.js just orchestrates.

The result: you can write millions of rows without breaking a sweat — or a heap limit.

---

## The Killer Feature: Memory Modes 💾

This is where things get wild.

`rusc-xlsx` offers **three memory modes**, and you pick the one that fits your workload:

- **Standard Mode** — Full feature set, optimized for typical datasets. Think tens of thousands of rows with charts, formatting, and data validation.
- **Low Memory Mode** — Optimized for large datasets with repeated string values. Uses a shared string table to cut memory dramatically.
- **Constant Memory Mode** — The showstopper. Rows are flushed to disk as you write them. Memory usage stays flat *no matter how many rows you write*.

Here are real benchmark numbers for a 50/50 string-and-number dataset:

| Rows | Standard Mode | Low Memory Mode | Constant Memory Mode |
|---|---|---|---|
| 100K | 18.0 MB | 3.0 MB | **0.0215 MB** |
| 1M | 216.8 MB | 41.7 MB | **0.0215 MB** |
| 10M | ~2.2 GB | ~420 MB | **0.0215 MB** |

Read that last row again. **Ten million rows. 0.0215 MB. That's 21 kilobytes.**

The Standard mode already beats most pure-JS libraries. Low Memory mode is a solid middle ground. But Constant Memory mode is in a different universe entirely. Memory usage doesn't grow. It's flat. A horizontal line on the chart while your row count climbs into the millions.

> 🚀 **Why this matters for Serverless & Cloud:**
> AWS Lambda functions have a hard memory ceiling (typically 128 MB–3 GB). Cloud Run and Azure Functions have similar limits. With Constant Memory mode, you can generate massive Excel reports inside a Lambda function without ever worrying about hitting that wall. Your infrastructure costs stay low, and your exports stay reliable.

[Insert benchmark bar chart comparing 2.2 GB vs 0.02 MB here]

---

## Getting Started

Installation is one line:

```bash
npm install @nazarkalytiuk/rusc-xlsx
```

Prebuilt binaries are provided for Windows (x64, ARM64), macOS (Universal), and Linux (x64, ARM64, musl) — so there's no Rust toolchain needed for most users.

### Basic Example: Formatted Headers in 10 Lines

```javascript
const { Workbook, Format } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

const headerFormat = new Format()
  .setBold()
  .setBackgroundColor('blue')
  .setFontColor('white');

worksheet.write(0, 0, 'Name', headerFormat);
worksheet.write(0, 1, 'Age', headerFormat);
worksheet.write(1, 0, 'Alice');
worksheet.write(1, 1, 30);

workbook.save('output.xlsx');
```

That's it. A clean, formatted Excel file in under a second. The API is intentionally simple: create a workbook, add worksheets, write cells, save.

[Insert screenshot of generated Excel file]

### Constant Memory Mode: For Massive Datasets

When you're dealing with hundreds of thousands (or millions) of rows, switch to Constant Memory mode with a single method call:

```javascript
const { Workbook } = require('@nazarkalytiuk/rusc-xlsx');

const workbook = new Workbook();
// This is the magic line:
const worksheet = workbook.addWorksheetWithConstantMemory('Massive Data');

for (let row = 0; row < 1_000_000; row++) {
  worksheet.write(row, 0, `Row ${row}`);
  worksheet.write(row, 1, row * 100);
}

workbook.save('large_export.xlsx');
```

That's one million rows written with roughly **21 KB of RAM**. The API is identical — you just swap `addWorksheet` for `addWorksheetWithConstantMemory`. Your existing code needs almost zero changes.

[Insert Carbon.now.sh screenshot of this code block]

---

## Not Just Fast — Feature Complete 📊

Speed without features is just a CSV writer. `rusc-xlsx` doesn't make you choose.

Here's what you get out of the box: **25 chart types** (bar, column, line, pie, scatter, radar, stock, and more), full **conditional formatting** (color scales, data bars, icon sets, top-N rules, formula-based rules), **Excel tables** with 60 built-in styles and total rows, **sparklines** (line, column, win/lose) with 36 styles, **data validation** (dropdowns, number ranges, custom rules with input/error messages), **images** (PNG, JPG, GIF, BMP with scaling and positioning), **rich text** within cells, **hyperlinks**, **named ranges**, **cell comments**, **page setup** for print, **freeze panes**, **autofilters**, and **worksheet protection** with granular permissions.

And everything ships with **100% TypeScript definitions** — full IntelliSense, autocomplete, and type safety in VS Code without any extra setup.

> ⚠️ **One thing to know:** `rusc-xlsx` is a **write-only** library. It generates `.xlsx` files — it doesn't read or parse existing ones. This is by design. By focusing exclusively on the write path, every optimization goes toward making generation as fast and memory-efficient as possible. If you need to read Excel files, pair it with a reader library like `xlsx` or `exceljs` for the read side.

---

## Try It Out 🚀

If you're building anything in Node.js that exports data to Excel — reports, admin dashboards, ETL pipelines, SaaS data exports — give `rusc-xlsx` a spin.

```bash
npm install @nazarkalytiuk/rusc-xlsx
```

Check out the **[GitHub repository](https://github.com/NazarKalytiuk/napi_rust_xlsxwrighter)** for full documentation, examples, and the source code. And if it saves you from a 2 AM heap crash, consider dropping a ⭐️ on the repo — it means more than you'd think for open-source maintainers.

I'd love to hear what you build with it. Find me on GitHub, or drop a comment below.

Happy exporting. 🦀📊

---

**Tags:** `Nodejs` · `Rust` · `Web Development` · `JavaScript` · `Open Source`

---

> **📝 Post-Publication Checklist:**
>
> - **Hero Image:** Create a 1200×630 cover image featuring the Node.js logo + Rust logo (Ferris 🦀) + Excel logo
> - **Code Screenshots:** Paste the Constant Memory code block into [carbon.now.sh](https://carbon.now.sh/) and embed the image above the text code block
> - **Benchmark Chart:** Create a bar chart comparing Standard (2.2 GB) vs Constant Memory (0.02 MB) at 10M rows — the visual contrast is striking
> - **Medium Tags:** `Nodejs`, `Rust`, `Web Development`, `JavaScript`, `Open Source`
> - **Reddit Cross-Post:** Post to `r/node`, `r/javascript`, and `r/rust` with the title: *"I got tired of Node.js running out of memory generating Excel files, so I rewrote it in Rust (0.02MB RAM for 1M rows)."*
