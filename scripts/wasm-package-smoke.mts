// Checks the packages a release publishes, the way a browser app uses them.
//
// It packs the root package and the WebAssembly package (npm/wasm32-wasip1,
// filled by `napi create-npm-dirs` and `npm run artifacts`, as the release
// workflow does) and installs both into a new Vite app. The app writes a
// workbook in Chromium, Firefox and WebKit, from a production build and from
// the Vite dev server, both without any Vite configuration, and each workbook
// is read back with ExcelJS. Last,
// Node.js must fall back to the WebAssembly package when no native package is
// installed.
//
// Usage: node scripts/wasm-package-smoke.mts [package root, default: cwd]

import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import ExcelJS from 'exceljs';
import { chromium, firefox, webkit } from 'playwright';

const require = createRequire(import.meta.url);
const packageRoot = path.resolve(process.argv[2] ?? '.');
const rootManifest = JSON.parse(fs.readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));
const wasmPackageDir = path.join(packageRoot, 'npm', 'wasm32-wasip1');
const wasmManifest = JSON.parse(fs.readFileSync(path.join(wasmPackageDir, 'package.json'), 'utf8'));
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';

function run(command: string, args: string[], cwd: string): string {
  return execFileSync(command, args, {
    cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'inherit'],
  });
}

function pack(dir: string, destination: string): string {
  // Scripts are skipped: the release workflow packs files it has already built.
  const [{ filename }] = JSON.parse(
    run(npm, ['pack', '--json', '--ignore-scripts', '--pack-destination', destination], dir),
  );
  return path.join(destination, filename);
}

const APP_SOURCE = `
import { Format, Workbook } from '${rootManifest.name}';

const workbook = new Workbook();
const worksheet = workbook.addWorksheet('Smoke');
worksheet.writeRow(0, 0, ['Browser', 'Rows'], new Format().setBold());
worksheet.writeRow(1, 0, [navigator.userAgent, 1000]);
for (let row = 2; row < 1002; row++) {
  worksheet.writeRow(row, 0, ['row ' + row, row]);
}
const buffer = workbook.saveToBuffer();
const link = document.createElement('a');
link.href = URL.createObjectURL(new Blob([buffer]));
link.download = 'smoke.xlsx';
link.textContent = 'Download';
document.body.append(link);
window.smokeResult = { base64: buffer.toString('base64') };
`;

const ENTRY_SOURCE = `
// Loading the app through import() turns a failed WebAssembly load into a
// rejection the smoke test can report.
import('./app.js').catch((error) => {
  window.smokeResult = { error: String(error?.stack ?? error) };
});
`;

function serve(dir: string): Promise<http.Server> {
  const types: Record<string, string> = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.wasm': 'application/wasm',
  };
  const server = http.createServer((request, response) => {
    const urlPath = new URL(request.url ?? '/', 'http://localhost').pathname;
    const file = path.join(dir, urlPath === '/' ? 'index.html' : urlPath);
    if (!file.startsWith(dir) || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
      response.writeHead(404).end();
      return;
    }
    response.writeHead(200, {
      'Content-Type': types[path.extname(file)] ?? 'application/octet-stream',
    });
    fs.createReadStream(file).pipe(response);
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)));
}

/** Opens a workbook from base64 text with ExcelJS. */
async function loadWorkbook(base64: string): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  // ExcelJS's typings take an ArrayBuffer rather than Node.js's Buffer.
  await workbook.xlsx.load(new Uint8Array(Buffer.from(base64, 'base64')).buffer);
  return workbook;
}

async function checkWorkbook(base64: string, label: string): Promise<void> {
  const workbook = await loadWorkbook(base64);
  const worksheet = workbook.getWorksheet('Smoke');
  if (!worksheet) {
    throw new Error(`${label}: the workbook has no Smoke worksheet`);
  }
  const header = worksheet.getCell('A1');
  const last = worksheet.getCell('A1002').text;
  if (header.text !== 'Browser' || !header.font?.bold || last !== 'row 1001') {
    throw new Error(`${label}: unexpected workbook contents (A1 ${header.text}, A1002 ${last})`);
  }
}

/** Opens the app in every browser and checks the workbook each one writes. */
async function checkBrowsers(url: string, label: string): Promise<void> {
  for (const browserType of [chromium, firefox, webkit]) {
    const name = `${label} / ${browserType.name()}`;
    const browser = await browserType.launch();
    try {
      const page = await browser.newPage();
      await page.goto(url);
      // These callbacks run in the page, where the app sets `smokeResult`.
      await page.waitForFunction(() => 'smokeResult' in globalThis, undefined, {
        timeout: 60_000,
      });
      const result = await page.evaluate(
        () =>
          (globalThis as unknown as { smokeResult: { base64?: string; error?: string } })
            .smokeResult,
      );
      if (result.error !== undefined || result.base64 === undefined) {
        throw new Error(`${name}: ${result.error}`);
      }
      await checkWorkbook(result.base64, name);
      console.log(`${name}: ok`);
    } finally {
      await browser.close();
    }
  }
}

const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'rusc-xlsx-wasm-smoke-'));
let server: http.Server | undefined;
try {
  const rootTarball = pack(packageRoot, workDir);
  const wasmTarball = pack(wasmPackageDir, workDir);

  const appDir = path.join(workDir, 'app');
  fs.mkdirSync(appDir);
  fs.writeFileSync(
    path.join(appDir, 'package.json'),
    JSON.stringify({
      private: true,
      type: 'module',
      dependencies: {
        [rootManifest.name]: `file:${rootTarball}`,
        [wasmManifest.name]: `file:${wasmTarball}`,
        vite: require('vite/package.json').version,
      },
    }),
  );
  run(npm, ['install', '--no-audit', '--no-fund'], appDir);
  fs.writeFileSync(
    path.join(appDir, 'index.html'),
    '<!doctype html><meta charset="utf-8"><title>smoke</title><script type="module" src="/main.js"></script>',
  );
  fs.writeFileSync(path.join(appDir, 'main.js'), ENTRY_SOURCE);
  fs.writeFileSync(path.join(appDir, 'app.js'), APP_SOURCE);
  run(process.execPath, [path.join('node_modules', 'vite', 'bin', 'vite.js'), 'build'], appDir);

  const distDir = path.join(appDir, 'dist');
  const assets = fs.readdirSync(path.join(distDir, 'assets'));
  if (!assets.some((asset) => asset.endsWith('.wasm'))) {
    throw new Error(`The production build has no .wasm asset: ${assets.join(', ')}`);
  }

  server = await serve(distDir);
  const address = server.address();
  if (address === null || typeof address === 'string') {
    throw new Error('The static server has no port');
  }
  await checkBrowsers(`http://127.0.0.1:${address.port}/`, 'production build');

  const { createServer } = await import(
    pathToFileURL(path.join(appDir, 'node_modules', 'vite', 'dist', 'node', 'index.js')).href
  );
  const devServer = await createServer({ root: appDir, logLevel: 'error', server: { port: 0 } });
  await devServer.listen();
  try {
    await checkBrowsers(devServer.resolvedUrls.local[0], 'dev server');
  } finally {
    await devServer.close();
  }

  // Without a native package, as on a platform with no prebuilt binary, the
  // Node.js loader falls back to the WebAssembly package.
  const scopeDir = path.join(appDir, 'node_modules', path.dirname(rootManifest.name));
  for (const entry of fs.readdirSync(scopeDir)) {
    const name = `${path.dirname(rootManifest.name)}/${entry}`;
    if (name.startsWith(`${rootManifest.name}-`) && name !== wasmManifest.name) {
      fs.rmSync(path.join(scopeDir, entry), { recursive: true });
    }
  }
  const nodeCheck = `
    const { __napiBindingTarget } = require('${rootManifest.name}/index.js');
    const { Workbook } = require('${rootManifest.name}');
    const workbook = new Workbook();
    workbook.addWorksheet('Smoke').writeRow(0, 0, ['Browser']);
    if (__napiBindingTarget !== 'wasm32-wasip1') throw new Error('loaded ' + __napiBindingTarget);
    process.stdout.write(workbook.saveToBuffer().toString('base64'));
  `;
  const nodeWorkbook = run(process.execPath, ['--no-warnings', '-e', nodeCheck], appDir);
  const nodeResult = await loadWorkbook(nodeWorkbook);
  if (nodeResult.getWorksheet('Smoke')?.getCell('A1').text !== 'Browser') {
    throw new Error('Node.js: unexpected workbook contents');
  }
  console.log('node (WebAssembly fallback): ok');
} finally {
  server?.close();
  fs.rmSync(workDir, { recursive: true, force: true });
}
