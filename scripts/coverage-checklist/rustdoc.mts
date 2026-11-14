// Produces rustdoc JSON for a rust_xlsxwriter release with a given feature set.
//
// rustdoc's JSON output is unstable, so it needs a nightly toolchain; the
// toolchain is pinned so that the JSON format (and thus the inventory) only
// changes when this file does. Each feature set is documented from a throwaway
// Cargo workspace under target/, which keeps the project's Cargo.lock untouched.

import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

import type { RustdocCrate } from './inventory.mts';

export const RUSTDOC_TOOLCHAIN = 'nightly-2026-11-13';
const FORMAT_VERSION = 61;

export interface RustdocRequest {
  version: string;
  features: string[];
}

export function rustdocJson(workDir: string, request: RustdocRequest): RustdocCrate {
  const features = [...request.features].toSorted();
  const name = `${request.version}-${features.join('+') || 'no-default'}`;
  const crateDir = path.join(workDir, 'crates', name);
  const output = path.join(workDir, 'json', `${name}.json`);
  if (!fs.existsSync(output)) {
    fs.mkdirSync(path.join(crateDir, 'src'), { recursive: true });
    fs.writeFileSync(path.join(crateDir, 'src', 'lib.rs'), '');
    fs.writeFileSync(
      path.join(crateDir, 'Cargo.toml'),
      [
        '[package]',
        'name = "coverage-inventory"',
        'version = "0.0.0"',
        'edition = "2024"',
        'publish = false',
        '',
        // An empty [workspace] keeps cargo from attaching to the project workspace.
        '[workspace]',
        '',
        '[dependencies]',
        `rust_xlsxwriter = { version = "=${request.version}", default-features = false, features = ${JSON.stringify(features)} }`,
        '',
      ].join('\n'),
    );
    const targetDir = path.join(workDir, 'target');
    execFileSync(
      'cargo',
      [
        `+${RUSTDOC_TOOLCHAIN}`,
        'rustdoc',
        '--quiet',
        '--package',
        'rust_xlsxwriter',
        '--lib',
        '--',
        '-Z',
        'unstable-options',
        '--output-format',
        'json',
      ],
      {
        cwd: crateDir,
        stdio: ['ignore', 'inherit', 'inherit'],
        env: { ...process.env, CARGO_TARGET_DIR: targetDir },
      },
    );
    fs.mkdirSync(path.dirname(output), { recursive: true });
    fs.copyFileSync(path.join(targetDir, 'doc', 'rust_xlsxwriter.json'), output);
  }
  const krate = JSON.parse(fs.readFileSync(output, 'utf8')) as RustdocCrate;
  if (krate.format_version !== FORMAT_VERSION) {
    throw new Error(
      `${RUSTDOC_TOOLCHAIN} produced rustdoc JSON format ${krate.format_version}, expected ${FORMAT_VERSION}`,
    );
  }
  if (krate.crate_version !== request.version) {
    throw new Error(
      `rustdoc JSON is for rust_xlsxwriter ${krate.crate_version}, expected ${request.version}`,
    );
  }
  return krate;
}

/** Feature names declared by a rust_xlsxwriter release (without `default`). */
export function declaredFeatures(workDir: string, version: string): string[] {
  const crateDir = path.join(workDir, 'crates', `${version}-metadata`);
  fs.mkdirSync(path.join(crateDir, 'src'), { recursive: true });
  fs.writeFileSync(path.join(crateDir, 'src', 'lib.rs'), '');
  fs.writeFileSync(
    path.join(crateDir, 'Cargo.toml'),
    `[package]\nname = "coverage-inventory"\nversion = "0.0.0"\nedition = "2024"\npublish = false\n\n[workspace]\n\n[dependencies]\nrust_xlsxwriter = "=${version}"\n`,
  );
  const metadata = JSON.parse(
    execFileSync('cargo', ['metadata', '--format-version', '1'], {
      cwd: crateDir,
      encoding: 'utf8',
      maxBuffer: 64 * 1024 * 1024,
    }),
  );
  const pkg = metadata.packages.find(
    (candidate: any) => candidate.name === 'rust_xlsxwriter' && candidate.version === version,
  );
  return Object.keys(pkg.features)
    .filter((feature) => feature !== 'default')
    .toSorted();
}

/** The rust_xlsxwriter version and resolved features this project compiles. */
export function projectBaseline(projectRoot: string): RustdocRequest {
  const metadata = JSON.parse(
    execFileSync('cargo', ['metadata', '--format-version', '1', '--locked'], {
      cwd: projectRoot,
      encoding: 'utf8',
      maxBuffer: 64 * 1024 * 1024,
    }),
  );
  const pkg = metadata.packages.find((candidate: any) => candidate.name === 'rust_xlsxwriter');
  const node = metadata.resolve.nodes.find((candidate: any) => candidate.id === pkg.id);
  return { version: pkg.version, features: [...node.features].toSorted() };
}
