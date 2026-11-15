'use strict';

// Release builds for the benchmarks that compare this package with direct
// rust_xlsxwriter calls. benchmarks/pure-rust is a member of the root Cargo
// workspace, so it compiles with the addon's Cargo.lock, release profile and
// rust_xlsxwriter declaration, and the two sides cannot drift apart.

const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const PURE_RUST_PACKAGE = 'rusc-xlsx-pure-rust-benchmark';

function runInherited(command, args) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    stdio: 'inherit',
    env: process.env,
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} exited with status ${result.status}`);
  }
}

function buildReleaseBinaries() {
  const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  process.stdout.write('Building the N-API package in release mode...\n');
  runInherited(npm, ['run', 'build']);

  process.stdout.write('Building the pure Rust benchmarks in release mode...\n');
  runInherited('cargo', ['build', '--release', '--locked', '--package', PURE_RUST_PACKAGE]);
}

function cargoMetadata() {
  const result = spawnSync('cargo', ['metadata', '--format-version', '1', '--locked'], {
    cwd: ROOT,
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
    env: process.env,
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`cargo metadata exited with status ${result.status}\n${result.stderr}`);
  }
  return JSON.parse(result.stdout);
}

/** Path of a release binary of the pure Rust benchmark package. */
function pureRustBinary(metadata, name) {
  return path.join(
    metadata.target_directory,
    'release',
    `${name}${process.platform === 'win32' ? '.exe' : ''}`,
  );
}

/** The rust_xlsxwriter release and resolved features both sides compile. */
function rustXlsxwriterBuild(metadata) {
  const pkg = metadata.packages.find((candidate) => candidate.name === 'rust_xlsxwriter');
  const node = metadata.resolve.nodes.find((candidate) => candidate.id === pkg.id);
  return {
    version: pkg.version,
    features: node.features.filter((feature) => feature !== 'default').toSorted(),
  };
}

module.exports = {
  ROOT,
  buildReleaseBinaries,
  cargoMetadata,
  pureRustBinary,
  rustXlsxwriterBuild,
};
