#!/usr/bin/env bash
# Measures line coverage of the Rust sources by running the Vitest suite
# against an instrumented debug build of the native addon. `npm run
# test:coverage` only measures wrapper.ts, which forwards to Rust.
#
# Needs cargo-llvm-cov and the llvm-tools component of the active Rust
# toolchain. Writes coverage/rust.lcov, and the per-file report to the output
# and coverage/rust.txt; extra arguments go to that report, for example
# `--summary-only` or `--show-missing-lines`.
set -euo pipefail
cd "$(dirname "$0")/.."

# Instruments this crate only and writes profiles under target/.
eval "$(cargo llvm-cov show-env --sh)"
cargo llvm-cov clean --profraw-only
host="$(rustc -vV | sed -n 's/^host: //p')"

# The instrumented addon replaces any build in the project root, where the
# loader looks for it; the builds there are set aside and put back afterwards.
shopt -s nullglob
set_aside="$(mktemp -d)"
restore_builds() {
  rm -f rusc_xlsx.*.node
  for build in "$set_aside"/*.node; do
    mv "$build" .
  done
  rmdir "$set_aside"
}
trap restore_builds EXIT
for build in rusc_xlsx.*.node; do
  mv "$build" "$set_aside/"
done
# The generated loader and type definitions stay as the release build made them.
npx napi build --platform --no-js --dts target/rust-coverage/index.d.ts --target "$host"

# Forked workers are stopped with SIGTERM before LLVM writes their profile, so
# the suite runs in worker threads of one process that exits normally. Files
# run one at a time because the allocator statistics cover the whole process.
npx vitest run --pool=threads --no-file-parallelism

mkdir -p coverage
cargo llvm-cov report --target "$host" --lcov --output-path coverage/rust.lcov
cargo llvm-cov report --target "$host" "$@" | tee coverage/rust.txt
