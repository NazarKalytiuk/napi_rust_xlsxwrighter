import { configDefaults, defineConfig } from 'vitest/config';

// Runs the library tests against the WebAssembly build (`npm run build:wasm`)
// instead of the native addon. The flavor is set only for the test workers,
// and only the library tests run: rolldown inside Vitest and oxc-parser in the
// coverage-checklist scripts are napi-rs packages too, and NAPI_RS_WASI_FLAVOR
// would redirect them as well.
export default defineConfig({
  test: {
    include: ['tests/**/*.test.mjs'],
    exclude: [...configDefaults.exclude, 'tests/browser/**'],
    env: { NAPI_RS_WASI_FLAVOR: 'wasm32-wasip1' },
  },
});
