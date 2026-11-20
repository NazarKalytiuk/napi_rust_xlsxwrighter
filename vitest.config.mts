import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.mjs', 'scripts/**/*.test.mts'],
    // Browser tests run in browsers: `npm run test:browser`.
    exclude: [...configDefaults.exclude, 'tests/browser/**'],
    coverage: {
      provider: 'v8',
      include: ['wrapper.ts'],
      reporter: ['text', 'html', 'lcov'],
    },
  },
});
