import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.mjs', 'scripts/**/*.test.mts'],
    coverage: {
      provider: 'v8',
      include: ['wrapper.ts'],
      reporter: ['text', 'html', 'lcov'],
    },
  },
});
