import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    coverage: {
      reportsDirectory: './coverage'
    },
    environment: 'node',
    globals: true,
    include: ['packages/*/test/**/*.test.ts']
  }
});
