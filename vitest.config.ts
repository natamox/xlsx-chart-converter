import path from 'node:path';
import { defineConfig } from 'vitest/config';

export default defineConfig({
  resolve: {
    alias: {
      '@natamox/xlsx-chart-converter-core': path.resolve(import.meta.dirname, 'packages/core/src/index.ts'),
      '@natamox/xlsx-chart-converter-echarts': path.resolve(import.meta.dirname, 'packages/echarts/src/index.ts'),
      '@natamox/xlsx-chart-converter-svg': path.resolve(import.meta.dirname, 'packages/svg/src/index.ts')
    }
  },
  test: {
    coverage: {
      reportsDirectory: './coverage'
    },
    environment: 'node',
    globals: true,
    include: ['packages/*/test/**/*.test.ts']
  }
});
