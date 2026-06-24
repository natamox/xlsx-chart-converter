import js from '@eslint/js';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  {
    ignores: [
      'dist/**',
      'packages/*/dist/**',
      'packages/*/src/**/*.d.ts',
      'coverage/**',
      'node_modules/**'
    ]
  },
  js.configs.recommended,
  ...tseslint.configs.strictTypeChecked.map((config) => ({
    ...config,
    files: ['**/*.ts']
  })),
  {
    files: ['**/*.ts'],
    languageOptions: {
      parserOptions: {
        project: './tsconfig.eslint.json',
        tsconfigRootDir: import.meta.dirname
      }
    },
    rules: {
      '@typescript-eslint/consistent-type-imports': ['error', { prefer: 'type-imports' }],
      '@typescript-eslint/no-confusing-void-expression': 'off',
      '@typescript-eslint/no-empty-object-type': 'off',
      '@typescript-eslint/no-extraneous-class': 'off',
      '@typescript-eslint/restrict-template-expressions': [
        'error',
        { allowNumber: true, allowBoolean: true }
      ]
    }
  }
);
