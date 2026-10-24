import js from '@eslint/js'
import prettierConfig from 'eslint-config-prettier/flat'
import prettierPlugin from 'eslint-plugin-prettier'
import tseslint from 'typescript-eslint'

export default tseslint.config(
  {ignores: ['dist/', 'node_modules/']},
  js.configs.recommended,
  tseslint.configs.recommended,
  prettierConfig,
  {
    files: ['**/*.ts'],
    languageOptions: {sourceType: 'script'},
    plugins: {prettier: prettierPlugin},
    rules: {
      // Apps Script files share one global scope, so top-level functions and constants are used across files
      // (and by menu handlers resolved by name), which per-file analysis would flag as unused.
      '@typescript-eslint/no-unused-vars': ['error', {vars: 'local'}],
      'no-console': 'warn',
      'prettier/prettier': 'error',
    },
  },
)
