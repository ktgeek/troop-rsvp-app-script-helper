import js from '@eslint/js'
import {defineConfig} from 'eslint/config'
import prettierConfig from 'eslint-config-prettier/flat'
import prettierPlugin from 'eslint-plugin-prettier'
import globals from 'globals'
import tseslint from 'typescript-eslint'

export default defineConfig(
  {ignores: ['dist/', 'node_modules/']},
  js.configs.recommended,
  tseslint.configs.recommended,
  prettierConfig,
  {
    // The deploy script and this config are plain Node ESM, so they get Node globals.
    files: ['**/*.mjs'],
    languageOptions: {globals: globals.node},
    plugins: {prettier: prettierPlugin},
    rules: {'prettier/prettier': 'error'},
  },
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
