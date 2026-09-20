import js from '@eslint/js';
import googleappsscript from 'eslint-plugin-googleappsscript';

export default [
  { ignores: ['node_modules/**'] },
  js.configs.recommended,
  {
    files: ['src/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: googleappsscript.environments.googleappsscript.globals,
    },
    rules: {
      // All files share one global scope; the checkJs type check resolves cross-file references.
      'no-undef': 'off',
      // Public API functions are called by consuming projects, never within this one.
      'no-unused-vars': ['error', { vars: 'local', args: 'after-used' }],
      'no-var': 'error',
      'prefer-const': 'error',
      'prefer-template': 'error',
      'prefer-arrow-callback': 'error',
      'no-useless-catch': 'error',
    },
  },
];
