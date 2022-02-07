module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'prettier',
  ],
  //   env: 'es2020',
  rules: {
    'no-extend-native': 'off',
    'no-var': 'off',
    'require-jsdoc': 'off',
    'valid-jsdoc': 'off',
    'no-unused-vars': ['off'],
    '@typescript-eslint/no-unused-vars': [
      'off',
      // 'error', { vars: 'local', args: 'after-used', ignoreRestSiblings: false },
    ],
  },
};
