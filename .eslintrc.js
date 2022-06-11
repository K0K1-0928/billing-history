module.exports = {
  env: {
    browser: true,
    commonjs: true,
    es2021: true,
    'googleappsscript/googleappsscript': true,
  },
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'prettier',
  ],
  parser: '@typescript-eslint/parser',
  parserOptions: {
    ecmaVersion: 'latest',
  },
  plugins: ['@typescript-eslint', 'googleappsscript'],
  rules: {},
};
