const { FlatCompat } = require("@eslint/eslintrc");
const typescriptPlugin = require("@typescript-eslint/eslint-plugin");

const compat = new FlatCompat({
  baseDirectory: __dirname,
  recommendedConfig: {
    parser: "@typescript-eslint/parser",
    parserOptions: {
      project: ["./tsconfig.json"],
      sourceType: "module",
    },
  },
});

module.exports = [
  ...compat.extends(
    "eslint:recommended",
    "plugin:@typescript-eslint/recommended",
    "prettier"
  ),

  {
    files: ["src/**/*.{ts,tsx}"],
    languageOptions: {
      parserOptions: {
        ecmaVersion: 2020,
        sourceType: "module",
      },
    },

    plugins: {
      "@typescript-eslint": typescriptPlugin,
    },

    rules: {},
  },
];
