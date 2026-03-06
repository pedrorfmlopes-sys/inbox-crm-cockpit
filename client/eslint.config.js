import js from "@eslint/js";
import { createRequire } from "node:module";
import reactHooks from "eslint-plugin-react-hooks";
import reactRefresh from "eslint-plugin-react-refresh";

const require = createRequire(import.meta.url);

let tsParser = null;
let tsPlugin = null;

try {
  tsParser = require("@typescript-eslint/parser");
  tsPlugin = require("@typescript-eslint/eslint-plugin");
} catch {
  // Keep lint runnable in restricted environments where TS ESLint packages cannot be installed.
}

const config = [
  {
    ignores: ["dist/**", "node_modules/**", "build/**", "public/**/*.min.mjs"],
  },
  js.configs.recommended,
  {
    files: ["**/*.{js,jsx,mjs,cjs}"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "module",
      globals: {
        console: "readonly",
        process: "readonly",
      },
    },
    plugins: {
      "react-hooks": reactHooks,
      "react-refresh": reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      "react-refresh/only-export-components": ["warn", { allowConstantExport: true }],
    },
  },
];

if (tsParser && tsPlugin) {
  config.push({
    files: ["src/**/*.{ts,tsx}"],
    languageOptions: {
      parser: tsParser,
      ecmaVersion: 2020,
      sourceType: "module",
    },
    plugins: {
      "@typescript-eslint": tsPlugin,
      "react-hooks": reactHooks,
      "react-refresh": reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      "react-refresh/only-export-components": ["warn", { allowConstantExport: true }],
      "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
      "no-undef": "off",
    },
  });
} else {
  const message = "[eslint] @typescript-eslint packages are unavailable; TypeScript files are skipped in this environment.";
  if (process.env.CI === "true") {
    throw new Error(`${message} Install devDependencies before running CI lint.`);
  }
  console.warn(message);
  config.push({
    ignores: ["src/**/*.{ts,tsx}"],
  });
}

export default config;
