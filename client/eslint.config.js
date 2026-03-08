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
  // Keep lint runnable in environments where TS ESLint packages are unavailable.
}

const baseConfig = {
  files: ["**/*.{js,jsx,ts,tsx,mjs,cjs}"],
  languageOptions: {
    ecmaVersion: 2020,
    sourceType: "module",
    parserOptions: {
      ecmaFeatures: {
        jsx: true,
      },
    },
    globals: {
      console: "readonly",
      process: "readonly",
      window: "readonly",
      document: "readonly",
      navigator: "readonly",
      localStorage: "readonly",
      fetch: "readonly",
      CustomEvent: "readonly",
      ResizeObserver: "readonly",
      URL: "readonly",
      URLSearchParams: "readonly",
      AbortController: "readonly",
      setTimeout: "readonly",
      clearTimeout: "readonly",
      setInterval: "readonly",
      clearInterval: "readonly",
      Office: "readonly",
      OfficeRuntime: "readonly",
    },
  },
  plugins: {
    "react-hooks": reactHooks,
    "react-refresh": reactRefresh,
  },
  rules: {
    ...reactHooks.configs.recommended.rules,
    "react-refresh/only-export-components": ["warn", { allowConstantExport: true }],
    "no-undef": "off",
  },
};

const config = [
  {
    ignores: [
      "dist/**",
      "node_modules/**",
      "build/**",
      "public/**",
      "**/*.min.js",
      "**/*.min.mjs",
      "**/*.min.css",
    ],
  },
  js.configs.recommended,
  baseConfig,
];

if (tsParser && tsPlugin) {
  config.push({
    files: ["src/**/*.{ts,tsx}"],
    languageOptions: {
      ...baseConfig.languageOptions,
      parser: tsParser,
    },
    plugins: {
      ...baseConfig.plugins,
      "@typescript-eslint": tsPlugin,
    },
    rules: {
      ...baseConfig.rules,
      "@typescript-eslint/no-unused-vars": ["warn", {
        argsIgnorePattern: "^_",
        varsIgnorePattern: "^_",
        caughtErrorsIgnorePattern: "^_",
      }],
      "@typescript-eslint/no-unused-expressions": "warn",
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/ban-ts-comment": "warn",
      "@typescript-eslint/no-empty-function": "warn",
      "@typescript-eslint/no-empty-object-type": "off",
      "@typescript-eslint/no-wrapper-object-types": "off",
      "@typescript-eslint/no-unsafe-function-type": "off",
      "no-redeclare": "off",
      "no-unused-vars": "off",
      "no-unused-expressions": "off",
      "no-empty": "warn",
      "no-case-declarations": "warn",
      "no-useless-escape": "warn",
      "no-dupe-else-if": "warn",
      "prefer-const": "warn",
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
