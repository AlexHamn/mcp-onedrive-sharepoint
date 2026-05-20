// @ts-check

/**
 * ESLint flat config (migrated from .eslintrc.cjs for ESLint 9 + typescript-eslint 8).
 *
 * Goal: keep the original rule customizations (control-character regexes
 * stay allowed for security-validator, pre-existing debt stays at warn) while
 * moving to the modern flat-config shape required by ESLint 9.
 */

import js from "@eslint/js";
import tseslint from "typescript-eslint";
import globals from "globals";

export default tseslint.config(
  {
    ignores: ["build/", "node_modules/", "**/*.d.ts"],
  },

  js.configs.recommended,
  ...tseslint.configs.recommended,

  {
    files: ["src/**/*.ts", "scripts/**/*.{ts,mjs,cjs}"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        ...globals.node,
        ...globals.es2022,
      },
    },
    rules: {
      // Security validators legitimately match control characters in regexes.
      "no-control-regex": "off",

      // Pre-existing debt: track as warning until cleaned up.
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/no-unused-vars": [
        "warn",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
        },
      ],
      "no-unused-vars": "off",
      "no-useless-escape": "warn",
      "no-case-declarations": "warn",

      // ban-types was removed in typescript-eslint 8 (split into
      // no-restricted-types / no-empty-object-type / no-wrapper-object-types
      // / no-unsafe-function-type). The recommended preset already covers
      // the common cases, so we don't need a custom rule for it.

      // Empty function bodies appear in mocks and interface placeholders.
      "@typescript-eslint/no-empty-function": "off",

      // Scripts use CommonJS requires; TS checker handles type imports.
      "@typescript-eslint/no-require-imports": "off",
    },
  },

  {
    // .cjs / .mjs scripts use the default ESLint parser.
    files: ["scripts/**/*.{mjs,cjs}", "*.cjs"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        ...globals.node,
      },
    },
  },
);
