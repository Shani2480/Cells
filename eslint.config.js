import eslint from "@eslint/js"
import prettier from "eslint-config-prettier"
import globals from "globals"
import tseslint from "typescript-eslint"

/** The engine and the platform interfaces must stay free of UI, DOM and shell code (PLAN.md §2). */
const pureRestrictions = {
  patterns: [
    { group: ["react", "react-*", "react/*"], message: "packages/* are UI-free: no React." },
    { group: ["@tauri-apps/*"], message: "packages/* are platform-free: no Tauri." },
    { group: ["@/*", "**/app/**"], message: "packages/* cannot import the app." },
    { group: ["node:*"], message: "packages/* run in browsers and workers too: no Node built-ins." },
  ],
}

export default tseslint.config(
  {
    ignores: ["dist/**", "node_modules/**", "tmp/**", "src-tauri/**", "tools/**", "*.tsbuildinfo"],
  },
  eslint.configs.recommended,
  ...tseslint.configs.recommended,
  prettier,
  {
    files: ["**/*.{ts,tsx,js,mjs,cjs}"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      globals: { ...globals.browser, ...globals.node },
    },
    rules: {
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrorsIgnorePattern: "^_" },
      ],
      "@typescript-eslint/no-unused-expressions": ["error", { allowShortCircuit: true, allowTernary: true }],
    },
  },
  {
    files: ["packages/**/*.ts"],
    ignores: ["packages/**/*.test.ts", "packages/**/*.bench.ts"],
    languageOptions: { globals: {} },
    rules: {
      "no-restricted-imports": ["error", pureRestrictions],
      "no-restricted-globals": ["error", "window", "document", "navigator", "localStorage", "process"],
    },
  },
)
