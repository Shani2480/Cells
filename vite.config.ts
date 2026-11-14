import path from "node:path"
import { fileURLToPath } from "node:url"
import tailwindcss from "@tailwindcss/vite"
import react from "@vitejs/plugin-react"
import { defineConfig } from "vitest/config"

const host = process.env.TAURI_DEV_HOST
const rootDir = path.dirname(fileURLToPath(import.meta.url))

export default defineConfig(() => ({
  root: path.resolve(rootDir, "app"),
  publicDir: path.resolve(rootDir, "app/public"),
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: { "@": path.resolve(rootDir, "app/src") },
  },
  build: {
    outDir: path.resolve(rootDir, "dist"),
    emptyOutDir: true,
  },
  test: {
    root: rootDir,
    projects: [
      {
        extends: true,
        test: { name: "core", environment: "node", include: ["packages/**/*.test.ts"] },
      },
      {
        extends: true,
        test: {
          name: "app",
          environment: "jsdom",
          setupFiles: ["app/src/test/setup.ts"],
          include: ["app/src/**/*.test.{ts,tsx}"],
        },
      },
      {
        extends: true,
        test: { name: "bench", environment: "node", include: ["packages/**/*.bench.ts"], testTimeout: 60_000 },
      },
    ],
  },
  // Tauri: fixed port, keep Rust errors visible, don't watch the Rust side.
  clearScreen: false,
  server: {
    port: 1430,
    strictPort: true,
    host: host || false,
    hmr: host ? { protocol: "ws", host, port: 1431 } : undefined,
    watch: { ignored: ["**/src-tauri/**"] },
  },
}))
