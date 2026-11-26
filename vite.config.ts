import { createReadStream, readFileSync } from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"
import tailwindcss from "@tailwindcss/vite"
import react from "@vitejs/plugin-react"
import type { Plugin } from "vite"
import { defineConfig } from "vitest/config"

const host = process.env.TAURI_DEV_HOST
const rootDir = path.dirname(fileURLToPath(import.meta.url))

/**
 * Text recognition's files (From PDF on scanned pages, `app/src/pdf/ocr.ts`): Tesseract's worker, its three
 * WebAssembly builds (it picks one by what the processor can do) and the English model, served from node_modules in
 * development and copied into the build, so nothing is fetched from the internet.
 */
const OCR_FILES: Record<string, string> = {
  "ocr/worker.min.js": "node_modules/tesseract.js/dist/worker.min.js",
  "ocr/core/tesseract-core-lstm.wasm.js": "node_modules/tesseract.js-core/tesseract-core-lstm.wasm.js",
  "ocr/core/tesseract-core-simd-lstm.wasm.js": "node_modules/tesseract.js-core/tesseract-core-simd-lstm.wasm.js",
  "ocr/core/tesseract-core-relaxedsimd-lstm.wasm.js":
    "node_modules/tesseract.js-core/tesseract-core-relaxedsimd-lstm.wasm.js",
  "ocr/lang/eng.traineddata.gz": "node_modules/@tesseract.js-data/eng/4.0.0_best_int/eng.traineddata.gz",
}

function ocrFiles(): Plugin {
  return {
    name: "cells-ocr-files",
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const key = (req.url ?? "").split("?")[0].replace(/^\//, "")
        const source = OCR_FILES[key]
        if (!source) return next()
        res.setHeader("Content-Type", key.endsWith(".js") ? "text/javascript" : "application/octet-stream")
        createReadStream(path.resolve(rootDir, source)).pipe(res)
      })
    },
    generateBundle() {
      // Android reads scans with ML Kit (AndroidOcr.kt), so its app goes without Tesseract's 15 MB.
      if (process.env.TAURI_ENV_PLATFORM === "android") return
      for (const [fileName, source] of Object.entries(OCR_FILES))
        this.emitFile({ type: "asset", fileName, source: readFileSync(path.resolve(rootDir, source)) })
    },
  }
}

export default defineConfig(() => ({
  root: path.resolve(rootDir, "app"),
  publicDir: path.resolve(rootDir, "app/public"),
  plugins: [react(), tailwindcss(), ocrFiles()],
  resolve: {
    alias: {
      "@": path.resolve(rootDir, "app/src"),
      // Spelling: the English Hunspell files (the package itself only exports a Node loader).
      "dictionary-en-files": path.resolve(rootDir, "node_modules/dictionary-en"),
    },
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
