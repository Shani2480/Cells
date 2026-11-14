// Fails if a pure package imports anything outside itself and the other pure packages (PLAN.md §2).
// ESLint and the DOM-free tsconfig catch most of this; this check also covers bare package imports.
import { readdirSync, readFileSync, statSync } from "node:fs"
import path from "node:path"

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..")
const allowed = {
  core: [],
  platform: ["@cells/core"],
  io: ["@cells/core", "@cells/chart", "@zip.js/zip.js", "saxes"],
  chart: [
    "d3-geo",
    "geojson",
    "saxes",
    "topojson-client",
    "world-atlas/countries-110m.json",
    "us-atlas/states-10m.json",
  ],
}
const importRe = /(?:import|export)\s[^"']*?from\s*["']([^"']+)["']|import\(\s*["']([^"']+)["']\s*\)/g
const problems = []

function walk(dir, files = []) {
  for (const name of readdirSync(dir)) {
    const full = path.join(dir, name)
    if (statSync(full).isDirectory()) walk(full, files)
    else if (/\.ts$/.test(name) && !/\.(test|bench)\.ts$/.test(name)) files.push(full)
  }
  return files
}

for (const [pkg, deps] of Object.entries(allowed)) {
  for (const file of walk(path.join(root, "packages", pkg, "src"))) {
    const text = readFileSync(file, "utf8")
    for (const match of text.matchAll(importRe)) {
      const spec = match[1] ?? match[2]
      if (spec.startsWith(".")) continue
      if (!deps.includes(spec)) problems.push(`${path.relative(root, file)} imports "${spec}"`)
    }
  }
}

if (problems.length) {
  console.error("Package boundary violations:\n  " + problems.join("\n  "))
  process.exit(1)
}
console.log("Package boundaries OK")
