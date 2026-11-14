import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { RELEASE_NOTES } from "./release-notes"

const root = join(__dirname, "../../..")
const version = (file: string) => (JSON.parse(readFileSync(join(root, file), "utf8")) as { version: string }).version

describe("release notes", () => {
  it("the newest notes are for the version being built", () => {
    expect(RELEASE_NOTES[0].version).toBe(version("package.json"))
    expect(RELEASE_NOTES[0].version).toBe(version("src-tauri/tauri.conf.json"))
    expect(RELEASE_NOTES[0].notes.length).toBeGreaterThan(0)
  })

  it("versions are semver and newest first", () => {
    const parts = RELEASE_NOTES.map((n) => n.version.split(".").map(Number))
    for (const p of parts) expect(p.length === 3 && p.every(Number.isInteger)).toBe(true)
    for (let i = 1; i < parts.length; i++) {
      const [a, b] = [parts[i - 1], parts[i]]
      const newer = a[0] - b[0] || a[1] - b[1] || a[2] - b[2]
      expect(newer).toBeGreaterThan(0)
    }
  })
})
