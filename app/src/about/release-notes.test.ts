import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { CURRENT_VERSION, RELEASE_NOTES, whatsNewOnStart } from "./release-notes"

const root = join(__dirname, "../../..")
const version = (file: string) => (JSON.parse(readFileSync(join(root, file), "utf8")) as { version: string }).version

describe("release notes", () => {
  it("the newest notes are for the version being built", () => {
    expect(RELEASE_NOTES[0].version).toBe(version("package.json"))
    expect(RELEASE_NOTES[0].version).toBe(version("src-tauri/tauri.conf.json"))
    expect(RELEASE_NOTES[0].notes.length).toBeGreaterThan(0)
  })

  it("the Cargo version and the Flatpak's release list follow the notes", () => {
    const cargo = readFileSync(join(root, "src-tauri/Cargo.toml"), "utf8")
    expect(/^version = "([^"]+)"/m.exec(cargo)?.[1]).toBe(CURRENT_VERSION)
    const metainfo = readFileSync(join(root, "packaging/flatpak/app.cells.desktop.metainfo.xml"), "utf8")
    const listed = [...metainfo.matchAll(/<release version="([^"]+)" date="([^"]+)"/g)].map((m) => `${m[1]} ${m[2]}`)
    expect(listed).toEqual(RELEASE_NOTES.map((n) => `${n.version} ${n.date}`))
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

  it("What's New opens once after an update, not on a first start", () => {
    const store = new Map<string, string>()
    const storage = {
      getItem: (k: string) => store.get(k) ?? null,
      setItem: (k: string, v: string) => void store.set(k, v),
    }
    expect(whatsNewOnStart(storage)).toBe(false)
    expect(whatsNewOnStart(storage)).toBe(false)
    store.set("cells.seen-version", "0.0.1")
    expect(whatsNewOnStart(storage)).toBe(true)
    expect(store.get("cells.seen-version")).toBe(CURRENT_VERSION)
    expect(whatsNewOnStart(storage)).toBe(false)
  })
})
