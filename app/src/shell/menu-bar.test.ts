import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { isSystemShortcut, keyInputFrom } from "@/sheet/keys"
import { isTextField, MENU_BAR_IDS, menuBarAction } from "./menu-bar"

const key = (key: string, mods: { ctrl?: boolean; meta?: boolean; shift?: boolean; alt?: boolean } = {}) => ({
  key,
  ctrlKey: Boolean(mods.ctrl),
  metaKey: Boolean(mods.meta),
  shiftKey: Boolean(mods.shift),
  altKey: Boolean(mods.alt),
})

describe("the macOS menu bar", () => {
  it("handles every item menu.rs sends, and menu.rs sends every one", () => {
    const rust = readFileSync(join(__dirname, "../../../src-tauri/src/menu.rs"), "utf8")
    const table = rust.slice(rust.indexOf("const ITEMS"), rust.indexOf("];"))
    const ids = [...table.matchAll(/\("([A-Za-z]+)", "/g)].map((m) => m[1])
    expect([...ids].sort()).toEqual([...MENU_BAR_IDS].sort())
    for (const id of ids) expect(menuBarAction(id, false), id).not.toBeNull()
  })

  it("Undo, Redo and Select All work on the text being edited, else on the workbook", () => {
    expect(menuBarAction("Undo", false)).toEqual({ kind: "command", id: "Undo" })
    expect(menuBarAction("Undo", true)).toEqual({ kind: "text", command: "undo" })
    expect(menuBarAction("Redo", true)).toEqual({ kind: "text", command: "redo" })
    expect(menuBarAction("SelectAll", false)).toEqual({ kind: "selectAll" })
    expect(menuBarAction("SelectAll", true)).toEqual({ kind: "text", command: "selectAll" })
    expect(menuBarAction("Quit", false)).toEqual({ kind: "quit" })
    expect(menuBarAction("Settings", false)).toEqual({ kind: "dialog", name: "options" })
    expect(menuBarAction("nonsense", false)).toBeNull()
  })

  it("tells a text field from the grid", () => {
    const grid = document.createElement("textarea")
    grid.setAttribute("data-grid-input", "")
    const box = document.createElement("input")
    const check = document.createElement("input")
    check.type = "checkbox"
    const button = document.createElement("button")
    expect(isTextField(grid, false)).toBe(false)
    expect(isTextField(grid, true)).toBe(true)
    expect(isTextField(box, false)).toBe(true)
    expect(isTextField(check, false)).toBe(false)
    expect(isTextField(button, false)).toBe(false)
    expect(isTextField(null, false)).toBe(false)
  })
})

describe("shortcut keys on macOS", () => {
  it("⌘ is the primary key, and Control for Replace, Go To and Show Formulas", () => {
    expect(keyInputFrom(key("s", { meta: true }), "macos").primary).toBe(true)
    expect(keyInputFrom(key("s", { ctrl: true }), "macos").primary).toBe(false)
    expect(keyInputFrom(key("h", { ctrl: true }), "macos").primary).toBe(true)
    expect(keyInputFrom(key("g", { ctrl: true }), "macos").primary).toBe(true)
    expect(keyInputFrom(key("`", { ctrl: true }), "macos").primary).toBe(true)
    // Elsewhere Ctrl is the primary key and ⌘ (the Windows key) isn't.
    expect(keyInputFrom(key("h", { ctrl: true }), "windows").primary).toBe(true)
    expect(keyInputFrom(key("s", { meta: true }), "linux").primary).toBe(false)
  })

  it("leaves macOS its own ⌘ shortcuts", () => {
    for (const k of ["h", "m", "q", ",", "`", "g", "H"])
      expect(isSystemShortcut(key(k, { meta: true }), "macos")).toBe(true)
    expect(isSystemShortcut(key("s", { meta: true }), "macos")).toBe(false)
    expect(isSystemShortcut(key("h", { ctrl: true }), "macos")).toBe(false)
    expect(isSystemShortcut(key("h", { meta: true }), "windows")).toBe(false)
  })
})
