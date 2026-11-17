import { readFileSync } from "node:fs"
import { join } from "node:path"
import { parseRange, selectRange } from "@cells/core"
import { describe, expect, it } from "vitest"
import { spellerFrom } from "@/proofing/speller"
import { SheetController } from "./controller"
import { SpellingSession } from "./spelling"

const root = join(__dirname, "../../..")
const dict = () =>
  spellerFrom(
    readFileSync(join(root, "node_modules/dictionary-en/index.aff"), "utf8"),
    readFileSync(join(root, "node_modules/dictionary-en/index.dic"), "utf8"),
  )

function setup(texts: [string, string][]) {
  const c = new SheetController()
  const cells = texts.map(([ref, value]) => {
    const r = parseRange(ref)!
    return { row: r.r0, col: r.c0, data: value.startsWith("=") ? { value: 0, input: value } : { value } }
  })
  c.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "x", before: c.selection })
  return c
}

describe("Spelling", () => {
  it("goes through the sheet's text, skipping formulas, and changes a word in place", () => {
    const c = setup([
      ["A1", "Monthly recieve report"],
      ["B1", '="teh"'],
      ["A2", "Teh total"],
    ])
    const s = new SpellingSession(c, dict())
    const first = s.next()!
    expect(first).toMatchObject({ row: 0, col: 0, word: "recieve" })
    expect(first.suggestions).toContain("receive")
    const second = s.change("receive")!
    expect(c.sheet.cells.value(0, 0)).toBe("Monthly receive report")
    expect(second).toMatchObject({ row: 1, col: 0, word: "Teh" })
    expect(s.change("The")).toBeNull()
    expect(c.sheet.cells.value(1, 0)).toBe("The total")
    // Each change is its own undo step.
    c.undo()
    expect(c.sheet.cells.value(1, 0)).toBe("Teh total")
  })

  it("Ignore All, Change All and Add to Dictionary last for the rest of the check", () => {
    const c = setup([
      ["A1", "zorbly teh"],
      ["A2", "zorbly teh Teh"],
      ["A3", "flurbo"],
    ])
    const added: string[] = []
    const s = new SpellingSession(c, dict(), (w) => added.push(w))
    expect(s.next()!.word).toBe("zorbly")
    expect(s.ignoreAll()!.word).toBe("teh")
    expect(s.changeEvery("the")!.word).toBe("flurbo")
    expect(c.sheet.cells.value(1, 0)).toBe("zorbly the The")
    expect(s.addToDictionary()).toBeNull()
    expect(added).toEqual(["flurbo"])
  })

  it("checks only the selection when it's more than one cell", () => {
    const c = setup([
      ["A1", "teh"],
      ["C3", "wrnog"],
    ])
    c.setSelection(selectRange(parseRange("B2:D4")!))
    const s = new SpellingSession(c, dict())
    expect(s.next()).toMatchObject({ row: 2, col: 2, word: "wrnog" })
    expect(s.ignoreOnce()).toBeNull()
  })
})

describe("Spelling from a merged cell", () => {
  it("checks the whole sheet when the active cell is one merged cell", () => {
    const c = setup([
      ["A1", "Title"],
      ["C5", "wrnog"],
    ])
    c.wb.transact([{ kind: "merge", sheet: 0, ranges: [parseRange("A1:B2")!] }], { label: "m", before: c.selection })
    c.setSelection(selectRange(parseRange("A1:B2")!))
    expect(new SpellingSession(c, dict()).next()).toMatchObject({ row: 4, col: 2, word: "wrnog" })
  })
})
