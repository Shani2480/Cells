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
  it("goes through the sheet's text, skipping formulas, and changes a word in place", async () => {
    const c = setup([
      ["A1", "Monthly recieve report"],
      ["B1", '="teh"'],
      ["A2", "Teh total"],
    ])
    const s = new SpellingSession(c, dict())
    const first = (await s.next())!
    expect(first).toMatchObject({ row: 0, col: 0, word: "recieve" })
    expect(first.suggestions).toContain("receive")
    const second = (await s.change("receive"))!
    expect(c.sheet.cells.value(0, 0)).toBe("Monthly receive report")
    expect(second).toMatchObject({ row: 1, col: 0, word: "Teh" })
    expect(await s.change("The")).toBeNull()
    expect(c.sheet.cells.value(1, 0)).toBe("The total")
    // Each change is its own undo step.
    c.undo()
    expect(c.sheet.cells.value(1, 0)).toBe("Teh total")
  })

  it("Ignore All, Change All and Add to Dictionary last for the rest of the check", async () => {
    const c = setup([
      ["A1", "zorbly teh"],
      ["A2", "zorbly teh Teh"],
      ["A3", "flurbo"],
    ])
    const added: string[] = []
    const s = new SpellingSession(c, dict(), (w) => added.push(w))
    expect((await s.next())!.word).toBe("zorbly")
    expect((await s.ignoreAll())!.word).toBe("teh")
    expect((await s.changeEvery("the"))!.word).toBe("flurbo")
    expect(c.sheet.cells.value(1, 0)).toBe("zorbly the The")
    expect(await s.addToDictionary()).toBeNull()
    expect(added).toEqual(["flurbo"])
  })

  it("checks only the selection when it's more than one cell", async () => {
    const c = setup([
      ["A1", "teh"],
      ["C3", "wrnog"],
    ])
    c.setSelection(selectRange(parseRange("B2:D4")!))
    const s = new SpellingSession(c, dict())
    expect(await s.next()).toMatchObject({ row: 2, col: 2, word: "wrnog" })
    expect(await s.ignoreOnce()).toBeNull()
  })
})

describe("Spelling from a merged cell", () => {
  it("checks the whole sheet when the active cell is one merged cell", async () => {
    const c = setup([
      ["A1", "Title"],
      ["C5", "wrnog"],
    ])
    c.wb.transact([{ kind: "merge", sheet: 0, ranges: [parseRange("A1:B2")!] }], { label: "m", before: c.selection })
    c.setSelection(selectRange(parseRange("A1:B2")!))
    expect(await new SpellingSession(c, dict()).next()).toMatchObject({ row: 4, col: 2, word: "wrnog" })
  })
})

describe("Spelling in another language", () => {
  it("picking a dictionary checks on from the current word in it", async () => {
    const c = setup([["A1", "huis recieve"]])
    // A dictionary that knows only Dutch words (and suggests nothing).
    const dutch = {
      check: async (words: readonly string[]) => words.map((w) => ["huis"].includes(w)),
      suggest: async () => [],
      add: () => undefined,
    }
    const s = new SpellingSession(c, dict())
    expect((await s.next())!.word).toBe("huis")
    expect((await s.useSpeller(dutch))!.word).toBe("recieve")
  })
})
