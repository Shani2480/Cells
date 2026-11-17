import { DEFAULT_PROTECTION, hashPassword, parseRange, selectCell } from "@cells/core"
import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { parseRefersTo, refersToText } from "./edit-ranges"

function setup() {
  const c = new SheetController()
  const alerts: string[] = []
  const unlocks: { name: string; then: () => void }[] = []
  c.shell = { ...c.shell, alert: (m) => alerts.push(m), unlockRange: (name, then) => unlocks.push({ name, then }) }
  c.editRanges.save([
    { name: "Inputs", ranges: [parseRange("B2:B5")!] },
    { name: "Rates", ranges: [parseRange("D1")!], password: hashPassword("pw") },
  ])
  const { password: _p, ...options } = DEFAULT_PROTECTION
  c.sheets.protect(options, "")
  return { c, alerts, unlocks }
}

const type = (c: SheetController, cell: string, text: string) => {
  const r = parseRange(cell)!
  c.setSelection(selectCell({ row: r.r0, col: r.c0 }))
  c.beginEdit("enter", text)
  if (c.edit) c.commitEdit()
}

describe("Allow Edit Ranges on a protected sheet", () => {
  it("lets cells of a range without a password change; others refuse with Excel's message", () => {
    const { c, alerts } = setup()
    type(c, "B3", "42")
    expect(c.sheet.cells.value(2, 1)).toBe(42)
    type(c, "A1", "x")
    expect(c.sheet.cells.value(0, 0)).toBeUndefined()
    expect(alerts).toHaveLength(1)
  })

  it("asks for a range's password first, then carries on with the entry", () => {
    const { c, unlocks } = setup()
    c.setSelection(selectCell({ row: 0, col: 3 }))
    c.beginEdit("enter", "7")
    expect(c.edit).toBeNull()
    expect(unlocks[0].name).toBe("Rates")
    expect(c.editRanges.unlock("Rates", "wrong")).toBe(false)
    expect(c.editRanges.unlock("Rates", "pw")).toBe(true)
    unlocks[0].then()
    c.commitEdit()
    expect(c.sheet.cells.value(0, 3)).toBe(7)
  })

  it("reads and writes the dialog's references", () => {
    expect(parseRefersTo("=$B$2:$B$5,$D$1")).toEqual([parseRange("B2:B5"), parseRange("D1")])
    expect(parseRefersTo("=Data!$A$1")).toEqual([parseRange("A1")])
    expect(parseRefersTo("=nope")).toBeNull()
    expect(refersToText([parseRange("B2:B5")!, parseRange("D1")!])).toBe("=$B$2:$B$5,$D$1")
  })
})
