import { describe, expect, it } from "vitest"
import { cellName, DEFAULT_PROTECTION, parseRange, selectRange } from "@cells/core"
import { SheetController } from "./controller"

/**
 * Protect Sheet with "Select locked cells" off: only unlocked cells (B2, D2, B4, D5, F3 here) can be selected.
 * Arrows go to the nearest unlocked cell that way, Tab and Enter go round them like a form's fields, and a click or
 * drag over locked cells doesn't move the cursor.
 */
function setup(selectLocked = false) {
  const c = new SheetController()
  const unlocked = c.wb.styles.intern({ ...c.wb.styles.get(0), locked: false })
  for (const ref of ["B2", "D2", "B4", "D5", "F3"]) {
    const { r0, c0 } = parseRange(ref)!
    c.sheet.cells.set(r0, c0, { value: ref, style: unlocked })
  }
  c.setSelection(selectRange(parseRange("A1")!))
  c.sheets.protect({ ...DEFAULT_PROTECTION, selectLocked, selectUnlocked: true }, "")
  const at = () => cellName(c.selection.active)
  return { c, at }
}

describe("Select locked cells off", () => {
  it("protecting moves the cursor to the first unlocked cell", () => {
    expect(setup().at()).toBe("B2")
    expect(setup(true).at()).toBe("A1")
  })

  it("arrows go to the nearest unlocked cell that way, or stay", () => {
    const { c, at } = setup()
    c.arrow("right", false, false)
    expect(at()).toBe("D2")
    c.arrow("right", false, false)
    expect(at()).toBe("D2")
    c.arrow("down", false, false)
    expect(at()).toBe("D5")
    c.arrow("left", false, false)
    expect(at()).toBe("D5")
  })

  it("Tab goes row by row and wraps; Shift+Tab goes back; Enter goes column by column", () => {
    const { c, at } = setup()
    const tabs = [at()]
    for (let i = 0; i < 5; i++) {
      c.enterMove("right")
      tabs.push(at())
    }
    expect(tabs).toEqual(["B2", "D2", "F3", "B4", "D5", "B2"])
    c.enterMove("left")
    expect(at()).toBe("D5")
    c.enterMove("down")
    expect(at()).toBe("F3")
    c.enterMove("down")
    expect(at()).toBe("B2")
    c.enterMove("down")
    expect(at()).toBe("B4")
  })

  it("a click or drag over locked cells is refused", () => {
    const { c, at } = setup()
    c.selectCellAt({ row: 0, col: 0 })
    expect(at()).toBe("B2")
    c.dragTo({ row: 2, col: 2 })
    expect(c.selection.ranges[0]).toEqual({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.selectCellAt({ row: 4, col: 3 })
    expect(at()).toBe("D5")
  })
})
