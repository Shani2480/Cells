import { describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"
import { NOT_FOUND, NOTHING_TO_REPLACE } from "./find"

function book() {
  const c = new SheetController()
  const type = (row: number, col: number, text: string) => {
    c.goTo({ r0: row, c0: col, r1: row, c1: col })
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  type(0, 0, "apple")
  type(0, 1, "10")
  type(1, 0, "=B1*2")
  type(2, 1, "pineapple")
  c.sheets.insert(1)
  c.activate(1)
  type(4, 4, "Apple tart")
  c.activate(0)
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  return { c, type }
}
const at = (c: SheetController) => [c.sheetIndex, c.selection.active.row, c.selection.active.col]

describe("Find", () => {
  it("Find Next goes through the sheet and round; Workbook goes on to the next sheet", () => {
    const { c } = book()
    c.find.next({ text: "apple", lookIn: "formulas", within: "sheet" })
    expect(at(c)).toEqual([0, 2, 1])
    c.find.next()
    expect(at(c)).toEqual([0, 0, 0])
    c.find.next({ text: "apple", lookIn: "formulas", within: "workbook" })
    expect(at(c)).toEqual([0, 2, 1])
    c.find.next()
    expect(at(c)).toEqual([1, 4, 4])
    c.find.next()
    expect(at(c)).toEqual([0, 0, 0])
  })

  it("says so when nothing matches; Find All lists the matches", () => {
    const { c } = book()
    const alert = vi.fn()
    c.shell.alert = alert
    expect(c.find.next({ text: "kiwi", lookIn: "formulas", within: "sheet" })).toBe(false)
    expect(alert).toHaveBeenCalledWith(NOT_FOUND)
    const all = c.find.all({ text: "20", lookIn: "values", within: "workbook" })
    expect(all.map((x) => [x.sheetName, x.cell, x.value, x.formula])).toEqual([["Sheet1", "$A$2", "20", "=B1*2"]])
  })

  it("searches only the selection when more than one cell is selected, keeping it", () => {
    const { c } = book()
    c.goTo({ r0: 0, c0: 0, r1: 1, c1: 1 })
    c.find.next({ text: "apple", lookIn: "formulas", within: "sheet" })
    expect(c.selection.ranges).toEqual([{ r0: 0, c0: 0, r1: 1, c1: 1 }])
    expect(c.selection.active).toEqual({ row: 0, col: 0 })
  })
})

describe("Replace", () => {
  it("Replace All types the cells again (numbers stay numbers, formulas recalculate), in one undo step", () => {
    const { c } = book()
    const alert = vi.fn()
    c.shell.alert = alert
    expect(c.find.replaceAll({ text: "1", lookIn: "formulas", within: "sheet" }, "3")).toBe(2)
    expect(alert).toHaveBeenCalledWith("All done. We made 2 replacements.")
    expect(c.sheet.cells.get(0, 1)?.value).toBe(30)
    expect(c.sheet.cells.get(1, 0)?.input).toBe("=B3*2")
    c.undo()
    expect(c.sheet.cells.get(0, 1)?.value).toBe(10)
    expect(c.sheet.cells.get(1, 0)?.input).toBe("=B1*2")
  })

  it("Replace changes the active cell when it matches, then finds the next", () => {
    const { c } = book()
    c.shell.alert = vi.fn()
    c.find.replace({ text: "apple", lookIn: "formulas", within: "sheet" }, "pear")
    expect(c.sheet.cells.get(0, 0)?.value).toBe("pear")
    expect(at(c)).toEqual([0, 2, 1])
    c.find.replace({ text: "apple", lookIn: "formulas", within: "sheet" }, "pear")
    expect(c.sheet.cells.get(2, 1)?.value).toBe("pinepear")
    const alert = vi.fn()
    c.shell.alert = alert
    c.find.replaceAll({ text: "apple", lookIn: "formulas", within: "sheet" }, "x")
    expect(alert).toHaveBeenCalledWith(NOTHING_TO_REPLACE)
  })
})

describe("Go To Special", () => {
  it("selects the formulas, or says none were found", () => {
    const { c } = book()
    c.find.special({ kind: "formulas", types: { numbers: true, text: true, logicals: true, errors: true } })
    expect(c.selection.ranges).toEqual([{ r0: 1, c0: 0, r1: 1, c1: 0 }])
    const alert = vi.fn()
    c.shell.alert = alert
    expect(c.find.special({ kind: "notes" })).toBe(false)
    expect(alert).toHaveBeenCalledWith("No cells were found.")
  })
})
