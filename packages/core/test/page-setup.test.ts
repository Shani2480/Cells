// A sheet's page setup: print area, titles and page breaks follow inserted and deleted rows and columns, as Excel's
// Print_Area and Print_Titles names do; undo puts them back.
import { describe, expect, it } from "vitest"
import { defaultPageSetup, shiftPageSetup, Workbook } from "../src"

describe("page setup", () => {
  it("follows inserted and deleted rows, and undo", () => {
    const wb = new Workbook(["Sheet1"])
    const setup = defaultPageSetup()
    setup.printArea = [{ r0: 0, c0: 0, r1: 19, c1: 4 }]
    setup.titleRows = [0, 1]
    setup.rowBreaks = [10, 15]
    wb.transact([{ kind: "pageSetup", sheet: 0, setup }], { label: "Page Setup" })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 5, count: 2 }], { label: "Insert" })
    let p = wb.sheets[0].pageSetup
    expect(p.printArea).toEqual([{ r0: 0, c0: 0, r1: 21, c1: 4 }])
    expect(p.titleRows).toEqual([0, 1])
    expect(p.rowBreaks).toEqual([12, 17])
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 0, count: 1 }], { label: "Delete" })
    p = wb.sheets[0].pageSetup
    expect(p.printArea).toEqual([{ r0: 0, c0: 0, r1: 20, c1: 4 }])
    expect(p.titleRows).toEqual([0, 0])
    expect(p.rowBreaks).toEqual([11, 16])
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].pageSetup.rowBreaks).toEqual([10, 15])
    wb.undo()
    expect(wb.sheets[0].pageSetup.printArea).toBeNull()
  })

  it("loses what is deleted", () => {
    const p = defaultPageSetup()
    p.printArea = [
      { r0: 2, c0: 2, r1: 4, c1: 6 },
      { r0: 0, c0: 8, r1: 3, c1: 9 },
    ]
    p.titleCols = [2, 3]
    p.colBreaks = [3, 7]
    const next = shiftPageSetup(p, "cols", 2, -2)
    expect(next.printArea).toEqual([
      { r0: 2, c0: 2, r1: 4, c1: 4 },
      { r0: 0, c0: 6, r1: 3, c1: 7 },
    ])
    expect(next.titleCols).toBeNull()
    expect(next.colBreaks).toEqual([5])
    expect(p.colBreaks).toEqual([3, 7])
  })
})
