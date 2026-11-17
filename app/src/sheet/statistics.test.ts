import { describe, expect, it } from "vitest"
import { SheetController } from "./controller"
import { workbookStatistics } from "./statistics"

describe("Workbook Statistics", () => {
  it("counts cells, formulas and notes for the sheet and the whole workbook", () => {
    const c = new SheetController()
    c.sheets.insert(1)
    c.wb.transact(
      [
        {
          kind: "setCells",
          sheet: 0,
          cells: [
            { row: 0, col: 0, data: { value: 1 } },
            { row: 4, col: 2, data: { value: 2, input: "=1+1" } },
          ],
        },
        { kind: "setCells", sheet: 1, cells: [{ row: 0, col: 0, data: { value: "x" } }] },
      ],
      { label: "x", before: c.selection },
    )
    c.wb.sheets[1].annotations.set(0, 0, { kind: "note", author: "Alex", text: "Check" })
    const s = workbookStatistics(c.wb, 0)
    expect(s.sheet).toMatchObject({ name: "Sheet1", end: "C5", cells: 2, formulas: 1, notes: 0 })
    expect(s.workbook).toMatchObject({ sheets: 2, cells: 3, formulas: 1, notes: 1, comments: 0 })
  })
})
