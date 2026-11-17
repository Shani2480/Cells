import { describe, expect, it } from "vitest"
import { parseRange } from "./address"
import { canEditCell, hashPassword, protectionBlocks } from "./protection"
import { Workbook } from "./workbook"

const protectAll = {
  selectLocked: true,
  selectUnlocked: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  editObjects: false,
  editScenarios: false,
}

describe("Allow Edit Ranges", () => {
  it("opens a range's cells on a protected sheet; a password keeps them shut until unlocked", () => {
    const wb = new Workbook(["Data"])
    const sheet = wb.sheets[0]
    wb.transact(
      [
        {
          kind: "editRanges",
          sheet: 0,
          ranges: [
            { name: "Inputs", ranges: [parseRange("B2:B5")!] },
            { name: "Rates", ranges: [parseRange("D1")!], password: hashPassword("pw") },
          ],
        },
      ],
      { label: "x" },
    )
    sheet.protection = protectAll
    expect(canEditCell(sheet, wb.styles, 2, 1)).toBe(true)
    expect(canEditCell(sheet, wb.styles, 0, 0)).toBe(false)
    expect(canEditCell(sheet, wb.styles, 0, 3)).toBe(false)
    sheet.unlockedEditRanges.add("rates")
    expect(canEditCell(sheet, wb.styles, 0, 3)).toBe(true)
    const write = (row: number, col: number) =>
      protectionBlocks(sheet, wb.styles, { kind: "setCells", sheet: 0, cells: [{ row, col, data: { value: 1 } }] })
    expect(write(3, 1)).toBe(false)
    expect(write(0, 0)).toBe(true)
    // Formatting still follows Protect Sheet's own options.
    expect(protectionBlocks(sheet, wb.styles, { kind: "format", sheet: 0, ranges: [], patch: {} })).toBe(true)
  })

  it("follows inserted rows and goes back on undo", () => {
    const wb = new Workbook(["Data"])
    wb.transact([{ kind: "editRanges", sheet: 0, ranges: [{ name: "R", ranges: [parseRange("A5:A6")!] }] }], {
      label: "x",
    })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 2 }], { label: "y" })
    expect(wb.sheets[0].editRanges[0].ranges[0]).toEqual(parseRange("A7:A8"))
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].editRanges).toEqual([])
  })
})
