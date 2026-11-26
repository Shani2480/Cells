import { describe, expect, it, vi } from "vitest"
import { writeInputOp } from "@cells/core"
import { SheetController } from "./controller"
import { NO_FLASH_PATTERN } from "./data"

function sheet(rows: string[][]) {
  const c = new SheetController()
  c.wb.transact(
    rows.flatMap((r, row) => r.map((t, col) => writeInputOp(0, c.sheet, c.wb.styles, row, col, t))),
    { label: "Setup" },
  )
  c.shell.alert = vi.fn()
  return c
}

describe("Flash Fill (Ctrl+E)", () => {
  it("fills the column from the typed example, under a header, in one undo step", () => {
    const c = sheet([
      ["Full name", "First"],
      ["Nancy Davolio", "Nancy"],
      ["Andrew Fuller", ""],
      ["Janet Leverling", ""],
    ])
    c.goTo({ r0: 2, c0: 1, r1: 2, c1: 1 })
    expect(c.keyReady({ key: "e", primary: true, shift: false, alt: false })).toBe(true)
    expect([2, 3].map((r) => c.sheet.cells.get(r, 1)?.value)).toEqual(["Andrew", "Janet"])
    expect(c.sheet.cells.get(0, 1)?.value).toBe("First")
    c.wb.undo()
    expect(c.sheet.cells.get(2, 1)).toBeUndefined()
  })

  it("numbers are read as shown, and the results are entered as if typed (Excel 365: Order 0042 gives 42)", () => {
    const c = sheet([
      ["Order 1001", "1001"],
      ["Order 1002", ""],
      ["Order 0042", ""],
    ])
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    c.data.flashFill()
    expect(c.sheet.cells.get(1, 1)?.value).toBe(1002)
    expect(c.sheet.cells.get(2, 1)?.value).toBe(42)
  })

  it("says so when there is no pattern", () => {
    const c = sheet([
      ["Nancy", ""],
      ["Andrew", ""],
    ])
    c.goTo({ r0: 0, c0: 1, r1: 0, c1: 1 })
    expect(c.data.flashFill()).toBe(0)
    expect(c.shell.alert).toHaveBeenCalledWith(NO_FLASH_PATTERN)
  })
})
