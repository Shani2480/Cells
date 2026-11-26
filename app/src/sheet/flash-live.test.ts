import { writeInputOp } from "@cells/core"
import { afterEach, describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"
import { setAutoFlashFill } from "./data"

function sheet(rows: string[][]) {
  const c = new SheetController()
  c.wb.transact(
    rows.flatMap((r, row) => r.map((t, col) => writeInputOp(0, c.sheet, c.wb.styles, row, col, t))),
    { label: "Setup" },
  )
  c.shell.alert = vi.fn()
  return c
}

const names = () =>
  sheet([
    ["Full name", "First"],
    ["Nancy Davolio", "Nancy"],
    ["Andrew Fuller", ""],
    ["Janet Leverling", ""],
    ["Margaret Peacock", ""],
  ])

const key = (c: SheetController, k: string, shift = false) =>
  c.keyEditing({ key: k, primary: false, shift, alt: false })

/** Starts typing `text` in a cell, as the grid's editor does. */
function type(c: SheetController, row: number, col: number, text: string) {
  c.goTo({ r0: row, c0: col, r1: row, c1: col })
  c.beginEdit("enter", text)
}

const column = (c: SheetController, col: number, rows: number[]) => rows.map((r) => c.sheet.cells.get(r, col)?.value)

afterEach(() => setAutoFlashFill(true))

describe("live Flash Fill (Automatically Flash Fill)", () => {
  it("offers the rest of the column while the next entry is typed; Enter takes it, one step to undo", () => {
    const c = names()
    type(c, 2, 1, "An")
    expect(c.data.liveFlash()).toEqual({
      row: 2,
      col: 1,
      value: "Andrew",
      fills: [
        { row: 3, text: "Janet" },
        { row: 4, text: "Margaret" },
      ],
    })
    key(c, "Enter")
    expect(column(c, 1, [1, 2, 3, 4])).toEqual(["Nancy", "Andrew", "Janet", "Margaret"])
    expect(c.selection.active).toEqual({ row: 3, col: 1 })
    expect(c.wb.undoLabel).toBe("Flash Fill")
    c.undo()
    expect(column(c, 1, [2, 3, 4])).toEqual([undefined, undefined, undefined])
  })

  it("follows what's typed: no offer once it stops fitting the pattern", () => {
    const c = names()
    type(c, 2, 1, "A")
    expect(c.data.liveFlash()?.value).toBe("Andrew")
    c.setEditText("Ax")
    expect(c.data.liveFlash()).toBeNull()
    key(c, "Enter")
    expect(column(c, 1, [2, 3])).toEqual(["Ax", undefined])
  })

  it("Tab and Esc leave the offer; a formula, a cell with a value and the switch turned off get none", () => {
    const c = names()
    type(c, 2, 1, "An")
    key(c, "Tab")
    expect(column(c, 1, [2, 3])).toEqual(["An", undefined])
    type(c, 3, 1, "J")
    key(c, "Escape")
    expect(column(c, 1, [3])).toEqual([undefined])
    type(c, 3, 1, "=A4")
    expect(c.data.liveFlash()).toBeNull()
    c.cancelEdit()
    type(c, 2, 1, "A")
    expect(c.data.liveFlash()).toBeNull()
    c.cancelEdit()
    setAutoFlashFill(false)
    type(c, 3, 1, "J")
    expect(c.data.liveFlash()).toBeNull()
  })

  it("needs an example: the first entry in a column gets no offer", () => {
    const c = sheet([
      ["Nancy Davolio", ""],
      ["Andrew Fuller", ""],
    ])
    type(c, 0, 1, "Nancy")
    expect(c.data.liveFlash()).toBeNull()
  })
})
