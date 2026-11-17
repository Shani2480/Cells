import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { DataTableDialog, GoalSeekDialog } from "./WhatIfDialogs"

function setup(cells: [number, number, string][]) {
  const c = new SheetController()
  for (const [row, col, text] of cells) {
    c.goTo({ r0: row, c0: col, r1: row, c1: col })
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  c.wb.calculate(true)
  return c
}

describe("Goal Seek", () => {
  it("shows the status, keeps the answer on OK as one undo step", () => {
    const c = setup([
      [0, 0, "3"],
      [0, 1, "=A1*4+2"],
    ])
    c.goTo({ r0: 0, c0: 1, r1: 0, c1: 1 })
    const close = vi.fn()
    render(<GoalSeekDialog controller={c} onClose={close} />)
    expect(screen.getByLabelText("Set cell:")).toHaveValue("B1")
    fireEvent.change(screen.getByLabelText("To value:"), { target: { value: "30" } })
    fireEvent.change(screen.getByLabelText("By changing cell:"), { target: { value: "$A$1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(screen.getByText("Goal Seeking with Cell B1 found a solution.")).toBeInTheDocument()
    expect(c.sheet.cells.value(0, 0)).toBe(3)
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.cells.value(0, 0)).toBeCloseTo(7, 6)
    c.wb.calculate(true)
    expect(c.sheet.cells.value(0, 1)).toBeCloseTo(30, 3)
    c.undo()
    expect(c.sheet.cells.value(0, 0)).toBe(3)
  })

  it("refuses a set cell without a formula", () => {
    const c = setup([[0, 0, "3"]])
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    render(<GoalSeekDialog controller={c} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText("To value:"), { target: { value: "1" } })
    fireEvent.change(screen.getByLabelText("By changing cell:"), { target: { value: "A2" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(alert).toHaveBeenCalledWith("Cell must contain a formula.")
  })
})

describe("Data Table", () => {
  it("fills the selection with TABLE, shows braces in the formula bar and refuses a part edit", () => {
    // A1 rate, A2 = A1*100; D1 = A2; C2:C3 rates to try.
    const c = setup([
      [0, 0, "0.1"],
      [1, 0, "=A1*100"],
      [0, 3, "=A2"],
      [1, 2, "0.2"],
      [2, 2, "0.3"],
    ])
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    c.goTo({ r0: 0, c0: 2, r1: 2, c1: 3 })
    const close = vi.fn()
    render(<DataTableDialog controller={c} onClose={close} />)
    fireEvent.change(screen.getByLabelText("Column input cell:"), { target: { value: "$A$1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(close).toHaveBeenCalled()
    c.wb.calculate(true)
    expect([c.sheet.cells.value(1, 3), c.sheet.cells.value(2, 3)]).toEqual([20, 30])
    c.goTo({ r0: 1, c0: 3, r1: 1, c1: 3 })
    expect(c.formulaBarText()).toBe("{=TABLE(,A1)}")
    c.beginEdit("enter", "5")
    expect(alert).toHaveBeenCalledWith("Cannot change part of a data table.")
    expect(c.edit).toBeNull()
    c.clearContents()
    expect(c.sheet.cells.input(1, 3)).toBe("=TABLE(,A1)")
    c.goTo({ r0: 1, c0: 3, r1: 2, c1: 3 })
    c.clearContents()
    expect(c.sheet.cells.input(1, 3)).toBeUndefined()
  })

  it("needs a selection of two rows and two columns, and an input cell", () => {
    const c = setup([[0, 0, "1"]])
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    c.goTo({ r0: 0, c0: 0, r1: 3, c1: 0 })
    render(<DataTableDialog controller={c} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText("Row input cell:"), { target: { value: "B1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(alert).toHaveBeenCalledWith("Selection not valid.")
  })
})
