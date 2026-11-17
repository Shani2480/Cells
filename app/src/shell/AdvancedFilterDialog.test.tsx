import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { AdvancedFilterDialog } from "./AdvancedFilterDialog"

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

const list: [number, number, string][] = [
  [0, 0, "Name"],
  [0, 1, "Sales"],
  [1, 0, "Ann"],
  [1, 1, "10"],
  [2, 0, "Bob"],
  [2, 1, "50"],
  [3, 0, "Cy"],
  [3, 1, "70"],
  // Criteria in E1:E2.
  [0, 4, "Sales"],
  [1, 4, ">20"],
]

describe("Advanced Filter", () => {
  it("filters in place from the list around the active cell; Clear shows the rows again", () => {
    const c = setup(list)
    c.goTo({ r0: 1, c0: 0, r1: 1, c1: 0 })
    render(<AdvancedFilterDialog controller={c} onClose={() => undefined} />)
    expect(screen.getByLabelText("List range:")).toHaveValue("$A$1:$B$4")
    fireEvent.change(screen.getByLabelText("Criteria range:"), { target: { value: "E1:E2" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect([1, 2, 3].map((r) => c.sheet.rows.isHidden(r))).toEqual([true, false, false])
    expect(c.data.clearable).toBe(true)
    c.data.clearFilters()
    expect([1, 2, 3].map((r) => c.sheet.rows.isHidden(r))).toEqual([false, false, false])
    expect(c.data.clearable).toBe(false)
    c.undo()
    expect(c.sheet.rows.isHidden(1)).toBe(true)
  })

  it("copies to another location on the sheet, and refuses another sheet", () => {
    const c = setup(list)
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    render(<AdvancedFilterDialog controller={c} onClose={() => undefined} />)
    fireEvent.click(screen.getByLabelText("Copy to another location"))
    fireEvent.change(screen.getByLabelText("Criteria range:"), { target: { value: "$E$1:$E$2" } })
    fireEvent.change(screen.getByLabelText("Copy to:"), { target: { value: "Other!A1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(alert).toHaveBeenCalledWith("Reference isn't valid.")
    fireEvent.change(screen.getByLabelText("Copy to:"), { target: { value: "G1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    const col = (c0: number) => [0, 1, 2, 3].map((r) => c.sheet.cells.value(r, c0))
    expect(col(6)).toEqual(["Name", "Bob", "Cy", undefined])
    expect(col(7)).toEqual(["Sales", 50, 70, undefined])
    expect(c.sheet.rows.isHidden(1)).toBe(false)
  })
})
