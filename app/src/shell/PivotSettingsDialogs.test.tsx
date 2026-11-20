import { cleanup, fireEvent, render, screen } from "@testing-library/react"
import { afterEach, describe, expect, it, vi } from "vitest"
import { withFieldIn, Workbook, type Selection } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { FieldSettingsDialog, PivotOptionsDialog, ValueFieldSettingsDialog } from "./PivotSettingsDialogs"

afterEach(cleanup)

/** Units by Region (rows) and Channel (columns) at F1. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  c.shell.openDialog = vi.fn()
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Units"],
    ["North", "Shop", 4],
    ["South", "Online", 2],
    ["North", "Online", 5],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.pivots.create("A1:C4", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
  c.pivots.change((t, k) => withFieldIn(withFieldIn(withFieldIn(t, k, 0, "rows"), k, 1, "cols"), k, 2, "values"))
  return c
}

const at = (c: SheetController, row: number, col: number) => c.sheet.cells.get(row, col)?.value

describe("PivotTable settings", () => {
  it("knows which value field or field a cell belongs to, and a header's double-click opens it", () => {
    const c = setup()
    // F1 Sum of Units | Column Labels; F2 Row Labels | Online | Shop | Grand Total; F3 North ...
    expect(at(c, 0, 5)).toBe("Sum of Units")
    expect(c.pivots.settingsTargetAt(0, 5)).toEqual({ kind: "value", index: 0 })
    expect(c.pivots.settingsTargetAt(2, 6)).toEqual({ kind: "value", index: 0 })
    expect(c.pivots.settingsTargetAt(2, 5)).toEqual({ kind: "field", field: 0 })
    expect(c.pivots.settingsTargetAt(1, 6)).toEqual({ kind: "field", field: 1 })
    expect(c.pivots.openSettingsAt(0, 5)).toBe(true)
    expect(c.shell.openDialog).toHaveBeenCalledWith("pivotValue:0")
    expect(c.pivots.openSettingsAt(2, 6)).toBe(false)
  })

  it("Value Field Settings changes the function (renaming the automatic name) and the number format", () => {
    const c = setup()
    const onClose = vi.fn()
    render(<ValueFieldSettingsDialog controller={c} index={0} onClose={onClose} />)
    fireEvent.click(screen.getByRole("option", { name: "Max" }))
    expect(screen.getByLabelText("Custom Name:")).toHaveProperty("value", "Max of Units")
    fireEvent.click(screen.getByRole("button", { name: "Number Format" }))
    expect(screen.getByRole("dialog", { name: "Format Cells" })).toBeTruthy()
    fireEvent.click(screen.getAllByRole("button", { name: "Cancel" })[1])
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(at(c, 0, 5)).toBe("Max of Units")
    expect(c.sheet.pivots[0].values[0]).toMatchObject({ fn: "max" })
    c.undo()
    expect(at(c, 0, 5)).toBe("Sum of Units")
  })

  it("Value Field Settings says when the name is taken", () => {
    const c = setup()
    render(<ValueFieldSettingsDialog controller={c} index={0} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText("Custom Name:"), { target: { value: "Region" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.shell.alert).toHaveBeenCalledWith("PivotTable field name already exists.")
  })

  it("Field Settings renames a field and turns its subtotals off", () => {
    const c = setup()
    render(<FieldSettingsDialog controller={c} field={1} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText("Custom Name:"), { target: { value: "Where" } })
    fireEvent.click(screen.getByRole("radio", { name: "None" }))
    fireEvent.click(screen.getByRole("tab", { name: "Layout & Print" }))
    fireEvent.click(screen.getByRole("radio", { name: "Show item labels in tabular form" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.pivots[0].fields[1]).toMatchObject({ name: "Where", subtotals: [], outline: false })
  })

  it("PivotTable Options renames the table, hides the row grand totals and shows a caption for empty cells", () => {
    const c = setup()
    render(<PivotOptionsDialog controller={c} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText("PivotTable Name:"), { target: { value: "Sales" } })
    fireEvent.change(screen.getByLabelText("Empty cell caption"), { target: { value: "-" } })
    fireEvent.click(screen.getByRole("tab", { name: "Totals & Filters" }))
    fireEvent.click(screen.getByRole("checkbox", { name: "Show grand totals for rows" }))
    fireEvent.click(screen.getByRole("tab", { name: "Data" }))
    fireEvent.click(screen.getByRole("checkbox", { name: "Refresh data when opening the file" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    const t = c.sheet.pivots[0]
    expect(t).toMatchObject({ name: "Sales", colGrandTotals: false, captions: { missing: "-" } })
    expect(c.wb.pivotCaches[0].refreshOnLoad).toBe(true)
    // Columns Shop, Online and no Grand Total; South had no Shop sales, so its cell shows the caption.
    expect([5, 6, 7, 8].map((col) => at(c, 1, col))).toEqual(["Row Labels", "Shop", "Online", undefined])
    expect([5, 6, 7].map((col) => at(c, 3, col))).toEqual(["South", "-", 2])
    c.undo()
    expect(c.sheet.pivots[0].name).toBe("PivotTable1")
    expect(c.wb.pivotCaches[0].refreshOnLoad).toBe(false)
  })
})
