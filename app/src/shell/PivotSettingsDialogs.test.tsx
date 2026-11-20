import { cleanup, fireEvent, render, screen } from "@testing-library/react"
import { afterEach, describe, expect, it, vi } from "vitest"
import { withFieldIn, Workbook, type Selection } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import {
  CalculatedFieldDialog,
  FieldSettingsDialog,
  PivotOptionsDialog,
  ShowValuesAsDialog,
  ValueFieldSettingsDialog,
} from "./PivotSettingsDialogs"

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

  it("Show Values As: % of Column Total, then % Of the item North of Region", () => {
    const c = setup()
    render(<ValueFieldSettingsDialog controller={c} index={0} onClose={vi.fn()} />)
    fireEvent.click(screen.getByRole("tab", { name: "Show Values As" }))
    expect(screen.getByRole("listbox", { name: "Base field" })).toHaveAttribute("aria-disabled", "true")
    fireEvent.change(screen.getByLabelText("Show values as"), { target: { value: "percentOfCol" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    // F3 North: Shop 4 of 4, Online 5 of 7.
    expect([at(c, 2, 6), at(c, 2, 7)]).toEqual([1, 5 / 7])
    expect(c.sheet.pivots[0].values[0].numFmt).toBe("0.00%")
    cleanup()
    render(<ShowValuesAsDialog controller={c} index={0} as="percent" onClose={vi.fn()} />)
    expect(screen.getByRole("dialog", { name: "Show Values As (Sum of Units)" })).toBeTruthy()
    fireEvent.change(screen.getByLabelText("Base Item:"), { target: { value: "0" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    // South Online 2 against North's 5; South has no Shop units: #NULL!.
    expect(at(c, 3, 6)).toEqual({ error: "#NULL!" })
    expect(at(c, 3, 7)).toBe(2 / 5)
    expect(c.sheet.pivots[0].values[0]).toMatchObject({ showAs: "percent", baseField: 0, baseItem: 0 })
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

describe("Insert Calculated Field", () => {
  it("inserts fields into the formula, adds with Add and OK, and shows Excel's message", () => {
    const c = setup()
    c.goTo({ r0: 2, c0: 6, r1: 2, c1: 6 })
    const onClose = vi.fn()
    render(<CalculatedFieldDialog controller={c} onClose={onClose} />)
    const name = screen.getByLabelText("Name:") as HTMLInputElement
    const formula = screen.getByLabelText("Formula:") as HTMLInputElement
    expect(name.value).toBe("Field1")
    expect(formula.value).toBe("= 0")
    fireEvent.change(formula, { target: { value: "=" } })
    formula.setSelectionRange(1, 1)
    fireEvent.click(screen.getByRole("option", { name: "Units" }))
    fireEvent.click(screen.getByRole("button", { name: "Insert Field" }))
    expect(formula.value).toBe("=Units")
    fireEvent.change(formula, { target: { value: "=Units*A1" } })
    fireEvent.click(screen.getByRole("button", { name: "Add" }))
    expect(c.shell.alert).toHaveBeenCalledWith(
      "References, names, and arrays are not supported in PivotTable formulas.",
    )
    fireEvent.change(name, { target: { value: "Double" } })
    fireEvent.change(formula, { target: { value: "=Units*2" } })
    fireEvent.click(screen.getByRole("button", { name: "Add" }))
    expect(c.pivots.calcFields()).toEqual([{ name: "Double", formula: "Units*2" }])
    expect(onClose).not.toHaveBeenCalled()
    // The name is taken now: Modify and Delete.
    expect(screen.getByRole("button", { name: "Modify" })).toBeTruthy()
    fireEvent.click(screen.getByRole("button", { name: "Delete" }))
    expect(c.pivots.calcFields()).toEqual([])
    fireEvent.change(name, { target: { value: "Triple" } })
    fireEvent.change(formula, { target: { value: "=Units*3" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.sheet.pivots[0].values.map((v) => v.name)).toEqual(["Sum of Units", "Sum of Triple"])
  })
})
