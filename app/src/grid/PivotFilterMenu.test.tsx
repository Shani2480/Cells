import { cleanup, fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import {
  newPivotField,
  pivotCacheFrom,
  pivotCellsPlan,
  readPivotSource,
  Workbook,
  type PivotTable,
  type Selection,
} from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { PivotFilterDialog } from "@/shell/PivotFilterDialogs"
import { PivotFilterMenu } from "./PivotFilterMenu"

/** Sales by Region and Product (compact rows) on Sheet2 at A3, with Channel as a report filter in A1:B1. */
function setup() {
  const c = new SheetController(new Workbook<Selection>(["Sheet1", "Sheet2"]))
  c.shell.alert = vi.fn()
  const rows: (string | number)[][] = [
    ["Region", "Product", "Channel", "Sales"],
    ["North", "Apples", "Shop", 10],
    ["North", "Pears", "Online", 5],
    ["South", "Apples", "Shop", 20],
    ["East", "Plums", "Online", 7],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.wb.sheets[0].cells.set(row, col, { value })))
  const source = { kind: "range" as const, sheet: "Sheet1", ref: "A1:D5" }
  const read = readPivotSource(c.wb.sheets, c.wb.styles, source)
  if (typeof read === "string") throw new Error(read)
  const cache = pivotCacheFrom(1, source, read)
  const fields = cache.fields.map((f) => newPivotField(f.items?.length ?? 0))
  fields[0].axis = "row"
  fields[1].axis = "row"
  fields[2].axis = "page"
  fields[3].dataField = true
  const table: PivotTable = {
    name: "PivotTable1",
    cacheId: 1,
    at: { row: 3, col: 0 },
    location: { r0: 3, c0: 0, r1: 3, c1: 0 },
    fields,
    rows: [0, 1],
    cols: [],
    pages: [{ field: 2, item: null }],
    values: [{ name: "Sum of Sales", field: 3, fn: "sum" }],
    rowGrandTotals: true,
    colGrandTotals: true,
    compact: true,
    outline: true,
    compactData: true,
    captions: { data: "Values" },
    showError: false,
    showMissing: true,
    style: { name: null, rowHeaders: true, colHeaders: true, rowStripes: false, colStripes: false, lastColumn: true },
  }
  c.wb.pivotCaches = [cache]
  const plan = pivotCellsPlan(c.wb.sheets, c.wb.styles, 1, null, table, cache)
  c.wb.transact([...plan.ops, { kind: "pivots", sheet: 1, tables: [plan.table] }], { label: "Make" })
  c.wb.clearHistory()
  c.activate(1)
  return c
}
const labels = (c: SheetController) =>
  Array.from({ length: 12 }, (_, r) => c.sheet.cells.get(r + 3, 0)?.value).filter((v) => v !== undefined)

function open(c: SheetController, row: number, col: number) {
  const target = c.pivots.filterTargetAt(row, col)!
  const onClose = vi.fn()
  cleanup()
  render(<PivotFilterMenu controller={c} row={row} col={col} target={target} left={10} top={10} onClose={onClose} />)
  return onClose
}

describe("PivotTable filter dropdown", () => {
  it("Row Labels' dropdown picks a field, and unticking an item hides it as one undo step", () => {
    const c = setup()
    expect(labels(c)).toEqual([
      "Row Labels",
      "North",
      "Apples",
      "Pears",
      "South",
      "Apples",
      "East",
      "Plums",
      "Grand Total",
    ])
    expect(c.pivots.filterTargetAt(3, 0)).toEqual({ kind: "row", fields: [0, 1] })
    const onClose = open(c, 3, 0)
    expect(screen.getByRole("combobox")).toHaveValue("0")
    expect(screen.getAllByRole("checkbox").map((b) => b.closest("label")!.textContent)).toEqual([
      "(Select All)",
      "North",
      "South",
      "East",
    ])
    fireEvent.click(screen.getByRole("checkbox", { name: "South" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(labels(c)).toEqual(["Row Labels", "North", "Apples", "Pears", "East", "Plums", "Grand Total"])
    expect(c.sheet.cells.get(9, 1)?.value).toBe(22)
    c.undo()
    expect(labels(c)).toHaveLength(9)
  })

  it("the second field's list, its search, and Sort Z to A", () => {
    const c = setup()
    open(c, 3, 0)
    fireEvent.change(screen.getByRole("combobox"), { target: { value: "1" } })
    expect(screen.getByRole("checkbox", { name: "Plums" })).toBeChecked()
    fireEvent.change(screen.getByRole("searchbox", { name: "Search" }), { target: { value: "ea" } })
    expect(screen.queryByRole("checkbox", { name: "Apples" })).toBeNull()
    expect(screen.queryByRole("checkbox", { name: "Plums" })).toBeNull()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    // Only the match stays: Pears.
    expect(labels(c)).toEqual(["Row Labels", "North", "Pears", "Grand Total"])
    c.undo()
    open(c, 3, 0)
    fireEvent.click(screen.getByRole("menuitemradio", { name: "Sort Z to A" }))
    expect(labels(c)).toEqual([
      "Row Labels",
      "South",
      "Apples",
      "North",
      "Apples",
      "Pears",
      "East",
      "Plums",
      "Grand Total",
    ])
    expect(c.sheet.pivots[0].fields[0].sort).toBe("descending")
  })

  it("Clear Filter shows every item again", () => {
    const c = setup()
    c.pivots.showItems(3, 0, 0, new Set([0]))
    expect(labels(c)).toEqual(["Row Labels", "North", "Apples", "Pears", "Grand Total"])
    open(c, 3, 0)
    fireEvent.click(screen.getByRole("menuitem", { name: 'Clear Filter From "Region"' }))
    expect(labels(c)).toHaveLength(9)
  })

  it("a report filter picks one item, or several with Select Multiple Items", () => {
    const c = setup()
    expect(c.sheet.cells.get(1, 1)?.value).toBe("(All)")
    open(c, 1, 1)
    expect(screen.queryByRole("menuitemradio")).toBeNull()
    fireEvent.click(screen.getByRole("radio", { name: "Online" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.cells.get(1, 1)?.value).toBe("Online")
    expect(labels(c)).toEqual(["Row Labels", "North", "Pears", "East", "Plums", "Grand Total"])
    open(c, 1, 1)
    expect(screen.getByRole("radio", { name: "Online" })).toHaveAttribute("aria-checked", "true")
    fireEvent.click(screen.getByRole("checkbox", { name: "Select Multiple Items" }))
    fireEvent.click(screen.getByRole("checkbox", { name: "Shop" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.cells.get(1, 1)?.value).toBe("(All)")
    expect(c.sheet.pivots[0].fields[2].multipleItems).toBe(true)
    c.pivots.showItems(1, 1, 2, new Set([0]))
    expect(c.sheet.cells.get(1, 1)?.value).toBe("Shop")
    expect(labels(c)).toEqual(["Row Labels", "North", "Apples", "South", "Apples", "Grand Total"])
  })

  it("Alt+Down on Row Labels opens its dropdown", () => {
    const c = setup()
    const openMenu = vi.fn()
    c.pivots.openMenu = openMenu
    c.goTo({ r0: 3, c0: 0, r1: 3, c1: 0 })
    expect(c.pivots.openMenuAtActive()).toBe(true)
    expect(openMenu).toHaveBeenCalledWith(3, 0)
    c.goTo({ r0: 4, c0: 0, r1: 4, c1: 0 })
    expect(c.pivots.openMenuAtActive()).toBe(false)
  })

  it("Label Filters › Begins With... opens its dialog, whose OK filters the field", () => {
    const c = setup()
    c.shell.pivotFilterDialog = vi.fn()
    open(c, 3, 0)
    fireEvent.click(screen.getByRole("menuitem", { name: "Label Filters" }))
    fireEvent.click(screen.getByRole("menuitemcheckbox", { name: "Begins With..." }))
    expect(c.shell.pivotFilterDialog).toHaveBeenCalledWith({
      kind: "label",
      row: 3,
      col: 0,
      field: 0,
      op: "beginsWith",
    })
    cleanup()
    render(
      <PivotFilterDialog
        controller={c}
        request={{ kind: "label", row: 3, col: 0, field: 0, op: "beginsWith" }}
        onClose={vi.fn()}
      />,
    )
    expect(screen.getByRole("dialog", { name: "Label Filter (Region)" })).toBeInTheDocument()
    expect(screen.getByRole("combobox", { name: "Comparison" })).toHaveValue("beginsWith")
    fireEvent.change(screen.getByRole("textbox", { name: "Value" }), { target: { value: "s" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(labels(c)).toEqual(["Row Labels", "South", "Apples", "Grand Total"])
    expect(c.sheet.pivots[0].filters).toEqual([{ field: 0, id: 1, kind: "label", op: "beginsWith", value: "s" }])
    // The button's menu now offers Clear Filter in Label Filters, and a tick on it.
    open(c, 3, 0)
    fireEvent.click(screen.getByRole("menuitem", { name: "Label Filters" }))
    fireEvent.click(screen.getByRole("menuitem", { name: "Clear Filter" }))
    expect(labels(c)).toHaveLength(9)
    c.undo()
    expect(labels(c)).toHaveLength(4)
  })

  it("the Top 10 Filter dialog keeps the biggest items by a value field", () => {
    const c = setup()
    render(<PivotFilterDialog controller={c} request={{ kind: "top", row: 3, col: 0, field: 0 }} onClose={vi.fn()} />)
    expect(screen.getByRole("dialog", { name: "Top 10 Filter (Region)" })).toBeInTheDocument()
    fireEvent.change(screen.getByRole("spinbutton", { name: "How many" }), { target: { value: "1" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(labels(c)).toEqual(["Row Labels", "South", "Apples", "Grand Total"])
    // A Value Filter replaces it: North (15) and South (20) are over 10.
    cleanup()
    render(
      <PivotFilterDialog
        controller={c}
        request={{ kind: "value", row: 3, col: 0, field: 0, op: "greaterThan" }}
        onClose={vi.fn()}
      />,
    )
    fireEvent.change(screen.getByRole("textbox", { name: "Value" }), { target: { value: "10" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(labels(c)).toEqual(["Row Labels", "North", "Apples", "Pears", "South", "Apples", "Grand Total"])
    expect(c.sheet.pivots[0].filters).toHaveLength(1)
  })
})
