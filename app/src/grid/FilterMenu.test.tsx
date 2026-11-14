import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { CustomFilterDialog, rowFor, Top10Dialog } from "@/shell/FilterDialogs"
import { FilterMenu } from "./FilterMenu"

function table() {
  const c = new SheetController()
  const rows: (string | number)[][] = [
    ["Name", "Team", "Score"],
    ["Ann", "Blue", 5],
    ["Bob", "Red", 7],
    ["Cy", "Blue", 9],
    ["Dee", "", 2],
  ]
  rows.forEach((row, r) => row.forEach((value, col) => value !== "" && c.sheet.cells.set(r, col, { value })))
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  c.data.toggleFilter()
  return c
}
const hidden = (c: SheetController) => [1, 2, 3, 4].map((r) => c.sheet.rows.isHidden(r))

function open(c: SheetController, col: number, onDialog = vi.fn()) {
  const onClose = vi.fn()
  const view = render(<FilterMenu controller={c} col={col} left={10} top={10} onClose={onClose} onDialog={onDialog} />)
  return { onClose, onDialog, view }
}

describe("filter dropdown", () => {
  it("lists the values with (Blanks); unticking one and OK hides its rows", () => {
    const c = table()
    const { onClose } = open(c, 1)
    const boxes = screen.getAllByRole("checkbox").map((b) => b.closest("label")!.textContent)
    expect(boxes).toEqual(["(Select All)", "Blue", "Red", "(Blanks)"])
    fireEvent.click(screen.getByRole("checkbox", { name: "Red" }))
    fireEvent.click(screen.getByRole("checkbox", { name: "(Blanks)" }))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.sheet.autoFilter?.columns[1]).toEqual({ kind: "values", values: ["Blue"] })
    expect(hidden(c)).toEqual([false, true, false, true])
  })

  it("Select All off disables OK; the search filters to its matches", () => {
    const c = table()
    open(c, 0)
    fireEvent.click(screen.getByRole("checkbox", { name: "(Select All)" }))
    expect(screen.getByRole("button", { name: "OK" })).toBeDisabled()
    fireEvent.change(screen.getByRole("searchbox", { name: "Search" }), { target: { value: "e" } })
    expect(screen.getByRole("checkbox", { name: "Dee" })).toBeChecked()
    expect(screen.queryByRole("checkbox", { name: "Ann" })).toBeNull()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(hidden(c)).toEqual([true, true, true, false])
  })

  it("a number column offers Number Filters; Top 10 and Greater Than open their dialogs", () => {
    const c = table()
    const { onDialog } = open(c, 2)
    expect(screen.getByRole("menuitem", { name: "Sort Smallest to Largest" })).toBeInTheDocument()
    fireEvent.click(screen.getByRole("menuitem", { name: "Number Filters" }))
    fireEvent.click(screen.getByRole("menuitemcheckbox", { name: "Greater Than..." }))
    expect(onDialog).toHaveBeenCalledWith({ kind: "custom", col: 2, preset: { first: "greater" } })
  })

  it("Above Average applies at once; Clear Filter From takes it off", () => {
    const c = table()
    const first = open(c, 2)
    fireEvent.click(screen.getByRole("menuitem", { name: "Number Filters" }))
    fireEvent.click(screen.getByRole("menuitemcheckbox", { name: "Above Average" }))
    expect(hidden(c)).toEqual([true, false, false, true])
    first.view.unmount()
    open(c, 2)
    fireEvent.click(screen.getByRole("menuitem", { name: 'Clear Filter From "Score"' }))
    expect(hidden(c)).toEqual([false, false, false, false])
  })

  it("Sort Z to A sorts the filtered data", () => {
    const c = table()
    open(c, 0)
    fireEvent.click(screen.getByRole("menuitem", { name: "Sort Z to A" }))
    expect([1, 2, 3, 4].map((r) => c.sheet.cells.get(r, 0)?.value)).toEqual(["Dee", "Cy", "Bob", "Ann"])
  })
})

describe("Custom AutoFilter and Top 10", () => {
  it("begins with … or … ends with", () => {
    const c = table()
    render(<CustomFilterDialog controller={c} col={0} preset={{ first: "beginsWith" }} onClose={() => {}} />)
    fireEvent.change(screen.getByRole("combobox", { name: "First value" }), { target: { value: "a" } })
    fireEvent.click(screen.getByRole("radio", { name: "Or" }))
    fireEvent.change(screen.getByRole("combobox", { name: "Second condition" }), { target: { value: "endsWith" } })
    fireEvent.change(screen.getByRole("combobox", { name: "Second value" }), { target: { value: "y" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(c.sheet.autoFilter?.columns[0]).toEqual({
      kind: "custom",
      and: false,
      conditions: [
        { op: "eq", value: "a*" },
        { op: "eq", value: "*y" },
      ],
    })
    expect(hidden(c)).toEqual([false, true, false, true])
  })

  it("reads saved conditions back as the dialog's rows", () => {
    expect(rowFor({ op: "ne", value: "*x*" })).toEqual({ op: "notContains", value: "x" })
    expect(rowFor({ op: "eq", value: "x*" })).toEqual({ op: "beginsWith", value: "x" })
    expect(rowFor({ op: "eq", value: "*" })).toEqual({ op: "equals", value: "*" })
    expect(rowFor({ op: "le", value: "4" })).toEqual({ op: "lessOrEqual", value: "4" })
  })

  it("Top 2 items", () => {
    const c = table()
    render(<Top10Dialog controller={c} col={2} onClose={() => {}} />)
    fireEvent.change(screen.getByRole("spinbutton", { name: "How many" }), { target: { value: "2" } })
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(hidden(c)).toEqual([true, false, false, true])
  })
})
