import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { QuickGallery } from "./QuickAnalysis"

function book() {
  const c = new SheetController()
  const rows: (string | number)[][] = [
    ["Item", "Jan", "Feb"],
    ["Pens", 4, 6],
    ["Ink", 7, 2],
  ]
  rows.forEach((row, r) => row.forEach((v, col) => c.sheet.cells.set(r, col, { value: v })))
  c.goTo({ r0: 0, c0: 0, r1: 2, c1: 2 })
  return c
}

describe("Quick Analysis gallery", () => {
  it("previews the choice under the pointer, makes it on click and closes", () => {
    const c = book()
    const onClose = vi.fn()
    render(<QuickGallery controller={c} left={10} top={10} onClose={onClose} />)
    expect(screen.getByRole("tab", { name: "Formatting" })).toHaveAttribute("aria-selected", "true")
    const bars = screen.getByRole("button", { name: "Data Bars" })
    fireEvent.mouseEnter(bars)
    expect(c.conditional.preview?.rule.kind).toBe("dataBar")
    expect(screen.getByText(/A bar in each cell/)).toBeInTheDocument()
    fireEvent.mouseLeave(bars)
    expect(c.conditional.preview).toBeNull()
    fireEvent.click(bars)
    expect(onClose).toHaveBeenCalled()
    expect(c.sheet.conditional).toHaveLength(1)
  })

  it("switches tabs, previews totals on the sheet and takes them back when it closes", () => {
    const c = book()
    const { unmount } = render(<QuickGallery controller={c} left={10} top={10} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("tab", { name: "Totals" }))
    const sums = screen.getAllByRole("button", { name: "Sum" })
    expect(sums).toHaveLength(2)
    fireEvent.mouseEnter(sums[0])
    expect(c.sheet.cells.get(3, 1)?.input).toBe("=SUM(B2:B3)")
    unmount()
    expect(c.sheet.cells.get(3, 1)).toBeUndefined()
    expect(c.wb.canUndo).toBe(false)
  })

  it("offers recommended PivotTables, makes one on a new sheet, and Esc closes", () => {
    const c = new SheetController()
    const rows: (string | number)[][] = [
      ["Region", "Units"],
      ["North", 4],
      ["South", 2],
      ["North", 5],
    ]
    rows.forEach((row, r) => row.forEach((v, col) => c.sheet.cells.set(r, col, { value: v })))
    c.goTo({ r0: 0, c0: 0, r1: 3, c1: 1 })
    const onClose = vi.fn()
    const { unmount } = render(<QuickGallery controller={c} left={10} top={10} onClose={onClose} />)
    fireEvent.click(screen.getByRole("tab", { name: "Tables" }))
    const pivot = screen.getByRole("button", { name: "PivotTable" })
    fireEvent.mouseEnter(pivot)
    expect(screen.getByText("Sum of Units by Region. PivotTables summarize data, on a new sheet.")).toBeTruthy()
    expect(screen.getByRole("button", { name: "More" })).toBeTruthy()
    fireEvent.keyDown(window, { key: "Escape" })
    expect(onClose).toHaveBeenCalled()
    unmount()
    c.quick.apply("pivot:0")
    expect(c.wb.sheets).toHaveLength(2)
    const [table] = c.sheet.pivots
    expect(table.values.map((v) => v.name)).toEqual(["Sum of Units"])
    expect(c.sheet.cells.get(3, 0)?.value).toBe("North")
    expect(c.sheet.cells.get(3, 1)?.value).toBe(9)
  })

  it("has no PivotTables to recommend when every row is different, only More", () => {
    const c = book()
    render(<QuickGallery controller={c} left={10} top={10} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("tab", { name: "Tables" }))
    expect(screen.queryByRole("button", { name: "PivotTable" })).toBeNull()
    expect(screen.getByRole("button", { name: "More" })).toBeTruthy()
  })

  it("moves along the tabs with the arrow keys and down into the choices", () => {
    const c = book()
    render(<QuickGallery controller={c} left={10} top={10} onClose={() => undefined} />)
    const first = screen.getByRole("tab", { name: "Formatting" })
    first.focus()
    fireEvent.keyDown(first, { key: "ArrowRight" })
    expect(screen.getByRole("tab", { name: "Charts" })).toHaveAttribute("aria-selected", "true")
    fireEvent.keyDown(document.activeElement!, { key: "ArrowDown" })
    expect(document.activeElement).toHaveAccessibleName("Clustered Column")
  })
})
