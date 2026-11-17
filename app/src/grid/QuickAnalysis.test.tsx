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

  it("shows PivotTable greyed out and Esc closes", () => {
    const c = book()
    const onClose = vi.fn()
    render(<QuickGallery controller={c} left={10} top={10} onClose={onClose} />)
    fireEvent.click(screen.getByRole("tab", { name: "Tables" }))
    expect(screen.getByRole("button", { name: "PivotTable" })).toHaveAttribute("aria-disabled", "true")
    fireEvent.click(screen.getByRole("button", { name: "PivotTable" }))
    expect(onClose).not.toHaveBeenCalled()
    fireEvent.keyDown(window, { key: "Escape" })
    expect(onClose).toHaveBeenCalled()
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
