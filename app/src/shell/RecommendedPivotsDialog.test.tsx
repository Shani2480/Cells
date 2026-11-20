import { cleanup, fireEvent, render, screen } from "@testing-library/react"
import { afterEach, describe, expect, it, vi } from "vitest"
import { Workbook, type Selection } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { RecommendedPivotsDialog } from "./RecommendedPivotsDialog"

afterEach(cleanup)

function setup(rows: (string | number)[][]) {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  return c
}

const SALES = [
  ["Region", "Channel", "Units"],
  ["North", "Shop", 4],
  ["South", "Online", 2],
  ["North", "Online", 5],
]

describe("Recommended PivotTables", () => {
  it("lists the suggestions with pictures, shows the one picked large, and OK makes it on a new sheet", () => {
    const c = setup(SALES)
    const onClose = vi.fn()
    render(<RecommendedPivotsDialog controller={c} onClose={onClose} />)
    const options = screen.getAllByRole("option")
    expect(options.map((o) => o.querySelector(".recommend-title")?.textContent)).toEqual([
      "Sum of Units by Region",
      "Sum of Units by Channel",
      "Sum of Units by Channel and Region",
    ])
    expect(options[0]).toHaveAttribute("aria-selected", "true")
    fireEvent.click(options[2])
    expect(document.querySelector(".recommend-preview-title")?.textContent).toBe("Sum of Units by Channel and Region")
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet2", "Data"])
    const [table] = c.sheet.pivots
    expect([table.rows, table.cols]).toEqual([[1], [0]])
  })

  it("Blank PivotTable makes an empty one; no data says what to select", () => {
    const c = setup(SALES)
    render(<RecommendedPivotsDialog controller={c} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "Blank PivotTable" }))
    expect(c.sheet.pivots[0].rows).toEqual([])
    cleanup()
    const empty = setup([])
    render(<RecommendedPivotsDialog controller={empty} onClose={() => undefined} />)
    expect(screen.getByText(/to see recommended PivotTables/)).toBeTruthy()
  })
})
