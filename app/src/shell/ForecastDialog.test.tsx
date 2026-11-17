import { fireEvent, render, screen } from "@testing-library/react"
import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { ForecastDialog } from "./ForecastDialog"

/** Months as numbers 1–24 and sales with a season of four and a rising trend; one value missing. */
function setup() {
  const c = new SheetController()
  const put = (row: number, col: number, text: string) => {
    c.goTo({ r0: row, c0: col, r1: row, c1: col })
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  put(0, 0, "Month")
  put(0, 1, "Sales")
  const season = [5, -3, 8, -10]
  for (let i = 0; i < 24; i++) {
    put(i + 1, 0, String(i + 1))
    if (i !== 10) put(i + 1, 1, String(100 + 2 * i + season[i % 4]))
  }
  c.goTo({ r0: 3, c0: 1, r1: 3, c1: 1 })
  c.wb.calculate(true)
  return c
}

describe("Forecast Sheet", () => {
  it("previews the forecast and creates the sheet with its table, formulas and chart; undo takes it away", () => {
    const c = setup()
    render(<ForecastDialog controller={c} onClose={() => undefined} />)
    expect(screen.getByLabelText("Forecast End")).toHaveValue("32")
    expect(screen.getByLabelText("Preview").querySelector("svg")).not.toBeNull()
    fireEvent.click(screen.getByRole("button", { name: "Options" }))
    expect(screen.getByLabelText("Timeline Range")).toHaveValue("=Sheet1!$A$1:$A$25")
    expect(screen.getByLabelText("Values Range")).toHaveValue("=Sheet1!$B$1:$B$25")
    fireEvent.click(screen.getByLabelText("Include forecast statistics"))
    c.columnFit = (col) => 100 + col
    fireEvent.click(screen.getByRole("button", { name: "Create" }))
    expect([0, 1, 2].map((col) => c.sheet.cols.size(col))).toEqual([100, 101, 102])
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Sheet2", "Sheet1"])
    expect(c.sheetIndex).toBe(0)
    const s = c.sheet
    c.wb.calculate(true)
    expect([0, 1, 2, 3, 4].map((col) => s.cells.value(0, col))).toEqual([
      "Month",
      "Sales",
      "Forecast(Sales)",
      "Lower Confidence Bound(Sales)",
      "Upper Confidence Bound(Sales)",
    ])
    // The last point of the history starts the forecast lines; then 8 months of formulas.
    expect(s.cells.value(24, 2)).toBe(s.cells.value(24, 1))
    expect(s.cells.input(25, 2)).toBe("=FORECAST.ETS(A26,$B$2:$B$25,$A$2:$A$25,1,1)")
    expect(s.cells.input(25, 3)).toBe("=C26-FORECAST.ETS.CONFINT(A26,$B$2:$B$25,$A$2:$A$25,0.95,1,1)")
    expect(s.cells.value(32, 0)).toBe(32)
    const f = s.cells.value(25, 2) as number
    expect(Math.abs(f - (100 + 2 * 24 + 5))).toBeLessThan(3)
    expect((s.cells.value(25, 4) as number) > f && (s.cells.value(25, 3) as number) < f).toBe(true)
    expect(s.cells.value(11, 1)).toBeUndefined()
    expect(s.tables[0].range).toEqual({ r0: 0, c0: 0, r1: 32, c1: 4 })
    expect(s.charts).toHaveLength(1)
    expect(s.cells.value(1, 6)).toBe("Alpha")
    expect(s.cells.input(1, 7)).toBe("=FORECAST.ETS.STAT($B$2:$B$25,$A$2:$A$25,1,1,1)")
    c.undo()
    expect(c.wb.sheets.map((x) => x.name)).toEqual(["Sheet1"])
  })

  it("starts earlier when asked, forecasting the rest of the history from what came before", () => {
    const c = setup()
    render(<ForecastDialog controller={c} onClose={() => undefined} />)
    fireEvent.click(screen.getByRole("button", { name: "Options" }))
    fireEvent.change(screen.getByLabelText("Forecast Start"), { target: { value: "20" } })
    fireEvent.click(screen.getByLabelText("Confidence Interval"))
    fireEvent.click(screen.getByRole("button", { name: "Create" }))
    const s = c.sheet
    expect(s.cells.value(0, 3)).toBeUndefined()
    expect(s.cells.value(20, 2)).toBe(s.cells.value(20, 1))
    expect(s.cells.input(21, 2)).toBe("=FORECAST.ETS(A22,$B$2:$B$21,$A$2:$A$21,1,1)")
    expect(s.cells.value(19, 2)).toBeUndefined()
  })
})
