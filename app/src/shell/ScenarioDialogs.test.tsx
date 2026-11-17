import { fireEvent, render, screen, within } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { ScenarioManagerDialog, cellsText, parseCells } from "./ScenarioDialogs"

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

const ok = () => fireEvent.click(screen.getAllByRole("button", { name: "OK" }).at(-1)!)

function addScenario(name: string, values: string[]) {
  fireEvent.click(screen.getByRole("button", { name: "Add..." }))
  fireEvent.change(screen.getByLabelText("Scenario name:"), { target: { value: name } })
  ok()
  const fields = screen.getByText("Enter values for each of the changing cells.").parentElement!
  within(fields)
    .getAllByRole("textbox")
    .forEach((box, i) => fireEvent.change(box, { target: { value: values[i] } }))
  ok()
}

describe("Scenario Manager", () => {
  it("adds scenarios from the selection, shows one, and undoes", () => {
    // B1 rate, B2 amount, B3 = B1 * B2.
    const c = setup([
      [0, 1, "0.05"],
      [1, 1, "1000"],
      [2, 1, "=B1*B2"],
    ])
    c.goTo({ r0: 0, c0: 1, r1: 1, c1: 1 })
    render(<ScenarioManagerDialog controller={c} onClose={() => undefined} />)
    expect(screen.getByText("No Scenarios defined. Choose Add to add scenarios.")).toBeInTheDocument()
    addScenario("Low", ["0.03", "500"])
    addScenario("High", ["0.1", "2000"])
    expect(c.sheet.scenarios.scenarios.map((s) => s.name)).toEqual(["Low", "High"])
    expect(c.sheet.scenarios.scenarios[0].cells).toEqual([
      { row: 0, col: 1, value: 0.03 },
      { row: 1, col: 1, value: 500 },
    ])
    expect(screen.getByText("$B$1:$B$2")).toBeInTheDocument()
    fireEvent.click(screen.getByRole("option", { name: "Low" }))
    fireEvent.click(screen.getByRole("button", { name: "Show" }))
    c.wb.calculate(true)
    expect(c.sheet.cells.value(2, 1)).toBe(15)
    c.undo()
    c.wb.calculate(true)
    expect(c.sheet.cells.value(2, 1)).toBe(50)
  })

  it("refuses a taken name and edits and deletes", () => {
    const c = setup([[0, 0, "1"]])
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    render(<ScenarioManagerDialog controller={c} onClose={() => undefined} />)
    addScenario("One", ["5"])
    fireEvent.click(screen.getByRole("button", { name: "Add..." }))
    fireEvent.change(screen.getByLabelText("Scenario name:"), { target: { value: "one" } })
    ok()
    expect(alert).toHaveBeenCalledWith("Scenario name already exists. Please type a different name.")
    fireEvent.click(screen.getByRole("button", { name: "Cancel" }))
    fireEvent.click(screen.getByRole("button", { name: "Edit..." }))
    expect((screen.getByLabelText("Comment:") as HTMLTextAreaElement).value).toMatch(/\nModified by /)
    ok()
    fireEvent.change(screen.getByLabelText("$A$1"), { target: { value: "7" } })
    ok()
    expect(c.sheet.scenarios.scenarios[0].cells[0].value).toBe(7)
    fireEvent.click(screen.getByRole("button", { name: "Delete" }))
    expect(c.sheet.scenarios.scenarios).toEqual([])
  })

  it("makes a Scenario Summary sheet before the active one", () => {
    const c = setup([
      [0, 0, "2"],
      [1, 0, "=A1*10"],
    ])
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    const close = vi.fn()
    render(<ScenarioManagerDialog controller={c} onClose={close} />)
    addScenario("Three", ["3"])
    fireEvent.click(screen.getByRole("button", { name: "Summary..." }))
    fireEvent.change(screen.getByLabelText("Result cells:"), { target: { value: "A2" } })
    ok()
    expect(close).toHaveBeenCalled()
    expect(c.wb.sheets.map((s) => s.name)).toEqual(["Scenario Summary", "Sheet1"])
    expect(c.sheetIndex).toBe(0)
    const s = c.wb.sheets[0]
    expect([s.cells.value(1, 1), s.cells.value(1, 4), s.cells.value(5, 2), s.cells.value(5, 4)]).toEqual([
      "Scenario Summary",
      "Three",
      "$A$2",
      30,
    ])
    expect(c.wb.sheets[1].scenarios.resultCells).toEqual([{ r0: 1, c0: 0, r1: 1, c1: 0 }])
  })

  it("reads and writes changing cells as Excel shows them", () => {
    const c = new SheetController()
    expect(parseCells(c, "B1:B3, $D$4")).toEqual([
      { row: 0, col: 1 },
      { row: 1, col: 1 },
      { row: 2, col: 1 },
      { row: 3, col: 3 },
    ])
    expect(parseCells(c, "Other!A1")).toBeNull()
    expect(parseCells(c, "A1:A40")).toHaveLength(33)
    expect(
      cellsText([
        { row: 0, col: 1 },
        { row: 1, col: 1 },
        { row: 3, col: 3 },
      ]),
    ).toBe("$B$1:$B$2,$D$4")
  })
})
