import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { afterEach, describe, expect, it, vi } from "vitest"
import type { Platform } from "@cells/platform"
import { Workbook, type Selection } from "@cells/core"
import { PlatformContext } from "@/platform/context"
import { useSettings } from "@/settings/settings"
import { DEFAULT_SETTINGS } from "@/settings/shared"
import { SheetController } from "@/sheet/controller"
import { MemoryAi } from "@/test/memory-ai"
import { RecommendedPivotsDialog } from "./RecommendedPivotsDialog"

afterEach(() => {
  cleanup()
  useSettings.setState({ ...DEFAULT_SETTINGS })
})

/** The dialog with AI on (Ollama) and a test AI that answers `reply`. */
function renderWithAi(c: SheetController, reply: string, onClose = () => undefined) {
  const ai = new MemoryAi()
  ai.reply = reply
  useSettings.setState({ aiEnabled: true, aiProvider: "ollama" })
  render(
    <PlatformContext.Provider value={{ ai } as unknown as Platform}>
      <RecommendedPivotsDialog controller={c} onClose={onClose} />
    </PlatformContext.Provider>,
  )
  return ai
}

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

  it("with AI on, a description puts the AI's PivotTable first, picked, and OK makes it sorted biggest first", async () => {
    const c = setup(SALES)
    const onClose = vi.fn()
    const ai = renderWithAi(
      c,
      '{"rows": ["Region"], "values": [{"field": "Units", "summary": "sum"}], "sort": {"field": "Region", "by": "Units", "order": "descending"}}',
      onClose,
    )
    expect(screen.getByText(/Sends the column names and a few values from each to Ollama/)).toBeTruthy()
    const box = screen.getByLabelText("Describe the PivotTable you want")
    fireEvent.change(box, { target: { value: "units per region, most first" } })
    // Enter asks; it doesn't press OK.
    fireEvent.keyDown(box, { key: "Enter" })
    expect(onClose).not.toHaveBeenCalled()
    await waitFor(() => expect(screen.getAllByRole("option")).toHaveLength(4))
    expect(ai.asked[0].prompt).toBe("units per region, most first")
    expect(ai.asked[0].context).not.toContain("Shop, 4")
    const first = screen.getAllByRole("option")[0]
    expect(first).toHaveAttribute("aria-selected", "true")
    expect(first.querySelector(".recommend-title")?.textContent).toBe("Sum of Units by Region")
    expect(first.querySelector("[aria-label='Described']")).toBeTruthy()
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    expect(onClose).toHaveBeenCalled()
    const [table] = c.sheet.pivots
    expect(table.fields[0]).toMatchObject({ sort: "descending", sortBy: 0 })
    // North 9, South 2.
    expect([c.sheet.cells.get(3, 0)?.value, c.sheet.cells.get(4, 0)?.value]).toEqual(["North", "South"])
  })

  it("shows the AI's problem instead of a PivotTable, and offers nothing to describe with AI off", async () => {
    const c = setup(SALES)
    renderWithAi(c, '{"rows": ["Country"]}')
    fireEvent.change(screen.getByLabelText("Describe the PivotTable you want"), { target: { value: "by country" } })
    fireEvent.click(screen.getByRole("button", { name: "Ask" }))
    expect((await screen.findByRole("alert")).textContent).toMatch(/doesn't have: Country/)
    expect(screen.getAllByRole("option")).toHaveLength(3)
    cleanup()
    useSettings.setState({ aiEnabled: false })
    render(<RecommendedPivotsDialog controller={setup(SALES)} onClose={() => undefined} />)
    expect(screen.queryByLabelText("Describe the PivotTable you want")).toBeNull()
  })
})
