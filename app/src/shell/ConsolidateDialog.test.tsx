import { fireEvent, render, screen, within } from "@testing-library/react"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "@/sheet/controller"
import { ConsolidateDialog } from "./ConsolidateDialog"

function setup() {
  const c = new SheetController()
  c.sheets.insert(1)
  c.sheets.insert(2)
  const fill = (sheet: number, cells: [number, number, string][]) => {
    c.activate(sheet)
    for (const [row, col, text] of cells) {
      c.goTo({ r0: row, c0: col, r1: row, c1: col })
      c.beginEdit("enter", text)
      c.commitEdit()
    }
  }
  fill(1, [
    [0, 0, "Item"],
    [0, 1, "Qty"],
    [1, 0, "Pens"],
    [1, 1, "2"],
    [2, 0, "Ink"],
    [2, 1, "5"],
  ])
  fill(2, [
    [0, 0, "Item"],
    [0, 1, "Qty"],
    [1, 0, "Ink"],
    [1, 1, "10"],
  ])
  c.activate(0)
  c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
  c.wb.calculate(true)
  return c
}

describe("Consolidate", () => {
  it("adds references, consolidates by labels into the selection and remembers the settings", () => {
    const c = setup()
    const [s2, s3] = [c.wb.sheets[1].name, c.wb.sheets[2].name]
    const { unmount } = render(<ConsolidateDialog controller={c} onClose={() => undefined} />)
    const ref = screen.getByLabelText("Reference:")
    fireEvent.change(ref, { target: { value: `${s3}!A1:B2` } })
    fireEvent.click(screen.getByRole("button", { name: "Add" }))
    fireEvent.change(ref, { target: { value: `${s2}!a1:b3` } })
    fireEvent.click(screen.getByRole("button", { name: "Add" }))
    expect(
      within(screen.getByRole("listbox"))
        .getAllByRole("option")
        .map((o) => o.textContent),
    ).toEqual([`${s2}!$A$1:$B$3`, `${s3}!$A$1:$B$2`])
    fireEvent.click(screen.getByLabelText("Top row"))
    fireEvent.click(screen.getByLabelText("Left column"))
    fireEvent.click(screen.getByRole("button", { name: "OK" }))
    const v = (r: number, col: number) => c.sheet.cells.value(r, col)
    expect([v(0, 1), v(1, 0), v(1, 1), v(2, 0), v(2, 1)]).toEqual(["Qty", "Pens", 2, "Ink", 15])
    expect(c.sheet.consolidate).toMatchObject({ fn: 0, top: true, left: true, links: false })
    unmount()
    render(<ConsolidateDialog controller={c} onClose={() => undefined} />)
    expect(within(screen.getByRole("listbox")).getAllByRole("option")).toHaveLength(2)
    expect(screen.getByLabelText("Top row")).toBeChecked()
    c.undo()
    expect(v(2, 1)).toBeUndefined()
    expect(c.sheet.consolidate).toBeNull()
  })

  it("says when a reference isn't valid", () => {
    const c = setup()
    const alert = vi.fn()
    c.shell = { ...c.shell, alert }
    render(<ConsolidateDialog controller={c} onClose={() => undefined} />)
    fireEvent.change(screen.getByLabelText("Reference:"), { target: { value: "Nowhere!A1:B2" } })
    fireEvent.click(screen.getByRole("button", { name: "Add" }))
    expect(alert).toHaveBeenCalledWith("Reference isn't valid.")
  })
})
