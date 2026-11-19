import { act, cleanup, fireEvent, render, screen, within } from "@testing-library/react"
import { afterEach, describe, expect, it, vi } from "vitest"
import { Workbook, type Selection } from "@cells/core"
import { SheetController } from "@/sheet/controller"
import { PivotFieldsPane } from "./PivotFieldsPane"

afterEach(cleanup)

function open() {
  const c = new SheetController(new Workbook<Selection>(["Data"]))
  c.shell.alert = vi.fn()
  const rows: (string | number)[][] = [
    ["Region", "Channel", "Units"],
    ["North", "Shop", 4],
    ["South", "Online", 2],
    ["North", "Online", 5],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => c.sheet.cells.set(row, col, { value })))
  c.pivots.create("A1:C4", { kind: "existing", sheet: 0, at: { row: 0, col: 5 } })
  const view = render(<PivotFieldsPane controller={c} onDone={() => undefined} />)
  const rerender = () => view.rerender(<PivotFieldsPane controller={c} onDone={() => undefined} />)
  c.subscribe?.(rerender)
  const area = (name: string) => screen.getByRole("region", { name })
  const chipsIn = (name: string) =>
    within(area(name))
      .queryAllByRole("listitem")
      .map((li) => li.querySelector(".pivot-chip-name")?.textContent)
  const tick = (name: string) => {
    act(() => void fireEvent.click(screen.getByRole("checkbox", { name })))
    rerender()
  }
  return { c, area, chipsIn, tick, rerender }
}

/** A drag from `from` dropped on `to`, with one DataTransfer between them. */
function dragTo(from: Element, to: Element) {
  const store = new Map<string, string>()
  const dataTransfer = {
    setData: (k: string, v: string) => void store.set(k, v),
    getData: (k: string) => store.get(k) ?? "",
    effectAllowed: "",
    dropEffect: "",
  }
  act(() => {
    fireEvent.dragStart(from, { dataTransfer })
    fireEvent.dragOver(to, { dataTransfer })
    fireEvent.drop(to, { dataTransfer })
  })
}

describe("PivotTable Fields pane", () => {
  it("ticking a field puts text in Rows and numbers in Values; unticking takes it out", () => {
    const { c, chipsIn, tick } = open()
    tick("Region")
    tick("Units")
    expect(chipsIn("Rows")).toEqual(["Region"])
    expect(chipsIn("Values")).toEqual(["Sum of Units"])
    expect(c.sheet.cells.get(1, 5)?.value).toBe("North")
    expect(screen.getByRole("checkbox", { name: "Region" })).toHaveProperty("checked", true)
    tick("Region")
    expect(chipsIn("Rows")).toEqual([])
    expect(c.sheet.cells.get(0, 5)?.value).toBe("Sum of Units")
  })

  it("a chip's menu moves it to another area, and up and down", () => {
    const { chipsIn, tick, rerender } = open()
    tick("Region")
    tick("Channel")
    expect(chipsIn("Rows")).toEqual(["Region", "Channel"])
    act(() => void fireEvent.click(screen.getByRole("button", { name: "Channel options" })))
    rerender()
    act(() => void fireEvent.click(screen.getByRole("menuitem", { name: "Move Up" })))
    rerender()
    expect(chipsIn("Rows")).toEqual(["Channel", "Region"])
    act(() => void fireEvent.click(screen.getByRole("button", { name: "Region options" })))
    rerender()
    act(() => void fireEvent.click(screen.getByRole("menuitem", { name: "Move to Column Labels" })))
    rerender()
    expect(chipsIn("Rows")).toEqual(["Channel"])
    expect(chipsIn("Columns")).toEqual(["Region"])
    act(() => void fireEvent.click(screen.getByRole("button", { name: "Region options" })))
    rerender()
    act(() => void fireEvent.click(screen.getByRole("menuitem", { name: "Remove Field" })))
    rerender()
    expect(chipsIn("Columns")).toEqual([])
  })

  it("dragging a field from the list into an area, and a chip between areas", () => {
    const { c, area, chipsIn, rerender } = open()
    const item = (name: string) => screen.getByRole("checkbox", { name }).closest("li")!
    dragTo(item("Channel"), area("Columns"))
    rerender()
    dragTo(item("Units"), area("Values"))
    rerender()
    dragTo(item("Units"), area("Values"))
    rerender()
    expect(chipsIn("Columns")).toEqual(["Channel", "Σ Values"])
    expect(chipsIn("Values")).toEqual(["Sum of Units", "Sum of Units2"])
    const chip = within(area("Columns")).getAllByRole("listitem")[1]
    dragTo(chip, area("Rows"))
    rerender()
    expect(chipsIn("Rows")).toEqual(["Σ Values"])
    // A value chip dropped on the list leaves Values.
    dragTo(within(area("Values")).getAllByRole("listitem")[1], screen.getByRole("list", { name: "Fields" }))
    rerender()
    expect(chipsIn("Values")).toEqual(["Sum of Units"])
    expect(chipsIn("Rows")).toEqual([])
    expect(c.sheet.pivots[0].values).toHaveLength(1)
  })

  it("the search box narrows the list", () => {
    open()
    fireEvent.change(screen.getByRole("searchbox", { name: "Search" }), { target: { value: "ni" } })
    expect(screen.getAllByRole("checkbox").map((b) => b.closest("label")?.textContent)).toEqual(["Units"])
  })
})
