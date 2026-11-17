import { describe, expect, it } from "vitest"
import { OFFICE_THEME, type SheetSlicer } from "@cells/core"
import { objectSvg } from "@/drawing/render"
import { SheetController } from "./controller"

const ROWS = [
  ["Region", "Product", "Sales"],
  ["North", "Apples", 120],
  ["South", "Pears", 80],
  ["North", "Pears", 95],
  ["East", "Apples", 60],
  ["West", "Plums", 40],
]

function book() {
  const c = new SheetController()
  ROWS.forEach((row, r) => row.forEach((v, col) => c.sheet.cells.set(r, col, { value: v })))
  c.tables.create({ r0: 0, c0: 0, r1: ROWS.length - 1, c1: 2 }, true, "TableStyleMedium2")
  const table = c.sheet.tables[0]
  const slicer = (column: string, id: number): SheetSlicer => ({
    kind: "slicer",
    id,
    name: column,
    anchor: { from: { col: 4, row: 1, x: 0, y: 0 }, to: { col: 7, row: 13, x: 0, y: 0 } },
    table: table.id,
    column,
    caption: column,
    header: true,
    style: "SlicerStyleLight1",
    columns: 1,
    buttonHeight: 27,
    sort: "ascending",
    customLists: true,
    crossFilter: "showItemsWithDataAtTop",
    placement: "fixed",
  })
  c.sheet.objects.push(slicer("Region", 100), slicer("Product", 101))
  return { c, region: c.sheet.objects[0] as SheetSlicer, product: c.sheet.objects[1] as SheetSlicer }
}
const hidden = (c: SheetController) => [1, 2, 3, 4, 5].filter((r) => c.sheet.rows.isHidden(r))
const picked = (c: SheetController, o: SheetSlicer) =>
  c.slicers
    .view(o)
    .items.filter((i) => i.selected)
    .map((i) => i.text)

describe("slicers", () => {
  it("draws a slicer's caption and a button per value", () => {
    const { region } = book()
    const svg = objectSvg(region, 192, 243, OFFICE_THEME)
    for (const text of ["Region", "East", "North", "South", "West"]) expect(svg).toContain(`>${text}<`)
  })

  it("filters the table with a click, Ctrl+click and Shift+click, and clears it", () => {
    const { c, region, product } = book()
    c.slicers.pick(region, "north", "only")
    expect(picked(c, region)).toEqual(["North"])
    expect(hidden(c)).toEqual([2, 4, 5])
    expect(c.sheet.tables[0].autoFilter?.columns[0]).toEqual({ kind: "values", values: ["North"] })
    // Products the Region filter leaves no rows for come last, without data.
    expect(c.slicers.view(product).items.map((i) => [i.text, i.data])).toEqual([
      ["Apples", true],
      ["Pears", true],
      ["Plums", false],
    ])
    c.slicers.pick(region, "south", "toggle")
    expect(picked(c, region)).toEqual(["North", "South"])
    c.slicers.pick(region, "east", "only")
    c.slicers.pick(region, "south", "range")
    expect(picked(c, region)).toEqual(["East", "North", "South"])
    expect(c.slicers.view(region).filtered).toBe(true)
    c.slicers.clear(region)
    expect(hidden(c)).toEqual([])
    expect(c.slicers.view(region).filtered).toBe(false)
    c.undo()
    expect(picked(c, region)).toEqual(["East", "North", "South"])
  })

  it("keeps one button picked, and Multi-Select turns clicks into toggles", () => {
    const { c, region } = book()
    c.slicers.toggleMulti(region)
    c.slicers.pick(region, "west", "only")
    expect(picked(c, region)).toEqual(["East", "North", "South"])
    c.slicers.pick(region, "north", "only")
    c.slicers.pick(region, "south", "only")
    c.slicers.pick(region, "east", "only")
    expect(picked(c, region)).toEqual(["East"])
  })
})

describe("slicers and their table", () => {
  it("follow a renamed header, in the same undo step", () => {
    const { c, region } = book()
    c.goTo({ r0: 0, c0: 0, r1: 0, c1: 0 })
    c.beginEdit("enter", "Area")
    c.commitEdit()
    expect((c.sheet.objects[0] as SheetSlicer).column).toBe("Area")
    expect(c.slicers.view(c.sheet.objects[0] as SheetSlicer).items).toHaveLength(4)
    c.undo()
    expect((c.sheet.objects[0] as SheetSlicer).column).toBe(region.column)
  })
})
