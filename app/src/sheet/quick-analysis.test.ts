import { describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"
import { previewing, REPLACE_QUESTION } from "./quick-analysis"

/** Region, Q1, Q2, Q3 with three rows of numbers under a header. */
function book() {
  const c = new SheetController()
  const rows: (string | number)[][] = [
    ["Region", "Q1", "Q2", "Q3"],
    ["North", 10, 20, 30],
    ["South", 5, 15, 25],
    ["East", 1, 2, 3],
  ]
  rows.forEach((row, r) => row.forEach((v, col) => c.sheet.cells.set(r, col, { value: v })))
  c.shell.alert = vi.fn()
  c.goTo({ r0: 0, c0: 0, r1: 3, c1: 3 })
  return c
}

const input = (c: SheetController, row: number, col: number) => c.sheet.cells.get(row, col)?.input
const value = (c: SheetController, row: number, col: number) => c.sheet.cells.get(row, col)?.value

describe("Quick Analysis", () => {
  it("shows for a block of cells with data, not for one cell or empty cells", () => {
    const c = book()
    expect(c.quick.available()).toBe(true)
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    expect(c.quick.available()).toBe(false)
    c.goTo({ r0: 10, c0: 10, r1: 12, c1: 12 })
    expect(c.quick.available()).toBe(false)
  })

  it("Ctrl+Q on one cell selects the data around it first", () => {
    const c = book()
    c.goTo({ r0: 2, c0: 2, r1: 2, c1: 2 })
    const open = vi.fn()
    c.quick.openGallery = open
    expect(c.quick.show()).toBe(true)
    expect(c.selection.ranges[0]).toEqual({ r0: 0, c0: 0, r1: 3, c1: 3 })
    expect(open).toHaveBeenCalled()
  })

  it("reads the header, the columns of numbers and the kind of data", () => {
    const shape = book().quick.shape()!
    expect(shape).toMatchObject({ header: true, top: 1, numberCols: [1, 2, 3], kind: "numbers", positive: true })
  })

  it("offers Excel's formatting for numbers and for text", () => {
    const c = book()
    expect(c.quick.formatting(c.quick.shape()!).map((i) => i.label)).toEqual([
      "Data Bars",
      "Color Scale",
      "Icon Set",
      "Greater Than",
      "Top 10%",
      "Clear Format",
    ])
    c.goTo({ r0: 0, c0: 0, r1: 3, c1: 0 })
    expect(c.quick.formatting(c.quick.shape()!).map((i) => i.label)).toEqual([
      "Text Contains",
      "Duplicate Values",
      "Unique Values",
      "Equal To",
      "Clear Format",
    ])
  })

  it("previews a format without an undo step and makes it with one", () => {
    const c = book()
    c.quick.preview("cf:dataBar")
    expect(c.conditional.preview?.rule.kind).toBe("dataBar")
    expect(c.sheet.conditional).toEqual([])
    c.quick.apply("cf:dataBar")
    expect(c.conditional.preview).toBeNull()
    expect(c.sheet.conditional.map((e) => e.rule.kind)).toEqual(["dataBar"])
    expect(c.wb.undoLabel).toBe("Conditional Formatting")
  })

  it("writes totals in the row below: sums, % of the whole, running totals", () => {
    const c = book()
    c.quick.apply("total:below:sum")
    expect([1, 2, 3].map((col) => input(c, 4, col))).toEqual(["=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"])
    expect(value(c, 4, 1)).toBe(16)
    expect(value(c, 4, 0)).toBe("Sum")
    expect(c.wb.styles.get(c.sheet.styleAt(4, 1)).bold).toBe(true)
    c.undo()
    expect(c.sheet.cells.get(4, 1)).toBeUndefined()

    c.quick.apply("total:below:percent")
    expect(input(c, 4, 2)).toBe("=SUM(C2:C4)/SUM($B$2:$D$4)")
    expect(c.wb.styles.get(c.sheet.styleAt(4, 2)).numberFormat).toBe("0.00%")
    c.undo()
    c.quick.apply("total:below:running")
    expect([1, 2, 3].map((col) => input(c, 4, col))).toEqual(["=SUM($B$2:B4)", "=SUM($B$2:C4)", "=SUM($B$2:D4)"])
    expect(value(c, 4, 3)).toBe(111)
  })

  it("writes totals in the column to the right, with the header", () => {
    const c = book()
    c.quick.apply("total:right:average")
    expect(value(c, 0, 4)).toBe("Average")
    expect([1, 2, 3].map((row) => input(c, row, 4))).toEqual(["=AVERAGE(B2:D2)", "=AVERAGE(B3:D3)", "=AVERAGE(B4:D4)"])
    expect(value(c, 1, 4)).toBe(20)
    c.undo()
    c.quick.apply("total:right:running")
    expect(input(c, 3, 4)).toBe("=SUM(B$2:D4)")
  })

  it("gives totals the format of the cell before them", () => {
    const c = book()
    const blue = c.wb.styles.intern({ fill: "#1F4E78" })
    const money = c.wb.styles.intern({ numberFormat: "$#,##0.00" })
    c.sheet.cells.set(0, 3, { value: "Q3", style: blue })
    c.sheet.cells.set(3, 2, { value: 2, style: money })
    c.quick.apply("total:right:sum")
    expect(c.wb.styles.get(c.sheet.styleAt(0, 4))).toMatchObject({ fill: "#1F4E78", bold: true })
    c.undo()
    c.quick.apply("total:below:sum")
    expect(c.wb.styles.get(c.sheet.styleAt(4, 2))).toMatchObject({ numberFormat: "$#,##0.00", bold: true })
  })

  it("asks before writing totals over cells that hold something", () => {
    const c = book()
    c.sheet.cells.set(4, 2, { value: "note" })
    const confirm = vi.fn()
    c.shell.confirm = confirm
    c.quick.apply("total:below:sum")
    expect(confirm).toHaveBeenCalledWith(REPLACE_QUESTION, expect.any(Function))
    expect(value(c, 4, 2)).toBe("note")
    confirm.mock.calls[0][1]()
    expect(input(c, 4, 2)).toBe("=SUM(C2:C4)")
  })

  it("previews totals, a table and sparklines on the sheet and takes them back, outside undo", () => {
    const c = book()
    const counted: boolean[] = []
    c.wb.subscribe(() => counted.push(previewing.has(c.wb)))
    for (const id of ["total:below:sum", "table", "spark:line"]) {
      c.quick.preview(id)
      expect(c.quick.previewShown).toBe(true)
      c.quick.endPreview()
    }
    expect(c.sheet.cells.get(4, 1)).toBeUndefined()
    expect(c.sheet.tables).toEqual([])
    expect(c.sheet.sparklines).toEqual([])
    expect(c.wb.canUndo).toBe(false)
    // Every change the previews made was marked, so the document doesn't count it as an edit.
    expect(counted.length).toBe(6)
    expect(counted.every(Boolean)).toBe(true)
  })

  it("a choice made after its preview is one undo step", () => {
    const c = book()
    c.quick.preview("table")
    c.quick.apply("table")
    expect(c.sheet.tables).toHaveLength(1)
    expect(c.sheet.tables[0].range).toEqual({ r0: 0, c0: 0, r1: 3, c1: 3 })
    c.undo()
    expect(c.sheet.tables).toEqual([])
    expect(c.wb.canUndo).toBe(false)
  })

  it("makes sparklines of each row's numbers in the column to the right", () => {
    const c = book()
    c.quick.apply("spark:column")
    const g = c.sheet.sparklines[0]
    expect(g.type).toBe("column")
    expect(g.sparklines.map((s) => [s.row, s.col, s.source])).toEqual([
      [1, 4, "Sheet1!B2:D2"],
      [2, 4, "Sheet1!B3:D3"],
      [3, 4, "Sheet1!B4:D4"],
    ])
  })

  it("recommends charts for the data and draws them from it", () => {
    const c = book()
    const shape = c.quick.shape()!
    expect(c.quick.charts(shape)).toEqual(["column", "column-stacked", "hbar", "line", "area-stacked"])
    const spec = c.quick.chartSpec("line")!
    expect(spec.series.map((s) => s.name)).toEqual(["North", "South", "East"])
    expect(spec.categories).toEqual(["Q1", "Q2", "Q3"])
    c.goTo({ r0: 0, c0: 0, r1: 3, c1: 1 })
    expect(c.quick.charts(c.quick.shape()!)).toContain("pie")
  })

  it("leaves Table out for data already in a table; More opens Recommended PivotTables", () => {
    const c = book()
    c.shell.openDialog = vi.fn()
    c.quick.apply("table")
    const items = c.quick.tables(c.quick.shape()!)
    expect(items[0]).toMatchObject({ id: "table", disabled: "The data is already in a table." })
    expect(items.at(-1)?.id).toBe("more:pivot")
    c.quick.apply("more:pivot")
    expect(c.shell.openDialog).toHaveBeenCalledWith(expect.stringMatching(/^recommendPivots:A1:/))
  })
})
