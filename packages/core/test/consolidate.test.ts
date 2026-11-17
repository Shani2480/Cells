import { describe, expect, it } from "vitest"
import { combine, consolidateOps, Workbook, type ConsolidateOptions } from "../src"

type V = string | number
function put(wb: Workbook, sheet: number, rows: V[][], at = { row: 0, col: 0 }) {
  const cells = rows.flatMap((r, ri) =>
    r.map((v, ci) => ({ row: at.row + ri, col: at.col + ci, data: { value: v } })).filter((c) => c.data.value !== ""),
  )
  wb.transact([{ kind: "setCells", sheet, cells }], { label: "Load", history: false })
}

const range = (r0: number, c0: number, r1: number, c1: number) => ({ r0, c0, r1, c1 })
const opts = (o: Partial<ConsolidateOptions>): ConsolidateOptions => ({
  fn: 0,
  top: false,
  left: false,
  links: false,
  ...o,
})

function grid(wb: Workbook, sheet: number, r: { r0: number; c0: number; r1: number; c1: number }) {
  const out: unknown[][] = []
  for (let row = r.r0; row <= r.r1; row++) {
    const line: unknown[] = []
    for (let col = r.c0; col <= r.c1; col++) line.push(wb.sheets[sheet].cells.value(row, col) ?? null)
    out.push(line)
  }
  return out
}

function run(
  wb: Workbook,
  sheet: number,
  dest: ReturnType<typeof range>,
  sources: { sheet: number; range: ReturnType<typeof range> }[],
  o: ConsolidateOptions,
) {
  const result = consolidateOps(wb.sheets, wb.styles, sheet, dest, sources, o)
  if ("problem" in result) throw new Error(result.problem)
  wb.transact(result.ops, { label: "Consolidate" })
  wb.calculate(true)
  return result.area
}

describe("Consolidate", () => {
  it("combines values with each function", () => {
    expect(combine(0, [1, 2, "x", 3])).toBe(6)
    expect(combine(1, [1, "x"])).toBe(2) // Count is COUNTA
    expect(combine(6, [1, "x"])).toBe(1) // Count Numbers
    expect(combine(2, [2, 4])).toBe(3)
    expect(combine(5, [2, 4])).toBe(8)
    expect(combine(9, [2, 4])).toBe(2) // Var
    expect(combine(10, [2, 4])).toBe(1) // Varp
    expect(combine(7, [5])).toBeUndefined() // StdDev of one value
    expect(combine(0, ["x"])).toBeUndefined()
  })

  it("by position: the same cells of each range", () => {
    const wb = new Workbook(["East", "West", "Total"])
    put(wb, 0, [
      [1, 2],
      [3, 4],
    ])
    put(wb, 1, [
      [10, 20],
      [30, 40],
      [50, ""],
    ])
    const area = run(
      wb,
      2,
      range(1, 1, 1, 1),
      [
        { sheet: 0, range: range(0, 0, 1, 1) },
        { sheet: 1, range: range(0, 0, 2, 1) },
      ],
      opts({}),
    )
    expect(area).toEqual(range(1, 1, 3, 2))
    expect(grid(wb, 2, area)).toEqual([
      [11, 22],
      [33, 44],
      [50, null],
    ])
  })

  it("by top row and left column labels, matched without case, in the order first seen", () => {
    const wb = new Workbook(["East", "West", "Total"])
    put(wb, 0, [
      ["", "Jan", "Feb"],
      ["Pens", 1, 2],
      ["Ink", 3, 4],
    ])
    put(wb, 1, [
      ["", "feb", "Mar"],
      ["ink", 10, 5],
      ["Paper", 7, 8],
    ])
    const sources = [
      { sheet: 0, range: range(0, 0, 2, 2) },
      { sheet: 1, range: range(0, 0, 2, 2) },
    ]
    const area = run(wb, 2, range(0, 0, 0, 0), sources, opts({ top: true, left: true }))
    expect(grid(wb, 2, area)).toEqual([
      [null, "Jan", "Feb", "Mar"],
      ["Pens", 1, 2, null],
      ["Ink", 3, 14, 5],
      ["Paper", null, 7, 8],
    ])
  })

  it("Average and a destination range that limits the result", () => {
    const wb = new Workbook(["A", "B", "Total"])
    put(wb, 0, [
      ["Item", "Qty"],
      ["x", 2],
      ["y", 4],
    ])
    put(wb, 1, [
      ["Item", "Qty"],
      ["x", 6],
      ["y", 8],
    ])
    const sources = [
      { sheet: 0, range: range(0, 0, 2, 1) },
      { sheet: 1, range: range(0, 0, 2, 1) },
    ]
    const area = run(wb, 2, range(0, 0, 1, 1), sources, opts({ fn: 2, top: true, left: true }))
    expect(area).toEqual(range(0, 0, 1, 1))
    expect(grid(wb, 2, range(0, 0, 2, 1))).toEqual([
      [null, "Qty"],
      ["x", 4],
      [null, null],
    ])
  })

  it("with links: detail rows linking to the sources, grouped and hidden under each summary", () => {
    const wb = new Workbook(["East", "West", "Total"])
    put(wb, 0, [
      ["Item", "Qty"],
      ["Pens", 1],
      ["Ink", 3],
    ])
    put(wb, 1, [
      ["Item", "Qty"],
      ["Ink", 10],
    ])
    const sources = [
      { sheet: 0, range: range(0, 0, 2, 1) },
      { sheet: 1, range: range(0, 0, 1, 1) },
    ]
    const area = run(wb, 2, range(0, 0, 0, 0), sources, opts({ top: true, left: true, links: true }))
    const t = wb.sheets[2]
    expect(area).toEqual(range(0, 0, 5, 2))
    expect(grid(wb, 2, area)).toEqual([
      [null, null, "Qty"],
      [null, "East", 1],
      ["Pens", null, 1],
      [null, "East", 3],
      [null, "West", 10],
      ["Ink", null, 13],
    ])
    expect(t.cells.input(3, 2)).toBe("=East!$B$3")
    expect(t.cells.input(5, 2)).toBe("=SUM(C4:C5)")
    expect([1, 2, 3, 4, 5].map((r) => [t.rows.level(r), t.rows.isHidden(r)])).toEqual([
      [1, true],
      [0, false],
      [1, true],
      [1, true],
      [0, false],
    ])
    expect(t.rows.isCollapsed(5)).toBe(true)
    wb.undo()
    expect(t.rows.isHidden(1)).toBe(false)
    expect(t.cells.value(5, 2)).toBeUndefined()
  })

  it("refuses links to the same sheet and sources that overlap the result", () => {
    const wb = new Workbook(["S"])
    put(wb, 0, [[1], [2]])
    const same = [{ sheet: 0, range: range(0, 0, 1, 0) }]
    expect(consolidateOps(wb.sheets, wb.styles, 0, range(0, 3, 0, 3), same, opts({ links: true }))).toHaveProperty(
      "problem",
    )
    expect(consolidateOps(wb.sheets, wb.styles, 0, range(1, 0, 1, 0), same, opts({}))).toEqual({
      problem: "Source references overlap the destination area.",
    })
    expect(consolidateOps(wb.sheets, wb.styles, 0, range(0, 3, 0, 3), [], opts({}))).toHaveProperty("problem")
  })
})
