import { describe, expect, it } from "vitest"
import { parseInput, Workbook } from "../src"

// GROUPBY, PIVOTBY, PERCENTOF and TRIMRANGE aren't in the Excel build used for the golden cases: these follow
// Microsoft's documentation.

const DATA = [
  ["Region", "Product", "Sales"],
  ["West", "Pens", "10"],
  ["East", "Ink", "20"],
  ["West", "Ink", "5"],
  ["East", "Pens", "7"],
  ["North", "Pens", "3"],
]

/** Results of a formula entered in H1 next to DATA (A1:C6), as rows. */
function run(formula: string): unknown[][] {
  const wb = new Workbook(["Data"])
  const cells = DATA.flatMap((r, row) =>
    r.map((text, col) => {
      const p = parseInput(text)!
      return { row, col, data: { value: p.value, input: p.input } }
    }),
  )
  const p = parseInput(formula)!
  cells.push({ row: 0, col: 7, data: { value: p.value, input: p.input } })
  wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load", history: false })
  wb.calculate(true)
  const s = wb.sheets[0]
  const area = wb.calc.spillAt(0, 0, 7)
  const v = (r: number, c: number) => {
    const x = s.cells.value(r, c)
    return x === undefined ? "" : x
  }
  if (!area) return [[v(0, 7)]]
  const out: unknown[][] = []
  for (let r = area.r0; r <= area.r1; r++) {
    const row: unknown[] = []
    for (let c = area.c0; c <= area.c1; c++) row.push(v(r, c))
    out.push(row)
  }
  return out
}

describe("GROUPBY", () => {
  it("sums each group, sorted by the field, with a grand total", () => {
    expect(run("=GROUPBY(A2:A6,C2:C6,SUM)")).toEqual([
      ["East", 27],
      ["North", 3],
      ["West", 15],
      ["Total", 45],
    ])
  })

  it("finds the headers itself and sorts by the values descending", () => {
    expect(run("=GROUPBY(A1:A6,C1:C6,SUM,,0,-2)")).toEqual([
      ["East", 27],
      ["West", 15],
      ["North", 3],
    ])
  })

  it("shows headers and subtotals for two levels", () => {
    expect(run("=GROUPBY(A1:B6,C1:C6,SUM,3)")).toEqual([
      ["Region", "Product", "Sales"],
      ["East", "Ink", 20],
      ["East", "Pens", 7],
      ["East", "", 27],
      ["North", "Pens", 3],
      ["North", "", 3],
      ["West", "Ink", 5],
      ["West", "Pens", 10],
      ["West", "", 15],
      ["Total", "", 45],
    ])
  })

  it("takes a LAMBDA, PERCENTOF and a filter", () => {
    expect(run("=GROUPBY(A2:A6,C2:C6,LAMBDA(v,MAX(v)),,0)")).toEqual([
      ["East", 20],
      ["North", 3],
      ["West", 10],
    ])
    expect(run('=GROUPBY(A2:A6,C2:C6,PERCENTOF,,0,,B2:B6="Pens")')).toEqual([
      ["East", 0.35],
      ["North", 0.15],
      ["West", 0.5],
    ])
  })
})

describe("PIVOTBY", () => {
  it("lays out rows by one field and columns by another, with totals", () => {
    expect(run("=PIVOTBY(A2:A6,B2:B6,C2:C6,SUM)")).toEqual([
      ["", "Ink", "Pens", "Total"],
      ["East", 20, 7, 27],
      ["North", "", 3, 3],
      ["West", 5, 10, 15],
      ["Total", 25, 20, 45],
    ])
  })

  it("is relative to the column totals for PERCENTOF", () => {
    expect(run("=PIVOTBY(A2:A6,B2:B6,C2:C6,PERCENTOF,,0,,0)")).toEqual([
      ["", "Ink", "Pens"],
      ["East", 0.8, 0.35],
      ["North", "", 0.15],
      ["West", 0.2, 0.5],
    ])
  })
})

describe("PERCENTOF and TRIMRANGE", () => {
  it("divides one sum by another", () => {
    expect(run("=PERCENTOF(C2:C3,C2:C6)")).toEqual([[30 / 45]])
    expect(run("=PERCENTOF(1,0)")).toEqual([[{ error: "#DIV/0!" }]])
  })

  it("trims blank rows and columns from the edges", () => {
    expect(run('=ROWS(TRIMRANGE(A1:E100))&"x"&COLUMNS(TRIMRANGE(A1:E100))')).toEqual([["6x3"]])
    expect(run("=ROWS(TRIMRANGE(A:A))")).toEqual([[6]])
    expect(run("=ROWS(TRIMRANGE(A1:A100,0))")).toEqual([[100]])
    expect(run("=TRIMRANGE(J1:K5)")).toEqual([[{ error: "#REF!" }]])
  })
})
