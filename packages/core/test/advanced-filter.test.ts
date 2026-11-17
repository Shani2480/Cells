import { describe, expect, it } from "vitest"
import { advancedCopyCells, advancedCriterion, advancedFilterRows, Workbook, type AdvancedFilter } from "../src"

type V = string | number
function book(rows: V[][], at = { row: 0, col: 0 }, sheet = 0, wb = new Workbook(["Data", "Other"])) {
  const cells = rows.flatMap((r, ri) =>
    r.map((v, ci) => ({
      row: at.row + ri,
      col: at.col + ci,
      data:
        typeof v === "string" && v.startsWith("=") && !v.startsWith("==")
          ? { input: v, value: 0 }
          : { value: typeof v === "string" && v.startsWith("==") ? v.slice(1) : v },
    })),
  )
  wb.transact([{ kind: "setCells", sheet, cells: cells.filter((c) => c.data.value !== "") }], {
    label: "Load",
    history: false,
  })
  wb.calculate(true)
  return wb
}

const LIST: V[][] = [
  ["Name", "Region", "Sales"],
  ["Smith", "East", 100],
  ["Jones", "West", 250],
  ["Smithers", "East", 40],
  ["Brown", "North", 300],
  ["Smith", "East", 100],
]
const list = { r0: 0, c0: 0, r1: 5, c1: 2 }

function run(criteria: V[][], unique = false) {
  const wb = book(LIST)
  book(criteria, { row: 0, col: 5 }, 0, wb)
  const filter: AdvancedFilter = {
    list,
    criteria: { sheet: 0, range: { r0: 0, c0: 5, r1: criteria.length - 1, c1: 5 + criteria[0].length - 1 } },
    unique,
  }
  return { wb, filter, shown: advancedFilterRows(wb.sheets, 0, filter, 0) }
}

describe("Advanced Filter criteria", () => {
  it("text begins the cell's text; = is exact; operators compare", () => {
    expect(advancedCriterion("Sm")!("Smithers")).toBe(true)
    expect(advancedCriterion("=Smith")!("Smithers")).toBe(false)
    expect(advancedCriterion("=Smith")!("smith")).toBe(true)
    expect(advancedCriterion(">100")!(250)).toBe(true)
    expect(advancedCriterion("<>East")!("West")).toBe(true)
    expect(advancedCriterion("=")!(null)).toBe(true)
    expect(advancedCriterion(100)!(100)).toBe(true)
    expect(advancedCriterion("")).toBeNull()
  })

  it("conditions in a row all hold; rows are alternatives", () => {
    expect(
      run([
        ["Region", "Sales"],
        ["East", ">50"],
        ["North", ""],
      ]).shown,
    ).toEqual([true, false, false, true, true])
  })

  it("a computed criterion (no matching heading) is a formula for the first data row", () => {
    expect(run([["Above average"], ["=C2>AVERAGE($C$2:$C$6)"]]).shown).toEqual([false, true, false, true, false])
  })

  it("unique records only leaves out repeated rows", () => {
    expect(run([["Name"], ["Smith"]], true).shown).toEqual([true, false, true, false, false])
  })

  it("copies the headers and rows shown, all columns or the ones named", () => {
    const { wb, filter, shown } = run([["Region"], ["East"]])
    const all = advancedCopyCells(wb.sheets[0], filter.list, shown, { r0: 10, c0: 0, r1: 10, c1: 0 })!
    expect(all.area).toEqual({ r0: 10, c0: 0, r1: 13, c1: 2 })
    const sheet = wb.sheets[0]
    sheet.cells.set(20, 0, { value: "Sales" })
    sheet.cells.set(20, 1, { value: "name" })
    const some = advancedCopyCells(sheet, filter.list, shown, { r0: 20, c0: 0, r1: 20, c1: 1 })!
    const at = (row: number, col: number) => some.cells.find((c) => c.row === row && c.col === col)?.data?.value
    expect([at(20, 0), at(20, 1), at(21, 0), at(21, 1), at(23, 1)]).toEqual(["Sales", "Name", 100, "Smith", "Smith"])
    sheet.cells.set(20, 1, { value: "Missing" })
    expect(advancedCopyCells(sheet, filter.list, shown, { r0: 20, c0: 0, r1: 20, c1: 1 })).toBeNull()
  })

  it("in place: hides rows with one undo step, and moves with inserted rows", () => {
    const { wb, filter, shown } = run([["Region"], ["East"]])
    const rows = shown.map((s, i) => ({ index: 1 + i, size: null, hidden: !s }))
    wb.transact([{ kind: "advancedFilter", sheet: 0, filter, rows }], { label: "Advanced Filter" })
    const s = wb.sheets[0]
    expect([1, 2, 3, 4, 5].map((r) => s.rows.isHidden(r))).toEqual([false, true, false, true, false])
    wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 1 }], { label: "Insert" })
    expect(s.advancedFilter?.list).toEqual({ r0: 1, c0: 0, r1: 6, c1: 2 })
    wb.undo()
    wb.undo()
    expect(s.advancedFilter).toBeNull()
    expect([1, 2, 3, 4, 5].some((r) => s.rows.isHidden(r))).toBe(false)
  })
})
