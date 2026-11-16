import { describe, expect, it } from "vitest"
import {
  clearOutlineOps,
  detailOps,
  groupOps,
  outlineGroups,
  removeSubtotalsOps,
  showLevelOps,
  snapshotSheet,
  subtotalOps,
  ungroupOps,
  Workbook,
} from "../src"

const levels = (wb: Workbook, n = 12) => Array.from({ length: n }, (_, r) => wb.sheets[0].rows.level(r)).join("")
const hidden = (wb: Workbook, n = 12) =>
  Array.from({ length: n }, (_, r) => (wb.sheets[0].rows.isHidden(r) ? "h" : ".")).join("")

describe("outline groups", () => {
  it("groups, nests, collapses, opens and ungroups rows, all undoable", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    s.rows.setSize(2, 40)
    wb.transact(groupOps(s, 0, "rows", 1, 6)!, { label: "Group" })
    wb.transact(groupOps(s, 0, "rows", 2, 3)!, { label: "Group" })
    expect(levels(wb, 9)).toBe("012211100")
    expect(s.rows.maxLevel).toBe(2)
    expect(s.rows.size(2)).toBe(40)
    expect(outlineGroups(s.rows, true)).toEqual([
      { level: 1, start: 1, end: 6, summary: 7 },
      { level: 2, start: 2, end: 3, summary: 4 },
    ])
    // Level 1: only the top rows; then open the outer group: the inner one stays shut.
    wb.transact(showLevelOps(s, 0, "rows", 1), { label: "Level" })
    expect(hidden(wb, 9)).toBe(".hhhhhh..")
    expect(s.rows.isCollapsed(7) && s.rows.isCollapsed(4)).toBe(true)
    wb.transact(detailOps(s, 0, "rows", 7, 7, true)!, { label: "Show Detail" })
    expect(hidden(wb, 9)).toBe("..hh.....")
    wb.transact(detailOps(s, 0, "rows", 4, 4, true)!, { label: "Show Detail" })
    expect(hidden(wb, 9)).toBe(".........")
    wb.transact(detailOps(s, 0, "rows", 2, 2, false)!, { label: "Hide Detail" })
    expect(hidden(wb, 9)).toBe("..hh.....")
    // Ungrouping the collapsed inner group shows its rows.
    wb.transact(ungroupOps(s, 0, "rows", 2, 3)!, { label: "Ungroup" })
    expect([levels(wb, 9), hidden(wb, 9), s.rows.isCollapsed(4)]).toEqual(["011111100", ".........", false])
    wb.undo()
    expect([levels(wb, 9), hidden(wb, 9)]).toEqual(["012211100", "..hh....."])
    wb.transact(clearOutlineOps(s, 0, "rows"), { label: "Clear Outline" })
    expect(s.rows.maxLevel).toBe(0)
    expect([s.rows.storedSize(2), hidden(wb, 9)]).toEqual([40, "..hh....."])
    expect(ungroupOps(s, 0, "rows", 0, 5)).toBeNull()
  })

  it("follows inserted and deleted rows, and copies with the sheet", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    wb.transact(groupOps(s, 0, "rows", 2, 4)!, { label: "Group" })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 3, count: 2 }], { label: "Insert" })
    expect(levels(wb, 9)).toBe("001111100")
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 1, count: 3 }], { label: "Delete" })
    expect(levels(wb, 9)).toBe("011100000")
    wb.undo()
    expect(levels(wb, 9)).toBe("001111100")
    expect(
      snapshotSheet(s)
        .rows.filter((e) => e.level)
        .map((e) => e.index),
    ).toEqual([2, 3, 4, 5, 6])
  })

  it("puts summaries above the detail when Settings says so", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    wb.transact(
      [{ kind: "outlineSettings", sheet: 0, settings: { summaryBelow: false, summaryRight: true, showSymbols: true } }],
      { label: "Settings" },
    )
    wb.transact(groupOps(s, 0, "rows", 3, 5)!, { label: "Group" })
    expect(outlineGroups(s.rows, s.outline.summaryBelow)).toEqual([{ level: 1, start: 3, end: 5, summary: 2 }])
    wb.undo()
    wb.undo()
    expect(s.outline.summaryBelow).toBe(true)
  })
})

describe("Subtotal", () => {
  const data = () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    const rows: [string, number][] = [
      ["East", 10],
      ["East", 20],
      ["West", 5],
      ["North", 1],
      ["North", 2],
    ]
    s.cells.set(0, 0, { value: "Region" })
    s.cells.set(0, 1, { value: "Sales" })
    rows.forEach(([r, v], i) => {
      s.cells.set(i + 1, 0, { value: r })
      s.cells.set(i + 1, 1, { value: v })
    })
    return wb
  }
  const column = (wb: Workbook, col: number, n: number) =>
    Array.from({ length: n }, (_, r) => wb.sheets[0].cells.input(r, col) ?? wb.sheets[0].cells.value(r, col) ?? "")

  it("adds a total after each run and a grand total, grouped in three levels; Remove All takes them out", () => {
    const wb = data()
    const s = wb.sheets[0]
    const range = { r0: 0, c0: 0, r1: 5, c1: 1 }
    const result = subtotalOps(s, wb.styles, 0, range, {
      by: 0,
      fn: 0,
      cols: [1],
      replace: true,
      pageBreaks: false,
      below: true,
    })
    if ("problem" in result) throw new Error(result.problem)
    wb.transact(result.ops, { label: "Subtotal" })
    wb.calculate(true)
    expect(column(wb, 0, 10)).toEqual([
      "Region",
      "East",
      "East",
      "East Total",
      "West",
      "West Total",
      "North",
      "North",
      "North Total",
      "Grand Total",
    ])
    expect(column(wb, 1, 10).slice(3, 10)).toEqual([
      "=SUBTOTAL(9,B2:B3)",
      5,
      "=SUBTOTAL(9,B5:B5)",
      1,
      2,
      "=SUBTOTAL(9,B7:B8)",
      "=SUBTOTAL(9,B2:B9)",
    ])
    expect(s.cells.value(9, 1)).toBe(38)
    expect(levels(wb, 11)).toBe("02212122100")
    expect(result.range).toEqual({ r0: 0, c0: 0, r1: 9, c1: 1 })
    // Level 2 shows the totals only.
    wb.transact(showLevelOps(s, 0, "rows", 2), { label: "Level" })
    expect(hidden(wb, 10)).toBe(".hh.h.hh..")
    // Subtotal again with Replace: the same result, not doubled.
    const again = subtotalOps(s, wb.styles, 0, result.range, {
      by: 0,
      fn: 0,
      cols: [1],
      replace: true,
      pageBreaks: false,
      below: true,
    })
    if ("problem" in again) throw new Error(again.problem)
    wb.transact(again.ops, { label: "Subtotal" })
    expect(column(wb, 0, 11)).toEqual([
      "Region",
      "East",
      "East",
      "East Total",
      "West",
      "West Total",
      "North",
      "North",
      "North Total",
      "Grand Total",
      "",
    ])
    wb.transact(removeSubtotalsOps(s, 0, again.range), { label: "Remove All" })
    expect(column(wb, 0, 7)).toEqual(["Region", "East", "East", "West", "North", "North", ""])
    expect(s.rows.maxLevel).toBe(0)
    expect(hidden(wb, 7)).toBe(".......")
  })

  it("puts totals above with Summary below data off, and counts with Count", () => {
    const wb = data()
    const s = wb.sheets[0]
    const result = subtotalOps(
      s,
      wb.styles,
      0,
      { r0: 0, c0: 0, r1: 5, c1: 1 },
      { by: 0, fn: 1, cols: [1], replace: true, pageBreaks: true, below: false },
    )
    if ("problem" in result) throw new Error(result.problem)
    wb.transact(result.ops, { label: "Subtotal" })
    wb.calculate(true)
    expect(column(wb, 0, 10)).toEqual([
      "Region",
      "Grand Count",
      "East Count",
      "East",
      "East",
      "West Count",
      "West",
      "North Count",
      "North",
      "North",
    ])
    expect(s.cells.input(2, 1)).toBe("=SUBTOTAL(3,B4:B5)")
    expect(s.cells.value(1, 1)).toBe(5)
    expect(s.pageSetup.rowBreaks).toEqual([5, 7])
  })
})

describe("Subtotal's list", () => {
  it("leaves out empty rows at its end", () => {
    const wb = new Workbook()
    const s = wb.sheets[0]
    ;["Region", "East", "East", "West"].forEach((v, r) => s.cells.set(r, 0, { value: v }))
    ;[0, 1, 2, 3].forEach((r) => r && s.cells.set(r, 1, { value: r }))
    const result = subtotalOps(
      s,
      wb.styles,
      0,
      { r0: 0, c0: 0, r1: 5, c1: 1 },
      { by: 0, fn: 0, cols: [1], replace: true, pageBreaks: false, below: true },
    )
    if ("problem" in result) throw new Error(result.problem)
    wb.transact(result.ops, { label: "Subtotal" })
    expect(Array.from({ length: 7 }, (_, r) => s.cells.value(r, 0) ?? "")).toEqual([
      "Region",
      "East",
      "East",
      "East Total",
      "West",
      "West Total",
      "Grand Total",
    ])
  })
})
