import { describe, expect, it } from "vitest"
import {
  ConditionalEval,
  clearRulesIn,
  dateSerial,
  defaultIconThresholds,
  parseCell,
  parseInput,
  parseRange,
  Workbook,
  type CFEntry,
  type CFRule,
  type CellWrite,
} from "../src"

const RED = { fill: "#FFC7CE", fontColor: "#9C0006" }

function book(cells: Record<string, string>) {
  const wb = new Workbook(["Sheet1"])
  const writes: CellWrite[] = Object.entries(cells).map(([ref, text]) => {
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)
    return { row, col, data: parsed ? { value: parsed.value, input: parsed.input } : undefined }
  })
  wb.transact([{ kind: "setCells", sheet: 0, cells: writes }], { label: "Typing" })
  return wb
}

let nextId = 1
function rule(wb: Workbook, ranges: string, r: CFRule, stopIfTrue = false): CFEntry {
  const entry: CFEntry = { id: nextId++, ranges: ranges.split(",").map((s) => parseRange(s)!), rule: r, stopIfTrue }
  wb.transact([{ kind: "conditional", sheet: 0, entries: [...wb.sheets[0].conditional, entry] }], { label: "CF" })
  return entry
}

const TODAY = dateSerial(2026, 11, 14) // a Saturday

function hits(wb: Workbook, refs: string[]) {
  const ev = new ConditionalEval(wb.sheets, 0, TODAY)
  return refs.filter((ref) => {
    const { row, col } = parseCell(ref)!
    return ev.at(row, col)?.style !== undefined
  })
}

describe("conditional formatting rules", () => {
  it("Cell Value greater than, less than (blank counts as 0), between either way round, equal to text", () => {
    const wb = book({ A1: "3", A2: "10", A3: "7", A4: "apple" })
    const all = ["A1", "A2", "A3", "A4", "A5"]
    rule(wb, "A1:A5", { kind: "cellIs", operator: "greaterThan", values: ["5"], style: RED })
    // Text is greater than any number (Excel's order).
    expect(hits(wb, all)).toEqual(["A2", "A3", "A4"])
    wb.transact([{ kind: "conditional", sheet: 0, entries: [] }], { label: "Clear" })
    rule(wb, "A1:A5", { kind: "cellIs", operator: "lessThan", values: ["5"], style: RED })
    expect(hits(wb, all)).toEqual(["A1", "A5"])
    wb.undo()
    rule(wb, "A1:A5", { kind: "cellIs", operator: "between", values: ["8", "3"], style: RED })
    expect(hits(wb, all)).toEqual(["A1", "A3"])
    wb.undo()
    rule(wb, "A1:A5", { kind: "cellIs", operator: "equal", values: ['"APPLE"'], style: RED })
    expect(hits(wb, all)).toEqual(["A4"])
  })

  it("values may be references: =$C$1 follows the cell it names", () => {
    const wb = book({ A1: "3", A2: "10", C1: "4" })
    rule(wb, "A1:A2", { kind: "cellIs", operator: "greaterThan", values: ["$C$1"], style: RED })
    expect(hits(wb, ["A1", "A2"])).toEqual(["A2"])
  })

  it("text that contains, begins and ends with (any case)", () => {
    const wb = book({ A1: "North East", A2: "south", A3: "Northwest", A4: "12345" })
    rule(wb, "A1:A4", { kind: "text", operator: "contains", text: "north", style: RED })
    expect(hits(wb, ["A1", "A2", "A3", "A4"])).toEqual(["A1", "A3"])
    wb.undo()
    rule(wb, "A1:A4", { kind: "text", operator: "contains", text: "234", style: RED })
    expect(hits(wb, ["A1", "A2", "A3", "A4"])).toEqual(["A4"])
  })

  it("a date occurring: yesterday, this week (Sunday to Saturday), last month", () => {
    const wb = book({ A1: "11/13/2026", A2: "11/8/2026", A3: "11/7/2026", A4: "10/31/2026", A5: "11/14/2026" })
    const all = ["A1", "A2", "A3", "A4", "A5"]
    rule(wb, "A1:A5", { kind: "date", period: "yesterday", style: RED })
    expect(hits(wb, all)).toEqual(["A1"])
    wb.undo()
    rule(wb, "A1:A5", { kind: "date", period: "thisWeek", style: RED })
    expect(hits(wb, all)).toEqual(["A1", "A2", "A5"])
    wb.undo()
    rule(wb, "A1:A5", { kind: "date", period: "lastMonth", style: RED })
    expect(hits(wb, all)).toEqual(["A4"])
  })

  it("duplicate and unique values (text without case; blanks never)", () => {
    const wb = book({ A1: "a", A2: "A", A3: "b", A4: "1", A5: "1" })
    const all = ["A1", "A2", "A3", "A4", "A5", "A6"]
    rule(wb, "A1:A6", { kind: "duplicate", unique: false, style: RED })
    expect(hits(wb, all)).toEqual(["A1", "A2", "A4", "A5"])
    wb.undo()
    rule(wb, "A1:A6", { kind: "duplicate", unique: true, style: RED })
    expect(hits(wb, all)).toEqual(["A3"])
  })

  it("top and bottom items and percentages, ties included; above and below average", () => {
    const cells: Record<string, string> = {}
    ;[5, 9, 1, 9, 3, 7, 2, 8, 6, 4].forEach((n, i) => (cells[`A${i + 1}`] = String(n)))
    const wb = book(cells)
    const all = Object.keys(cells)
    rule(wb, "A1:A10", { kind: "top", bottom: false, percent: false, rank: 2, style: RED })
    expect(hits(wb, all)).toEqual(["A2", "A4"])
    wb.undo()
    rule(wb, "A1:A10", { kind: "top", bottom: true, percent: true, rank: 20, style: RED })
    expect(hits(wb, all)).toEqual(["A3", "A7"])
    wb.undo()
    rule(wb, "A1:A10", { kind: "average", below: false, equal: false, style: RED })
    expect(hits(wb, all)).toEqual(["A2", "A4", "A6", "A8", "A9"])
  })

  it("a formula rule moves its relative references with each cell", () => {
    const wb = book({ A1: "1", B1: "x", A2: "0", B2: "y", A3: "1", B3: "z" })
    rule(wb, "B1:B3", { kind: "expression", formula: "$A1=1", style: RED })
    expect(hits(wb, ["B1", "B2", "B3"])).toEqual(["B1", "B3"])
  })

  it("priority: the first rule wins a field; Stop If True ends the list", () => {
    const wb = book({ A1: "10" })
    rule(wb, "A1", { kind: "cellIs", operator: "greaterThan", values: ["5"], style: { fill: "#FFC7CE" } })
    rule(wb, "A1", { kind: "cellIs", operator: "greaterThan", values: ["1"], style: { fill: "#C6EFCE", bold: true } })
    const ev = new ConditionalEval(wb.sheets, 0, TODAY)
    expect(ev.at(0, 0)?.style).toEqual({ fill: "#FFC7CE", bold: true })
    const entries = wb.sheets[0].conditional.map((e, i) => (i === 0 ? { ...e, stopIfTrue: true } : e))
    wb.transact([{ kind: "conditional", sheet: 0, entries }], { label: "Stop" })
    expect(new ConditionalEval(wb.sheets, 0, TODAY).at(0, 0)?.style).toEqual({ fill: "#FFC7CE" })
  })

  it("colour scales blend between min, midpoint (50th percentile) and max", () => {
    const wb = book({ A1: "0", A2: "50", A3: "100", A4: "25" })
    rule(wb, "A1:A4", {
      kind: "colorScale",
      points: [
        { at: { type: "min" }, color: "#F8696B" },
        { at: { type: "percentile", value: "50" }, color: "#FFEB84" },
        { at: { type: "max" }, color: "#63BE7B" },
      ],
    })
    const ev = new ConditionalEval(wb.sheets, 0, TODAY)
    expect(ev.at(0, 0)?.fill).toBe("#F8696B")
    expect(ev.at(2, 0)?.fill).toBe("#63BE7B")
    // The median of 0, 25, 50, 100 is 37.5; 50 lies a fifth of the way from it to 100.
    expect(ev.at(1, 0)?.fill).toBe("#E0E282")
  })

  it("data bars: lengths from zero, negative bars left of the axis", () => {
    const wb = book({ A1: "10", A2: "5", A3: "-5" })
    rule(wb, "A1:A3", {
      kind: "dataBar",
      min: { type: "automatic" },
      max: { type: "automatic" },
      color: "#638EC6",
      gradient: true,
    })
    const ev = new ConditionalEval(wb.sheets, 0, TODAY)
    const a1 = ev.at(0, 0)!.bar!
    expect(a1.axis).toBeCloseTo(1 / 3)
    expect(a1.end).toBe(1)
    const a3 = ev.at(2, 0)!.bar!
    expect(a3.negative).toBe(true)
    expect([a3.start, a3.end]).toEqual([0, 1 / 3])
  })

  it("icon sets split the range at 33% and 67%", () => {
    const wb = book({ A1: "0", A2: "40", A3: "70", A4: "100" })
    rule(wb, "A1:A4", {
      kind: "iconSet",
      set: "3Arrows",
      reverse: false,
      thresholds: defaultIconThresholds(3),
      showValue: true,
    })
    const ev = new ConditionalEval(wb.sheets, 0, TODAY)
    expect([0, 1, 2, 3].map((r) => ev.at(r, 0)?.icon?.index)).toEqual([0, 1, 2, 2])
  })

  it("rules follow inserted and deleted rows; undo puts them back", () => {
    const wb = book({ A5: "1" })
    rule(wb, "A3:A6", { kind: "blanks", blank: false, style: RED })
    wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 2 }], { label: "Insert" })
    expect(wb.sheets[0].conditional[0].ranges).toEqual([parseRange("A5:A8")])
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 4, count: 4 }], { label: "Delete" })
    expect(wb.sheets[0].conditional).toEqual([])
    wb.undo()
    wb.undo()
    expect(wb.sheets[0].conditional[0].ranges).toEqual([parseRange("A3:A6")])
  })

  it("Clear Rules from Selected Cells keeps the rest of each rule's range", () => {
    const wb = book({})
    const e = rule(wb, "A1:C3", { kind: "blanks", blank: true, style: RED })
    const left = clearRulesIn([e], [parseRange("B2")!])
    expect(left[0].ranges).toEqual([parseRange("A1:C1"), parseRange("A3:C3"), parseRange("A2"), parseRange("C2")])
    expect(clearRulesIn([e], [parseRange("A1:D9")!])).toEqual([])
  })
})
