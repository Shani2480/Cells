import { describe, expect, it } from "vitest"
import { MAX_COLS, MAX_ROWS, cellName, parseCell, rangeName } from "../src/address"
import { currentRegion, dataEdge, lastCell, moveWithinSelection, step, type Direction } from "../src/navigation"
import { selectCell, selectRange } from "../src/selection"
import { grid } from "./helpers"

const at = (name: string) => parseCell(name)!

describe("Ctrl+arrow", () => {
  // Column A: 1..3 filled, 4..5 empty, 6..7 filled, then nothing.
  const { sheet } = grid(["x", "x", "x", ".", ".", "x", "x"])

  it.each<[string, Direction, string]>([
    ["A1", "down", "A3"], // inside a run: to its end
    ["A3", "down", "A6"], // at a run's end: to the next filled cell
    ["A4", "down", "A6"], // from an empty cell: to the next filled cell
    ["A6", "down", "A7"],
    ["A7", "down", `A${MAX_ROWS}`], // nothing below: the sheet's edge
    ["A7", "up", "A6"],
    ["A6", "up", "A3"],
    ["A3", "up", "A1"],
    ["A1", "up", "A1"],
    ["A5", "up", "A3"],
  ])("from %s %s → %s", (from, dir, to) => {
    expect(cellName(dataEdge(sheet, at(from), dir))).toBe(to)
  })

  it("works across columns", () => {
    const { sheet: row } = grid([". x x . . x"])
    expect(cellName(dataEdge(row, at("A1"), "right"))).toBe("B1")
    expect(cellName(dataEdge(row, at("B1"), "right"))).toBe("C1")
    expect(cellName(dataEdge(row, at("C1"), "right"))).toBe("F1")
    expect(cellName(dataEdge(row, at("F1"), "right"))).toBe("XFD1")
    expect(cellName(dataEdge(row, at("F1"), "left"))).toBe("C1")
    expect(cellName(dataEdge(row, at("B1"), "left"))).toBe("A1")
  })

  it("passes over hidden rows", () => {
    const { sheet: s } = grid(["x", "x", "x", "x"])
    s.rows.setHidden(3, true) // A4 hidden
    expect(cellName(dataEdge(s, at("A1"), "down"))).toBe("A3")
    s.rows.setHidden(1, true) // A2 hidden: A1 → A3 counts as adjacent
    expect(cellName(dataEdge(s, at("A1"), "down"))).toBe("A3")
    expect(cellName(step(s, at("A1"), "down"))).toBe("A3")
  })

  it("finds data far away across empty chunks", () => {
    const { sheet: s } = grid(["x"])
    s.cells.set(500_000, 0, { value: 1 })
    expect(cellName(dataEdge(s, at("A1"), "down"))).toBe("A500001")
    expect(cellName(dataEdge(s, at("A500001"), "up"))).toBe("A1")
  })
})

describe("other moves", () => {
  it("stops at the sheet's edges", () => {
    const { sheet } = grid([])
    expect(step(sheet, at("A1"), "up")).toEqual(at("A1"))
    expect(step(sheet, at("A1"), "left")).toEqual(at("A1"))
    expect(step(sheet, { row: MAX_ROWS - 1, col: MAX_COLS - 1 }, "down")).toEqual({
      row: MAX_ROWS - 1,
      col: MAX_COLS - 1,
    })
  })

  it("Ctrl+End goes to the used range's corner", () => {
    const { sheet } = grid([". . x", "x"])
    expect(cellName(lastCell(sheet))).toBe("C2")
    expect(cellName(lastCell(grid([]).sheet))).toBe("A1")
  })

  it("finds the current region through diagonals", () => {
    const { sheet } = grid(["x x . .", "x x . .", ". . x .", ". . . .", "x . . ."])
    expect(rangeName(currentRegion(sheet, at("A1")))).toBe("A1:C3")
    expect(rangeName(currentRegion(sheet, at("D4")))).toBe("A1:D5")
    expect(rangeName(currentRegion(sheet, at("A5")))).toBe("A5")
  })

  it("Enter and Tab walk a selection", () => {
    const { sheet } = grid([])
    let sel = selectRange({ r0: 0, c0: 0, r1: 1, c1: 1 })
    const path: string[] = []
    for (let i = 0; i < 5; i++) {
      sel = moveWithinSelection(sheet, sel, "down")
      path.push(cellName(sel.active))
    }
    expect(path).toEqual(["A2", "B1", "B2", "A1", "A2"])
    sel = selectRange({ r0: 0, c0: 0, r1: 1, c1: 1 })
    const tabs: string[] = []
    for (let i = 0; i < 4; i++) {
      sel = moveWithinSelection(sheet, sel, "right")
      tabs.push(cellName(sel.active))
    }
    expect(tabs).toEqual(["B1", "A2", "B2", "A1"])
    expect(cellName(moveWithinSelection(sheet, selectCell(at("C3")), "up").active)).toBe("C2")
  })
})
