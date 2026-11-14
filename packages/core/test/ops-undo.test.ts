import { describe, expect, it } from "vitest"
import type { Op } from "../src/ops"
import { CLEAR_ALL } from "../src/styles"
import { Workbook } from "../src/workbook"
import { rng, snapshot } from "./helpers"

/** Random operations near the top-left (and at the sheet's far edges, where inserts push cells off). */
function randomOp(random: () => number): Op {
  const pick = (n: number) => Math.floor(random() * n)
  const row = () => (random() < 0.05 ? 1_048_570 + pick(6) : pick(30))
  const col = () => (random() < 0.05 ? 16_380 + pick(4) : pick(12))
  const roll = random()
  if (roll < 0.45) {
    const cells = Array.from({ length: 1 + pick(6) }, () => ({
      row: row(),
      col: col(),
      data:
        random() < 0.2
          ? undefined
          : random() < 0.2
            ? { style: 1 + pick(3) }
            : { value: random() < 0.5 ? pick(100) : `t${pick(9)}`, ...(random() < 0.3 ? { style: 1 + pick(3) } : {}) },
    }))
    return { kind: "setCells", sheet: 0, cells }
  }
  if (roll < 0.55) return { kind: "insertRows", sheet: 0, at: row(), count: 1 + pick(3) }
  if (roll < 0.65) return { kind: "deleteRows", sheet: 0, at: pick(30), count: 1 + pick(3) }
  if (roll < 0.72) return { kind: "insertCols", sheet: 0, at: col(), count: 1 + pick(2) }
  if (roll < 0.79) return { kind: "deleteCols", sheet: 0, at: pick(12), count: 1 + pick(2) }
  if (roll < 0.82) {
    const r0 = pick(20)
    const c0 = pick(10)
    const shape = random()
    const range =
      shape < 0.2
        ? { r0: 0, r1: 1_048_575, c0, c1: c0 + pick(2) }
        : shape < 0.4
          ? { r0, r1: r0 + pick(2), c0: 0, c1: 16_383 }
          : { r0, c0, r1: r0 + pick(4), c1: c0 + pick(3) }
    const patches = [{ bold: true as const }, { fill: "#ffff00" }, { bold: null }, { numberFormat: "0.00" }, CLEAR_ALL]
    return { kind: "format", sheet: 0, ranges: [range], patch: patches[pick(patches.length)] }
  }
  if (roll < 0.86) {
    const r0 = pick(20)
    const c0 = pick(10)
    const range = { r0, c0, r1: r0 + 1 + pick(2), c1: c0 + pick(2) }
    return { kind: random() < 0.7 ? "merge" : "unmerge", sheet: 0, ranges: [range] }
  }
  if (roll < 0.9) {
    const r0 = pick(20)
    const c0 = pick(10)
    const range = { r0, c0, r1: r0 + pick(3), c1: c0 + pick(3) }
    return random() < 0.5
      ? { kind: "insertCells", sheet: 0, range, shift: random() < 0.5 ? "down" : "right" }
      : { kind: "deleteCells", sheet: 0, range, shift: random() < 0.5 ? "up" : "left" }
  }
  return {
    kind: "axis",
    sheet: 0,
    axis: random() < 0.5 ? "rows" : "cols",
    entries: [{ index: pick(30), size: random() < 0.3 ? null : 5 + pick(80), hidden: random() < 0.3 }],
  }
}

describe("undo and redo", () => {
  it.each([1, 2, 3, 4, 5, 6, 7, 8])("random history %i: undo all restores the start, redo all the end", (seed) => {
    const random = rng(seed)
    const wb = new Workbook(["Sheet1"], 1000)
    const states = [snapshot(wb.sheets[0])]
    for (let i = 0; i < 60; i++) {
      wb.transact([randomOp(random), ...(random() < 0.3 ? [randomOp(random)] : [])], { label: `step ${i}` })
      states.push(snapshot(wb.sheets[0]))
    }
    for (let i = states.length - 2; i >= 0; i--) {
      wb.undo()
      expect(snapshot(wb.sheets[0])).toEqual(states[i])
    }
    expect(wb.canUndo).toBe(false)
    for (let i = 1; i < states.length; i++) {
      wb.redo()
      expect(snapshot(wb.sheets[0])).toEqual(states[i])
    }
  })

  it("keeps a 1,000-step history and drops the oldest past the limit", () => {
    const wb = new Workbook(["Sheet1"], 1000)
    for (let i = 0; i < 1200; i++) {
      wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: { value: i } }] }], {
        label: `Typing ${i}`,
      })
    }
    expect(wb.undoDepth).toBe(1000)
    for (let i = 0; i < 1000; i++) wb.undo()
    expect(wb.sheets[0].cells.value(0, 0)).toBe(199)
    expect(wb.undo()).toBeNull()
  })

  it("defaults to Excel's 100 levels, clears redo on a new change and carries caller state", () => {
    const wb = new Workbook<string>()
    expect(wb.undoLimit).toBe(100)
    wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: { value: 1 } }] }], {
      label: "Typing",
      before: "sel-before",
      after: "sel-after",
    })
    expect(wb.undoLabel).toBe("Typing")
    expect(wb.undo()!.before).toBe("sel-before")
    expect(wb.redo()!.after).toBe("sel-after")
    wb.undo()
    wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 1, col: 0, data: { value: 2 } }] }], { label: "Typing" })
    expect(wb.canRedo).toBe(false)
  })

  it("notifies listeners and bumps the version", () => {
    const wb = new Workbook()
    const seen: unknown[] = []
    wb.subscribe((changes) => seen.push(changes))
    wb.transact([{ kind: "insertRows", sheet: 0, at: 0, count: 1 }], { label: "Insert" })
    expect(wb.version).toBe(1)
    expect(seen).toEqual([[{ sheet: 0, structural: true }]])
  })
})
