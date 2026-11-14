import { MAX_COLS, MAX_ROWS } from "../src/address"
import type { Sheet } from "../src/sheet"
import { Workbook } from "../src/workbook"
import { contentForInput } from "../src/edit"
import { cellName, parseCell } from "../src/address"

/** Deterministic PRNG (mulberry32) for property tests. */
export function rng(seed: number) {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) >>> 0
    let t = a
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/** Everything observable about a sheet, for equality checks. */
export function snapshot(sheet: Sheet) {
  const cells = sheet.cells
    .entries({ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 })
    .map((e) => `${cellName(e)}=${JSON.stringify(e.data)}`)
  return {
    cells,
    rows: sheet.rows.overrides(0, MAX_ROWS - 1),
    cols: sheet.cols.overrides(0, MAX_COLS - 1),
  }
}

/**
 * Builds a workbook from a picture of cells: { A1: "x", B3: 5 }. Strings go through the typed-input parser.
 */
export function book(cells: Record<string, string | number | boolean> = {}) {
  const wb = new Workbook()
  const sheet = wb.sheets[0]
  for (const [name, value] of Object.entries(cells)) {
    const at = parseCell(name)!
    sheet.cells.set(at.row, at.col, contentForInput(wb.styles, undefined, String(value)))
  }
  return { wb, sheet }
}

/** Same, from rows of text: "x . x" (any non-"." token is a filled cell), starting at A1. */
export function grid(lines: string[]) {
  const wb = new Workbook()
  const sheet = wb.sheets[0]
  lines.forEach((line, row) => {
    line
      .trim()
      .split(/\s+/)
      .forEach((token, col) => {
        if (token !== ".") sheet.cells.set(row, col, { value: token })
      })
  })
  return { wb, sheet }
}
