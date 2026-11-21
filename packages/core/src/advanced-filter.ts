import type { Range } from "./address"
import { shiftRange, ValueContext } from "./conditional"
import { copyFormula } from "./formula/adjust"
import { evaluateFormula } from "./formula/evaluate"
import { advancedCriterion } from "./formula/functions/criteria"
import { parseFormula } from "./formula/parser"
import { isMatrix, isRef, type Scalar } from "./formula/types"
import type { CellWrite } from "./ops"
import type { Sheet } from "./sheet"
import type { CellValue } from "./values"

/**
 * Data › Sort & Filter › Advanced: a list filtered by a criteria range, in place (rows hidden) or copied elsewhere.
 * The criteria range's first row names columns of the list; each row under it is one set of conditions that must all
 * hold, and a list row shows when any set holds. A condition under a heading that names no column (or no heading) is
 * a formula worked out for each row, written for the list's first data row.
 */
export type AdvancedFilter = {
  list: Range
  criteria: { sheet: number; range: Range }
  unique: boolean
}

export function copyAdvancedFilter(f: AdvancedFilter | null): AdvancedFilter | null {
  return f
    ? { list: { ...f.list }, criteria: { sheet: f.criteria.sheet, range: { ...f.criteria.range } }, unique: f.unique }
    : null
}

/** Rows or columns inserted or deleted on `sheetIndex`: the list (and criteria there) move; a deleted list ends it. */
export function shiftAdvancedFilter(
  f: AdvancedFilter,
  sheetIndex: number,
  axis: "rows" | "cols",
  at: number,
  count: number,
): AdvancedFilter | null {
  const list = shiftRange(f.list, axis, at, count)
  if (!list) return null
  const range = f.criteria.sheet === sheetIndex ? shiftRange(f.criteria.range, axis, at, count) : f.criteria.range
  if (!range) return null
  return { ...f, list, criteria: { sheet: f.criteria.sheet, range } }
}

const textOf = (v: CellValue | undefined): string =>
  v === undefined ? "" : typeof v === "object" ? v.error : String(v)

export { advancedCriterion }

type Condition = { col: number; test: (v: Scalar) => boolean } | { formula: string }

/**
 * Which data rows of the list (its rows after the headers) meet the criteria: a flag per row, in order. `unique`
 * leaves out rows equal to one shown before (every column compared, text without regard to case).
 */
export function advancedFilterRows(
  sheets: readonly Sheet[],
  sheetIndex: number,
  filter: AdvancedFilter,
  today: number,
): boolean[] {
  const sheet = sheets[sheetIndex]
  const { list } = filter
  const headers = new Map<string, number>()
  for (let c = list.c0; c <= list.c1; c++) {
    const h = textOf(sheet.cells.value(list.r0, c)).trim().toLowerCase()
    if (h && !headers.has(h)) headers.set(h, c)
  }
  const crit = sheets[filter.criteria.sheet]
  const cr = filter.criteria.range
  const sets: Condition[][] = []
  for (let r = cr.r0 + 1; r <= cr.r1; r++) {
    const set: Condition[] = []
    for (let c = cr.c0; c <= cr.c1; c++) {
      const heading = textOf(crit.cells.value(cr.r0, c)).trim().toLowerCase()
      const col = headers.get(heading)
      const input = crit.cells.input(r, c)
      if (col === undefined) {
        // A computed criterion: a formula, true for the rows to show.
        if (input?.startsWith("=")) set.push({ formula: input })
        continue
      }
      const test = advancedCriterion(crit.cells.value(r, c))
      if (test) set.push({ col, test })
    }
    sets.push(set)
  }
  const ctx = new ValueContext(sheets, today)
  const holds = (cond: Condition, row: number): boolean => {
    if ("col" in cond) return cond.test(sheet.cells.value(row, cond.col) ?? null)
    const text = copyFormula(cond.formula, row - (list.r0 + 1), 0)
    const parsed = parseFormula(text.slice(1))
    if (!parsed.ok) return false
    ctx.sheet = sheetIndex
    ctx.row = row
    ctx.col = list.c0
    let v = evaluateFormula(parsed.ast, ctx)
    if (isRef(v)) v = ctx.cell(v.areas[0].sheet, v.areas[0].r0, v.areas[0].c0)
    if (isMatrix(v)) v = v.data[0] ?? null
    return v === true || (typeof v === "number" && v !== 0)
  }
  const out: boolean[] = []
  const seen = new Set<string>()
  for (let row = list.r0 + 1; row <= list.r1; row++) {
    // No criteria rows (or an empty one) shows every row.
    let shown = !sets.length || sets.some((set) => set.every((cond) => holds(cond, row)))
    if (shown && filter.unique) {
      const parts: string[] = []
      for (let c = list.c0; c <= list.c1; c++) {
        const v = sheet.cells.value(row, c)
        parts.push(typeof v === "string" ? "s" + v.toLowerCase() : typeof v + textOf(v))
      }
      const key = parts.join("\u0000")
      if (seen.has(key)) shown = false
      else seen.add(key)
    }
    out.push(shown)
  }
  return out
}

/**
 * Copy to another location: the headers and the rows shown, from `to`'s top-left cell. When `to` spans more than one
 * cell in its first row, its headings choose the columns (in their order); else every column is copied. Values and
 * formats are copied (formulas give their values). Null when a heading of `to` names no column of the list.
 */
export function advancedCopyCells(
  sheet: Sheet,
  list: Range,
  shown: boolean[],
  to: Range,
): { cells: CellWrite[]; area: Range } | null {
  let cols: number[] = []
  if (to.c1 > to.c0) {
    const headers = new Map<string, number>()
    for (let c = list.c0; c <= list.c1; c++) {
      const h = textOf(sheet.cells.value(list.r0, c)).trim().toLowerCase()
      if (h && !headers.has(h)) headers.set(h, c)
    }
    for (let c = to.c0; c <= to.c1; c++) {
      const h = textOf(sheet.cells.value(to.r0, c)).trim().toLowerCase()
      if (!h) continue
      const col = headers.get(h)
      if (col === undefined) return null
      cols.push(col)
    }
    if (!cols.length) return null
  } else {
    cols = Array.from({ length: list.c1 - list.c0 + 1 }, (_, i) => list.c0 + i)
  }
  const rows = [list.r0, ...shown.flatMap((s, i) => (s ? [list.r0 + 1 + i] : []))]
  const cells: CellWrite[] = []
  rows.forEach((row, ri) =>
    cols.forEach((col, ci) => {
      const data = sheet.cells.get(row, col)
      const out: CellWrite["data"] = {}
      if (data?.value !== undefined) out.value = data.value
      if (data?.style) out.style = data.style
      if (data?.input && !data.input.startsWith("=")) out.input = data.input
      cells.push({ row: to.r0 + ri, col: to.c0 + ci, data: out.value === undefined && !out.style ? undefined : out })
    }),
  )
  return { cells, area: { r0: to.r0, c0: to.c0, r1: to.r0 + rows.length - 1, c1: to.c0 + cols.length - 1 } }
}
