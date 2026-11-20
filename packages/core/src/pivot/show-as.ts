// Value Field Settings › Show Values As: a value cell shown against other cells of its PivotTable. The rules for
// empty cells, blanks and errors are Excel 365's, read from the PivotTables of io's fixtures/excel/pivot-showas.xlsx.
import { errorValue, type CellValue } from "../values"
import type { PivotDataField } from "./model"

/** What Show Values As needs from the layout: the real (non-Values) fields of each axis and a summary lookup. */
export type ShowAsContext = {
  rowFields: readonly number[]
  colFields: readonly number[]
  /** The item positions of a field in the order shown. */
  order: (field: number) => readonly number[]
  /** The item position of a cache item of a field (undefined: not one of its items). */
  posOf: (field: number, x: number) => number | undefined
  /** The value field's summary for the items `rows` × `cols` (shorter paths: subtotals; [] the grand total). */
  raw: (rows: readonly number[], cols: readonly number[]) => CellValue | undefined
}

const num = (v: CellValue | undefined): number | undefined => (typeof v === "number" ? v : undefined)
const isError = (v: CellValue | undefined): boolean => v !== undefined && v !== null && typeof v === "object"
const DIV0 = () => errorValue("#DIV/0!")
const NULL = () => errorValue("#NULL!")
const NA = () => errorValue("#N/A")

/** `a` over `b` as Excel's percentages do: an empty `a` counts as 0, an empty `b` leaves the cell blank. */
function ratio(a: CellValue | undefined, b: CellValue | undefined): CellValue | undefined {
  if (isError(a)) return a
  if (isError(b)) return b
  const d = num(b)
  if (d === undefined) return undefined
  if (d === 0) return DIV0()
  return (num(a) ?? 0) / d
}

/**
 * The value a cell shows for value field `df`, given its own summary `v` at row items `rows` and column items `cols`
 * (item positions per real field; [] on a grand total line). Undefined: the cell is blank.
 */
export function showValueAs(
  df: PivotDataField,
  ctx: ShowAsContext,
  rows: readonly number[],
  cols: readonly number[],
  v: CellValue | undefined,
): CellValue | undefined {
  const { raw } = ctx
  switch (df.showAs) {
    case undefined:
      return v
    case "percentOfTotal":
      return ratio(v, raw([], []))
    case "percentOfRow":
      return ratio(v, raw(rows, []))
    case "percentOfCol":
      return ratio(v, raw([], cols))
    case "percentOfParentRow":
      return ratio(v, raw(rows.slice(0, Math.max(rows.length - 1, 0)), cols))
    case "percentOfParentCol":
      return ratio(v, raw(rows, cols.slice(0, Math.max(cols.length - 1, 0))))
    case "index": {
      if (isError(v)) return v
      const all = num(raw([], []))
      const row = num(raw(rows, []))
      const col = num(raw([], cols))
      if (all === undefined || row === undefined || col === undefined) return undefined
      if (row === 0 || col === 0) return DIV0()
      return ((num(v) ?? 0) * all) / (row * col)
    }
  }
  // The rest compare along the base field: the cell's line must reach it.
  const base = df.baseField ?? -1
  const onRows = ctx.rowFields.indexOf(base)
  const onCols = ctx.colFields.indexOf(base)
  if (onRows < 0 && onCols < 0) return NA()
  const level = onRows >= 0 ? onRows : onCols
  const path = onRows >= 0 ? rows : cols
  if (path.length <= level) return undefined
  const own = path[level]
  /** The summary with the base field's item `pos` in place of the cell's (inner items kept). */
  const at = (pos: number) => {
    const moved = [...path]
    moved[level] = pos
    return onRows >= 0 ? raw(moved, cols) : raw(rows, moved)
  }
  const order = ctx.order(base)
  switch (df.showAs) {
    case "percentOfParent":
      return ratio(v, onRows >= 0 ? raw(rows.slice(0, level + 1), cols) : raw(rows, cols.slice(0, level + 1)))
    case "difference":
    case "percent":
    case "percentDiff": {
      const i = order.indexOf(own)
      let other: number | undefined
      if (df.baseItem === "previous") other = i > 0 ? order[i - 1] : undefined
      else if (df.baseItem === "next") other = i >= 0 && i < order.length - 1 ? order[i + 1] : undefined
      else {
        other = ctx.posOf(base, df.baseItem ?? 0)
        if (other === undefined) return NA()
      }
      if (other === undefined) return undefined
      if (other === own) return df.showAs === "percent" && v !== undefined ? (isError(v) ? v : 1) : undefined
      const w = at(other)
      if (isError(v)) return v
      if (isError(w)) return w
      if (df.showAs === "difference") return (num(v) ?? 0) - (num(w) ?? 0)
      if (num(v) === undefined) return NULL()
      const d = num(w)
      if (d === undefined) return undefined
      if (d === 0) return DIV0()
      return df.showAs === "percent" ? num(v)! / d : (num(v)! - d) / d
    }
    case "runTotal":
    case "percentOfRunningTotal": {
      let running = 0
      let total = 0
      let any = false
      let passed = false
      for (const pos of order) {
        const n = num(at(pos))
        if (n !== undefined) any = true
        if (!passed) running += n ?? 0
        total += n ?? 0
        if (pos === own) passed = true
      }
      if (df.showAs === "runTotal") return running
      if (!any) return undefined
      return total === 0 ? DIV0() : running / total
    }
    case "rankAscending":
    case "rankDescending": {
      const mine = num(v)
      if (mine === undefined) return isError(v) ? v : undefined
      // Ties share a rank and the next value takes the next one (1, 1, 2), as Excel ranks.
      const distinct = [...new Set(order.map((pos) => num(at(pos))).filter((n): n is number => n !== undefined))]
      const better = distinct.filter((n) => (df.showAs === "rankDescending" ? n > mine : n < mine)).length
      return better + 1
    }
  }
  return v
}

/** Excel's number format for a Show Values As choice when the value field has none of its own (0.00%). */
export function showAsNumFmt(showAs: PivotDataField["showAs"]): string | undefined {
  switch (showAs) {
    case "percentOfTotal":
    case "percentOfCol":
    case "percentOfRow":
    case "percent":
    case "percentOfParentRow":
    case "percentOfParentCol":
    case "percentOfParent":
    case "percentDiff":
    case "percentOfRunningTotal":
      return "0.00%"
  }
  return undefined
}
