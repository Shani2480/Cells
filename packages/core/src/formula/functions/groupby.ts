/**
 * GROUPBY, PIVOTBY, PERCENTOF and TRIMRANGE (Excel 365's summarising functions), built from Microsoft's documentation:
 * the Excel build Cells is checked against doesn't have them yet.
 *
 * Records are grouped by their row (and column) field values, one level per field column, sorted ascending unless a
 * sort order says otherwise. Each group's values go through the function: a LAMBDA (or built-in named without brackets)
 * of one parameter gets the group's values; one of two gets them and the set it is relative to (PERCENTOF).
 * Totals are labelled "Total"; a subtotal line shows its group's own field values with the deeper fields blank.
 */
import { isError, type CellError } from "../../values"
import { CALC, compareScalars, DIV0, REF, VALUE, toNumber } from "../coerce"
import type { EvalContext } from "../context"
import { invoke, toMatrix, toScalar } from "../evaluate"
import { isLambda, isMatrix, isRef, matrix, type LambdaValue, type Matrix, type Scalar, type Value } from "../types"
import { define } from "./registry"

const TOTAL = "Total"

type Line = {
  /** The field values down to this line's level (shorter for subtotals, empty for the grand total). */
  path: Scalar[]
  kind: "leaf" | "sub" | "total"
  members: number[]
  /** The records of the group one level up (all records at the top). */
  parent: number[]
}

/** Equal for grouping: text without regard to case, numbers and logicals as they are, blanks together. */
function keyOf(v: Scalar): string {
  if (v === null) return "b"
  if (typeof v === "string") return "s" + v.toLowerCase()
  if (typeof v === "number") return "n" + v
  if (typeof v === "boolean") return "l" + v
  return "e" + v.error
}

function compare(a: Scalar, b: Scalar): number {
  if (a === null || b === null) return a === b ? 0 : a === null ? 1 : -1
  if (isError(a) || isError(b)) return isError(a) === isError(b) ? 0 : isError(a) ? 1 : -1
  return compareScalars(a, b)
}

type Options = {
  depth: number
  /** Sort keys: 1-based field columns, then value columns after them; negative descending. */
  sort: number[]
  /** The value each group sorts by for a value column (k > fields). */
  valueKey: (members: number[], valueColumn: number) => Scalar
  table: boolean
}

/** The lines of one axis: groups level by level, sorted, with subtotals and the grand total placed by `depth`. */
function lines(fields: Scalar[][], records: number[], o: Options): Line[] {
  const levels = fields[0]?.length ?? 0
  const out: Line[] = []
  const top = o.depth < 0
  const depth = Math.abs(o.depth)
  const walk = (members: number[], level: number, path: Scalar[]) => {
    const groups = new Map<string, { value: Scalar; members: number[] }>()
    for (const r of members) {
      const v = fields[r][level]
      const k = keyOf(v)
      let g = groups.get(k)
      if (!g) groups.set(k, (g = { value: v, members: [] }))
      g.members.push(r)
    }
    const list = [...groups.values()]
    const sortKey = o.sort.find((s) => Math.abs(s) === level + 1 || Math.abs(s) > levels)
    list.sort((a, b) => {
      if (sortKey !== undefined && Math.abs(sortKey) > levels) {
        const c = compare(
          o.valueKey(a.members, Math.abs(sortKey) - levels - 1),
          o.valueKey(b.members, Math.abs(sortKey) - levels - 1),
        )
        if (c) return sortKey < 0 ? -c : c
      }
      const c = compare(a.value, b.value)
      return sortKey !== undefined && sortKey < 0 && Math.abs(sortKey) === level + 1 ? -c : c
    })
    for (const g of list) {
      const here = [...path, g.value]
      if (level === levels - 1) {
        out.push({ path: here, kind: "leaf", members: g.members, parent: members })
        continue
      }
      // Subtotals for the levels the depth reaches (depth 2: the first field's groups; 3: the second's too…).
      const sub = !o.table && level + 2 <= depth
      const line: Line = { path: here, kind: "sub", members: g.members, parent: members }
      if (sub && top) out.push(line)
      walk(g.members, level + 1, here)
      if (sub && !top) out.push(line)
    }
  }
  if (levels === 0) return [{ path: [], kind: "total", members: records, parent: records }]
  walk(records, 0, [])
  if (depth >= 1) {
    const total: Line = { path: [], kind: "total", members: records, parent: records }
    if (top) out.unshift(total)
    else out.push(total)
  }
  return out
}

/** A field or value argument as rows of values (one column per field). */
function columns(v: Value, ctx: EvalContext): Matrix | CellError {
  if (isLambda(v)) return VALUE
  return toMatrix(v, ctx)
}

function rowsOf(m: Matrix): Scalar[][] {
  const out: Scalar[][] = []
  for (let r = 0; r < m.rows; r++) out.push(m.data.slice(r * m.cols, (r + 1) * m.cols))
  return out
}

function optionalInt(v: Value | undefined, ctx: EvalContext): number | null | CellError {
  if (v === undefined || v === null) return null
  const s = toScalar(v, ctx)
  if (isError(s)) return s
  const n = toNumber(s)
  return isError(n) ? n : Math.trunc(n)
}

function sortArg(v: Value | undefined, ctx: EvalContext): number[] | CellError {
  if (v === undefined || v === null) return []
  const m = toMatrix(v, ctx)
  if (isError(m)) return m
  const out: number[] = []
  for (const x of m.data) {
    const n = toNumber(x)
    if (isError(n)) return n
    if (Math.trunc(n) === 0) return VALUE
    out.push(Math.trunc(n))
  }
  return out
}

/** Headers: whether the first row holds them, and whether they are shown. */
function headerMode(v: Value | undefined, values: Matrix, ctx: EvalContext, levels: number) {
  const n = optionalInt(v, ctx)
  if (isError(n)) return n
  if (n === null) {
    // Automatic: text over a number in the values means headers; shown when there is more than one level.
    const has = values.rows > 1 && typeof values.data[0] === "string" && typeof values.data[values.cols] === "number"
    return { has, show: has && levels > 1, generate: false }
  }
  if (n < 0 || n > 3) return VALUE
  return { has: n === 1 || n === 3, show: n >= 2, generate: n === 2 }
}

/** One aggregation: the function on a set of values (and the set it is relative to, for two-parameter functions). */
function aggregate(fn: LambdaValue, subset: Scalar[], all: Scalar[], ctx: EvalContext): Scalar {
  const column = (xs: Scalar[]) => matrix(xs.length, 1, xs)
  const two = fn.builtin === "PERCENTOF" || (!fn.builtin && fn.params.length === 2)
  const r = invoke(fn, two ? [column(subset), column(all)] : [column(subset)], ctx)
  if (isLambda(r)) return CALC
  if (isMatrix(r)) return r.rows === 1 && r.cols === 1 ? r.data[0] : CALC
  if (isRef(r)) {
    const a = r.areas[0]
    return a.r0 === a.r1 && a.c0 === a.c1 ? toScalar(r, ctx) : CALC
  }
  return r
}

function functionArg(v: Value | undefined): LambdaValue | CellError {
  if (v === undefined || v === null) return VALUE
  if (isError(v)) return v
  return isLambda(v) ? v : VALUE
}

type Prepared = {
  rows: Scalar[][]
  cols: Scalar[][] | null
  values: Scalar[][]
  rowNames: Scalar[]
  colNames: Scalar[]
  valueNames: Scalar[]
  records: number[]
  show: boolean
}

function prepare(
  rowArg: Value,
  colArg: Value | null,
  valuesArg: Value,
  headersArg: Value | undefined,
  filterArg: Value | undefined,
  ctx: EvalContext,
): Prepared | CellError {
  const R = columns(rowArg, ctx)
  if (isError(R)) return R
  const C = colArg === null ? null : columns(colArg, ctx)
  if (isError(C)) return C
  const V = columns(valuesArg, ctx)
  if (isError(V)) return V
  if (R.rows !== V.rows || (C && C.rows !== V.rows)) return VALUE
  const mode = headerMode(headersArg, V, ctx, Math.max(R.cols, C?.cols ?? 0))
  if (isError(mode)) return mode
  const start = mode.has ? 1 : 0
  const rows = rowsOf(R)
  const cols = C ? rowsOf(C) : null
  const values = rowsOf(V)
  const names = (m: Matrix, all: Scalar[][], generated: string): Scalar[] =>
    Array.from({ length: m.cols }, (_, i) =>
      mode.has ? all[0][i] : m.cols > 1 || generated !== "Value" ? `${generated} ${i + 1}` : generated,
    )
  let filter: Scalar[] | null = null
  if (filterArg !== undefined && filterArg !== null) {
    const F = toMatrix(filterArg, ctx)
    if (isError(F)) return F
    if (F.cols !== 1 || F.rows !== V.rows - start) return VALUE
    filter = F.data
  }
  const records: number[] = []
  for (let r = start; r < V.rows; r++) {
    const f = filter ? filter[r - start] : true
    if (isError(f)) return f
    if (f === true || (typeof f === "number" && f !== 0)) records.push(r)
  }
  return {
    rows,
    cols,
    values,
    rowNames: names(R, rows, "Row Field"),
    colNames: C ? names(C, cols!, "Col Field") : [],
    valueNames: names(V, values, "Value"),
    records,
    show: mode.show,
  }
}

const valuesIn = (p: Prepared, members: number[], column: number) => members.map((r) => p.values[r][column])

/** The default depth: grand totals, and subtotals where there are two or more field levels. */
function depthArg(v: Value | undefined, levels: number, ctx: EvalContext): number | CellError {
  const n = optionalInt(v, ctx)
  if (isError(n)) return n
  return n === null ? (levels > 1 ? 2 : 1) : n
}

define("GROUPBY", {
  min: 3,
  max: 8,
  call: (a, ctx) => {
    const fn = functionArg(a[2])
    if (isError(fn)) return fn
    const p = prepare(a[0], null, a[1], a[3], a[6], ctx)
    if (isError(p)) return p
    const levels = p.rows[0]?.length ?? 0
    const depth = depthArg(a[4], levels, ctx)
    if (isError(depth)) return depth
    const sort = sortArg(a[5], ctx)
    if (isError(sort)) return sort
    const rel = optionalInt(a[7], ctx)
    if (isError(rel)) return rel
    const all = (column: number) => valuesIn(p, p.records, column)
    const valueCols = p.values[0]?.length ?? 0
    const ls = lines(p.rows, p.records, {
      depth,
      sort,
      table: rel === 1,
      valueKey: (members, column) => aggregate(fn, valuesIn(p, members, column), all(column), ctx),
    })
    const out: Scalar[][] = []
    if (p.show) out.push([...p.rowNames, ...p.valueNames])
    for (const line of ls) {
      const labels: Scalar[] = Array.from({ length: levels }, (_, i) => line.path[i] ?? "")
      if (line.kind === "total") labels[0] = TOTAL
      const cells: Scalar[] = []
      for (let c = 0; c < valueCols; c++) cells.push(aggregate(fn, valuesIn(p, line.members, c), all(c), ctx))
      out.push([...labels, ...cells])
    }
    if (!out.length) return CALC
    return matrix(out.length, out[0].length, out.flat())
  },
})

define("PIVOTBY", {
  min: 4,
  max: 11,
  call: (a, ctx) => {
    const fn = functionArg(a[3])
    if (isError(fn)) return fn
    const p = prepare(a[0], a[1], a[2], a[4], a[9], ctx)
    if (isError(p)) return p
    const rowLevels = p.rows[0]?.length ?? 0
    const colLevels = p.cols![0]?.length ?? 0
    const rowDepth = depthArg(a[5], rowLevels, ctx)
    if (isError(rowDepth)) return rowDepth
    const rowSort = sortArg(a[6], ctx)
    if (isError(rowSort)) return rowSort
    const colDepth = depthArg(a[7], colLevels, ctx)
    if (isError(colDepth)) return colDepth
    const colSort = sortArg(a[8], ctx)
    if (isError(colSort)) return colSort
    const relative = optionalInt(a[10], ctx)
    if (isError(relative)) return relative
    const rel = relative ?? 0
    if (rel < 0 || rel > 4) return VALUE
    const valueCols = p.values[0]?.length ?? 0
    const total = (members: number[], column: number) =>
      aggregate(fn, valuesIn(p, members, column), valuesIn(p, p.records, column), ctx)
    const rowLines = lines(p.rows, p.records, { depth: rowDepth, sort: rowSort, table: false, valueKey: total })
    const colLines = lines(p.cols!, p.records, { depth: colDepth, sort: colSort, table: false, valueKey: total })

    const out: Scalar[][] = []
    const width = rowLevels + colLines.length * valueCols
    const header = (cells: Scalar[]) => out.push([...cells, ...new Array(width - cells.length).fill("")])
    // Header rows: one per column field level (the field names in the corner when headers are shown)…
    for (let level = 0; level < colLevels; level++) {
      const row: Scalar[] = new Array(rowLevels).fill("")
      if (p.show && level === colLevels - 1 && valueCols === 1) p.rowNames.forEach((n, i) => (row[i] = n))
      else if (p.show && level === 0) row[0] = p.colNames.join(" / ")
      for (const line of colLines)
        for (let v = 0; v < valueCols; v++)
          row.push(line.kind === "total" ? (level === 0 ? TOTAL : "") : (line.path[level] ?? ""))
      header(row)
    }
    // …and one naming the values when there are several.
    if (valueCols > 1) {
      const row: Scalar[] = p.show ? [...p.rowNames] : new Array(rowLevels).fill("")
      for (let i = 0; i < colLines.length; i++) row.push(...p.valueNames)
      header(row)
    }
    const inBoth = (r: number[], c: Set<number>) => r.filter((x) => c.has(x))
    for (const rl of rowLines) {
      const labels: Scalar[] = Array.from({ length: rowLevels }, (_, i) => rl.path[i] ?? "")
      if (rl.kind === "total") labels[0] = TOTAL
      const row: Scalar[] = [...labels]
      for (const cl of colLines) {
        const colSet = new Set(cl.members)
        const members = inBoth(rl.members, colSet)
        // The set a two-parameter function is relative to: column, row, grand or parent totals.
        const relativeTo =
          rel === 0
            ? cl.members
            : rel === 1
              ? rl.members
              : rel === 2
                ? p.records
                : rel === 3
                  ? inBoth(rl.members, new Set(cl.parent))
                  : inBoth(rl.parent, colSet)
        for (let v = 0; v < valueCols; v++)
          row.push(
            members.length || rl.kind !== "leaf" || cl.kind !== "leaf"
              ? aggregate(fn, valuesIn(p, members, v), valuesIn(p, relativeTo, v), ctx)
              : "",
          )
      }
      out.push(row)
    }
    if (!out.length) return CALC
    return matrix(out.length, width, out.flat())
  },
})

define("PERCENTOF", {
  min: 2,
  max: 2,
  call: (a, ctx) => {
    const sum = (v: Value): number | CellError => {
      const m = toMatrix(v, ctx)
      if (isError(m)) return m
      let s = 0
      for (const x of m.data) {
        if (isError(x)) return x
        if (typeof x === "number") s += x
      }
      return s
    }
    const part = sum(a[0])
    if (isError(part)) return part
    const whole = sum(a[1])
    if (isError(whole)) return whole
    return whole === 0 ? DIV0 : part / whole
  },
})

define("TRIMRANGE", {
  min: 1,
  max: 3,
  call: (a, ctx) => {
    const mode = (v: Value | undefined): number | CellError => {
      const n = optionalInt(v, ctx)
      if (isError(n)) return n
      if (n === null) return 3
      return n >= 0 && n <= 3 ? n : VALUE
    }
    const rowsMode = mode(a[1])
    if (isError(rowsMode)) return rowsMode
    const colsMode = mode(a[2])
    if (isError(colsMode)) return colsMode
    const src = a[0]
    if (isRef(src) && src.areas.length !== 1) return VALUE
    // A range is scanned from its start to the end of what is used in it (whole columns stay quick).
    const full = isRef(src) ? src.areas[0] : null
    const used = full ? ctx.usedArea(full) : null
    const scan = full && used ? { ...full, r1: Math.min(full.r1, used.r1), c1: Math.min(full.c1, used.c1) } : null
    const m = full ? (scan ? toMatrix({ kind: "ref", areas: [scan] }, ctx) : null) : toMatrix(src, ctx)
    if (isError(m)) return m
    // A range with nothing in it trims to nothing.
    if (!m) return REF
    const blankRow = (r: number) => m.data.slice(r * m.cols, (r + 1) * m.cols).every((v) => v === null)
    const blankCol = (c: number) =>
      Array.from({ length: m.rows }, (_, r) => m.data[r * m.cols + c]).every((v) => v === null)
    let r0 = 0
    let r1 = m.rows - 1
    let c0 = 0
    let c1 = m.cols - 1
    if (rowsMode & 1) while (r0 <= r1 && blankRow(r0)) r0++
    if (rowsMode & 2) while (r1 >= r0 && blankRow(r1)) r1--
    if (colsMode & 1) while (c0 <= c1 && blankCol(c0)) c0++
    if (colsMode & 2) while (c1 >= c0 && blankCol(c1)) c1--
    if (r0 > r1 || c0 > c1) return REF
    if (full) {
      return {
        kind: "ref",
        areas: [
          {
            sheet: full.sheet,
            r0: full.r0 + r0,
            r1: rowsMode & 2 ? full.r0 + r1 : full.r1,
            c0: full.c0 + c0,
            c1: colsMode & 2 ? full.c0 + c1 : full.c1,
          },
        ],
      }
    }
    const data: Scalar[] = []
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) data.push(m.data[r * m.cols + c])
    return matrix(r1 - r0 + 1, c1 - c0 + 1, data)
  },
})
