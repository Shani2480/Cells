import { columnName, MAX_COLS, MAX_ROWS } from "../../address"
import { isError, type CellError } from "../../values"
import type { Node } from "../ast"
import { CALC, compareScalars, compareText, NA, NUM, REF, toBool, toText, VALUE } from "../coerce"
import type { EvalContext } from "../context"
import { evaluate, toMatrix, toScalar } from "../evaluate"
import { parseFormula } from "../parser"
import { sheetPrefix } from "../print"
import { isMatrix, isRef, matrix, type Area, type Matrix, type RefValue, type Scalar, type Value } from "../types"
import { wildcardPattern } from "./criteria"
import { int, num } from "./helpers"
import { define } from "./registry"
import { pivotDataCell, pivotTableRange } from "../../pivot/getpivotdata"

// ── Searching ──────────────────────────────────────────────────────────────────────────────────────────────────

/** A one-dimensional list of values to search: a row or column of a reference, or of an array. */
type Vector = { length: number; at(i: number): Scalar; area?: Area; vertical: boolean }

function vectorOf(v: Value, ctx: EvalContext): Vector | CellError {
  if (isRef(v)) {
    if (v.areas.length !== 1) return NA
    const a = v.areas[0]
    const vertical = a.c0 === a.c1
    if (!vertical && a.r0 !== a.r1) return NA
    const used = ctx.usedArea(a)
    // Beyond the used rows everything is blank; trim long whole-column vectors to the used part.
    const length = vertical ? (used ? used.r1 - a.r0 + 1 : 0) : used ? used.c1 - a.c0 + 1 : 0
    return {
      length: Math.max(0, length),
      at: (i) => ctx.cell(a.sheet, vertical ? a.r0 + i : a.r0, vertical ? a.c0 : a.c0 + i),
      area: a,
      vertical,
    }
  }
  const m = isMatrix(v) ? v : matrix(1, 1, [v as Scalar])
  if (m.rows !== 1 && m.cols !== 1) return NA
  return { length: m.data.length, at: (i) => m.data[i], vertical: m.cols === 1 }
}

function key(v: Scalar): string | null {
  if (typeof v === "number") return `n${Number(v.toPrecision(15))}`
  if (typeof v === "string") return `s${v.toLowerCase()}`
  if (typeof v === "boolean") return `b${v}`
  return null
}

/** First (or last) exact position of each value, cached until a cell of the column changes. */
function exactIndex(vec: Vector, ctx: EvalContext, last: boolean): Map<string, number> | null {
  const a = vec.area
  if (!a || !ctx.cache || !ctx.columnVersion || !vec.vertical) return null
  const cacheKey = `x:${a.sheet}:${a.r0}:${a.c0}:${vec.length}:${last}`
  const version = ctx.columnVersion(a.sheet, a.c0)
  const hit = ctx.cache.get(cacheKey) as { version: number; map: Map<string, number> } | undefined
  if (hit && hit.version === version) return hit.map
  const map = new Map<string, number>()
  for (let i = 0; i < vec.length; i++) {
    const k = key(vec.at(i))
    if (k !== null && (last || !map.has(k))) map.set(k, i)
  }
  ctx.cache.set(cacheKey, { version, map })
  return map
}

function sameValue(x: Scalar, y: Scalar): boolean {
  if (x === null || y === null || isError(x) || isError(y)) return false
  if (typeof x !== typeof y) return false
  return compareScalars(x, y) === 0
}

/** Exact match, with * ? ~ wildcards for text when allowed. */
function findExact(vec: Vector, target: Scalar, ctx: EvalContext, wildcards: boolean, reverse = false): number {
  if (typeof target === "string" && wildcards) {
    const pattern = wildcardPattern(target)
    if (pattern) {
      for (let k = 0; k < vec.length; k++) {
        const i = reverse ? vec.length - 1 - k : k
        const v = vec.at(i)
        if (typeof v === "string" && pattern.test(v)) return i
      }
      return -1
    }
  }
  const index = exactIndex(vec, ctx, reverse)
  if (index) {
    const k = key(target)
    return k === null ? -1 : (index.get(k) ?? -1)
  }
  for (let k = 0; k < vec.length; k++) {
    const i = reverse ? vec.length - 1 - k : k
    if (sameValue(vec.at(i), target)) return i
  }
  return -1
}

/**
 * Approximate match in sorted data (MATCH 1 / -1, VLOOKUP TRUE): binary search for the last value ≤ target (or the
 * last ≥ target when descending), skipping values of another type the way Excel does.
 */
function findSorted(vec: Vector, target: Scalar, descending: boolean): number {
  let lo = 0
  let hi = vec.length - 1
  let found = -1
  while (lo <= hi) {
    let mid = (lo + hi) >> 1
    // Step over blanks and values of a different kind towards lo.
    let probe = mid
    while (probe >= lo && !comparable(vec.at(probe), target)) probe--
    if (probe < lo) {
      lo = mid + 1
      continue
    }
    mid = probe
    const c = compareScalars(vec.at(mid) as Exclude<Scalar, CellError>, target as Exclude<Scalar, CellError>)
    if (descending ? c >= 0 : c <= 0) {
      found = mid
      lo = mid + 1
    } else hi = mid - 1
  }
  return found
}

function comparable(v: Scalar, target: Scalar) {
  if (v === null || isError(v)) return false
  return typeof v === typeof target
}

// ── VLOOKUP / HLOOKUP / MATCH / LOOKUP ────────────────────────────────────────────────────────────────────────

function tableLookup(vertical: boolean) {
  return (a: Value[], ctx: EvalContext): Value => {
    const target = toScalar(a[0], ctx)
    if (isError(target)) return target
    const table = a[1]
    const n = int(a[2])
    if (isError(n)) return n
    const approx = a.length > 3 && a[3] !== null ? toBool(toScalar(a[3], ctx)) : true
    if (isError(approx)) return approx
    let area: Area | null = null
    let m: Matrix | null = null
    if (isRef(table)) {
      if (table.areas.length !== 1) return REF
      area = table.areas[0]
    } else {
      const mm = toMatrix(table, ctx)
      if (isError(mm)) return mm
      m = mm
    }
    const rows = area ? area.r1 - area.r0 + 1 : m!.rows
    const cols = area ? area.c1 - area.c0 + 1 : m!.cols
    if (n < 1) return VALUE
    if (n > (vertical ? cols : rows)) return REF
    const firstLine: Vector = area
      ? (vectorOf(
          { kind: "ref", areas: [vertical ? { ...area, c1: area.c0 } : { ...area, r1: area.r0 }] },
          ctx,
        ) as Vector)
      : { length: vertical ? rows : cols, at: (i) => (vertical ? m!.data[i * cols] : m!.data[i]), vertical }
    const i = approx ? findSorted(firstLine, target, false) : findExact(firstLine, target, ctx, true)
    if (i < 0) return NA
    if (area)
      return (
        ctx.cell(area.sheet, vertical ? area.r0 + i : area.r0 + n - 1, vertical ? area.c0 + n - 1 : area.c0 + i) ?? 0
      )
    return (vertical ? m!.data[i * cols + n - 1] : m!.data[(n - 1) * cols + i]) ?? 0
  }
}
define("VLOOKUP", { min: 3, max: 4, params: ["scalar", "any", "scalar", "scalar"], call: tableLookup(true) })
define("HLOOKUP", { min: 3, max: 4, params: ["scalar", "any", "scalar", "scalar"], call: tableLookup(false) })

define("MATCH", {
  min: 2,
  max: 3,
  params: ["scalar", "any", "scalar"],
  call: (a, ctx) => {
    const target = a[0] as Scalar
    const vec = vectorOf(a[1], ctx)
    if (isError(vec)) return vec
    const type = a.length > 2 && a[2] !== null ? num(a[2]) : 1
    if (isError(type)) return type
    let i: number
    if (type === 0) i = findExact(vec, target, ctx, true)
    else i = findSorted(vec, target, type < 0)
    return i < 0 ? NA : i + 1
  },
})
define("LOOKUP", {
  min: 2,
  max: 3,
  params: ["scalar", "any", "any"],
  call: (a, ctx) => {
    const target = a[0] as Scalar
    let lookIn: Value = a[1]
    let result: Value | undefined = a[2]
    if (result === undefined) {
      // Array form: search the first row or column (whichever is longer side), return from the last.
      const m = toMatrix(lookIn, ctx)
      if (isError(m)) return m
      const vertical = m.rows >= m.cols
      lookIn = matrix(
        vertical ? m.rows : m.cols,
        1,
        Array.from({ length: vertical ? m.rows : m.cols }, (_, k) => (vertical ? m.data[k * m.cols] : m.data[k])),
      )
      result = matrix(
        vertical ? m.rows : m.cols,
        1,
        Array.from({ length: vertical ? m.rows : m.cols }, (_, k) =>
          vertical ? m.data[k * m.cols + m.cols - 1] : m.data[(m.rows - 1) * m.cols + k],
        ),
      )
    }
    const vec = vectorOf(lookIn, ctx)
    if (isError(vec)) return vec
    const out = vectorOf(result, ctx)
    if (isError(out)) return out
    const i = findSorted(vec, target, false)
    if (i < 0) return NA
    return out.at(i) ?? (i < out.length ? 0 : NA)
  },
})

// ── XLOOKUP / XMATCH ──────────────────────────────────────────────────────────────────────────────────────────

function xfind(vec: Vector, target: Scalar, mode: number, search: number, ctx: EvalContext): number | CellError {
  if (search === 2 || search === -2) {
    // Binary search (sorted ascending or descending).
    const exact = findSorted(vec, target, search === -2)
    if (exact >= 0 && sameValue(vec.at(exact), target)) return exact
    if (mode === 0 || mode === 2) return -1
    if (mode === -1) return search === 2 ? exact : exact + 1 < vec.length ? exact + 1 : -1
    return search === 2 ? (exact + 1 < vec.length ? exact + 1 : -1) : exact
  }
  if (search !== 1 && search !== -1) return VALUE
  const reverse = search === -1
  if (mode === 0 || mode === 2) return findExact(vec, target, ctx, mode === 2, reverse)
  if (mode !== 1 && mode !== -1) return VALUE
  // Exact, else the next smaller (-1) or larger (1) item.
  let best = -1
  let bestValue: Scalar = null
  for (let k = 0; k < vec.length; k++) {
    const i = reverse ? vec.length - 1 - k : k
    const v = vec.at(i)
    if (!comparable(v, target)) continue
    const c = compareScalars(v as Exclude<Scalar, CellError>, target as Exclude<Scalar, CellError>)
    if (c === 0) return i
    if ((mode === -1 && c < 0) || (mode === 1 && c > 0)) {
      if (best < 0) {
        best = i
        bestValue = v
        continue
      }
      const d = compareScalars(v as Exclude<Scalar, CellError>, bestValue as Exclude<Scalar, CellError>)
      if ((mode === -1 && d > 0) || (mode === 1 && d < 0)) {
        best = i
        bestValue = v
      }
    }
  }
  return best
}

define("XMATCH", {
  min: 2,
  max: 4,
  params: ["scalar", "any", "scalar", "scalar"],
  call: (a, ctx) => {
    const vec = vectorOf(a[1], ctx)
    if (isError(vec)) return vec
    const mode = a.length > 2 && a[2] !== null ? int(a[2]) : 0
    const search = a.length > 3 && a[3] !== null ? int(a[3]) : 1
    if (isError(mode)) return mode
    if (isError(search)) return search
    const i = xfind(vec, a[0] as Scalar, mode, search, ctx)
    if (isError(i)) return i
    return i < 0 ? NA : i + 1
  },
})

define("XLOOKUP", {
  min: 3,
  max: 6,
  params: ["scalar", "any", "any", "any", "scalar", "scalar"],
  call: (a, ctx) => {
    const vec = vectorOf(a[1], ctx)
    if (isError(vec)) return VALUE
    const mode = a.length > 4 && a[4] !== null ? int(a[4]) : 0
    const search = a.length > 5 && a[5] !== null ? int(a[5]) : 1
    if (isError(mode)) return mode
    if (isError(search)) return search
    const i = xfind(vec, a[0] as Scalar, mode, search, ctx)
    if (isError(i)) return i
    if (i < 0) return a.length > 3 && a[3] !== null ? a[3] : NA
    // Return the matching row (or column) of the return array, which can be wider than one column.
    const ret = a[2]
    if (isRef(ret)) {
      const r = ret.areas[0]
      const area: Area = vec.vertical ? { ...r, r0: r.r0 + i, r1: r.r0 + i } : { ...r, c0: r.c0 + i, c1: r.c0 + i }
      if ((vec.vertical ? r.r1 - r.r0 + 1 : r.c1 - r.c0 + 1) < vecFullLength(a[1])) return VALUE
      return { kind: "ref", areas: [area] }
    }
    const m = toMatrix(ret, ctx)
    if (isError(m)) return m
    if (vec.vertical) return matrix(1, m.cols, m.data.slice(i * m.cols, (i + 1) * m.cols))
    return matrix(
      m.rows,
      1,
      Array.from({ length: m.rows }, (_, r) => m.data[r * m.cols + i]),
    )
  },
})
function vecFullLength(v: Value): number {
  if (isRef(v)) {
    const a = v.areas[0]
    return Math.max(a.r1 - a.r0, a.c1 - a.c0) + 1
  }
  return isMatrix(v) ? Math.max(v.rows, v.cols) : 1
}

// ── INDEX / OFFSET / INDIRECT / ROW / COLUMN / ADDRESS ────────────────────────────────────────────────────────

define("INDEX", {
  min: 2,
  max: 4,
  params: ["any", "scalar", "scalar", "scalar"],
  call: (a, ctx) => {
    const row = a[1] === null ? 0 : int(a[1])
    const col = a.length > 2 && a[2] !== null ? int(a[2]) : 0
    const areaNo = a.length > 3 && a[3] !== null ? int(a[3]) : 1
    if (isError(row)) return row
    if (isError(col)) return col
    if (isError(areaNo)) return areaNo
    const src = a[0]
    if (isRef(src)) {
      if (areaNo < 1 || areaNo > src.areas.length) return REF
      const r = src.areas[areaNo - 1]
      const rows = r.r1 - r.r0 + 1
      const cols = r.c1 - r.c0 + 1
      // INDEX(A1:A10, 3) on a single column or row picks along it.
      let rr = row
      let cc = col
      if (a.length === 2 && rows === 1) {
        cc = row
        rr = 0
      }
      if (rr < 0 || cc < 0 || rr > rows || cc > cols) return REF
      return {
        kind: "ref",
        areas: [
          {
            sheet: r.sheet,
            r0: rr ? r.r0 + rr - 1 : r.r0,
            r1: rr ? r.r0 + rr - 1 : r.r1,
            c0: cc ? r.c0 + cc - 1 : r.c0,
            c1: cc ? r.c0 + cc - 1 : r.c1,
          },
        ],
      }
    }
    const m = toMatrix(src, ctx)
    if (isError(m)) return m
    let rr = row
    let cc = col
    if (a.length === 2 && m.rows === 1) {
      cc = row
      rr = 0
    }
    if (rr < 0 || cc < 0 || rr > m.rows || cc > m.cols) return REF
    if (rr && cc) return m.data[(rr - 1) * m.cols + cc - 1]
    if (rr) return matrix(1, m.cols, m.data.slice((rr - 1) * m.cols, rr * m.cols))
    if (cc)
      return matrix(
        m.rows,
        1,
        Array.from({ length: m.rows }, (_, k) => m.data[k * m.cols + cc - 1]),
      )
    return m
  },
})

define("OFFSET", {
  min: 3,
  max: 5,
  volatile: true,
  params: ["any", "scalar", "scalar", "scalar", "scalar"],
  call: (a, ctx) => {
    if (!isRef(a[0]) || a[0].areas.length !== 1) return VALUE
    const base = a[0].areas[0]
    const dr = int(a[1])
    const dc = int(a[2])
    const h = a.length > 3 && a[3] !== null ? int(a[3]) : base.r1 - base.r0 + 1
    const w = a.length > 4 && a[4] !== null ? int(a[4]) : base.c1 - base.c0 + 1
    for (const v of [dr, dc, h, w]) if (isError(v)) return v
    const [rOff, cOff, height, width] = [dr, dc, h, w] as number[]
    if (height === 0 || width === 0) return REF
    let r0 = base.r0 + rOff
    let c0 = base.c0 + cOff
    let r1 = r0 + height + (height > 0 ? -1 : 1)
    let c1 = c0 + width + (width > 0 ? -1 : 1)
    ;[r0, r1] = [Math.min(r0, r1), Math.max(r0, r1)]
    ;[c0, c1] = [Math.min(c0, c1), Math.max(c0, c1)]
    if (r0 < 0 || c0 < 0 || r1 >= MAX_ROWS || c1 >= MAX_COLS) return REF
    const area = { sheet: base.sheet, r0, c0, r1, c1 }
    ctx.dynamicReference?.(area)
    return { kind: "ref", areas: [area] } satisfies RefValue
  },
})

define("INDIRECT", {
  min: 1,
  max: 2,
  volatile: true,
  params: ["scalar", "scalar"],
  call: (a, ctx) => {
    const text = toText(a[0] as Scalar)
    if (isError(text)) return text
    const a1 = a.length > 1 && a[1] !== null ? toBool(a[1] as Scalar) : true
    if (isError(a1)) return a1
    const refText = a1 ? text.trim() : r1c1ToA1(text.trim(), ctx)
    if (refText === null) return REF
    const parsed = parseFormula(refText)
    if (!parsed.ok) return REF
    const node: Node = parsed.ast
    if (node.type !== "ref" && !(node.type === "binary" && node.op === ":") && node.type !== "name") return REF
    const value = evaluate(node, ctx)
    if (!isRef(value)) return REF
    for (const area of value.areas) ctx.dynamicReference?.(area)
    return value
  },
})
/** R1C1 text for INDIRECT(…, FALSE): relative parts count from the formula cell. */
function r1c1ToA1(text: string, ctx: EvalContext): string | null {
  const bang = text.lastIndexOf("!")
  const prefix = bang >= 0 ? text.slice(0, bang + 1) : ""
  const body = bang >= 0 ? text.slice(bang + 1) : text
  const parts = body.split(":").map((p) => {
    const m = /^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i.exec(p)
    if (!m) return null
    const axis = (s: string | undefined, here: number) =>
      s === undefined ? here : s.startsWith("[") ? here + Number(s.slice(1, -1)) : Number(s) - 1
    const row = axis(m[1], ctx.row)
    const col = axis(m[2], ctx.col)
    if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) return null
    return `${columnName(col)}${row + 1}`
  })
  if (parts.some((p) => p === null)) return null
  return prefix + parts.join(":")
}

define("ROW", {
  min: 0,
  max: 1,
  call: (a, ctx) => {
    if (a.length === 0 || a[0] === null) return ctx.row + 1
    if (!isRef(a[0])) return VALUE
    const r = a[0].areas[0]
    if (r.r0 === r.r1) return r.r0 + 1
    return matrix(
      r.r1 - r.r0 + 1,
      1,
      Array.from({ length: r.r1 - r.r0 + 1 }, (_, k) => r.r0 + k + 1),
    )
  },
})
define("COLUMN", {
  min: 0,
  max: 1,
  call: (a, ctx) => {
    if (a.length === 0 || a[0] === null) return ctx.col + 1
    if (!isRef(a[0])) return VALUE
    const r = a[0].areas[0]
    if (r.c0 === r.c1) return r.c0 + 1
    return matrix(
      1,
      r.c1 - r.c0 + 1,
      Array.from({ length: r.c1 - r.c0 + 1 }, (_, k) => r.c0 + k + 1),
    )
  },
})
define("ROWS", {
  min: 1,
  max: 1,
  call: (a, ctx) => {
    if (isRef(a[0])) return a[0].areas.length === 1 ? a[0].areas[0].r1 - a[0].areas[0].r0 + 1 : REF
    const m = toMatrix(a[0], ctx)
    return isError(m) ? m : m.rows
  },
})
define("COLUMNS", {
  min: 1,
  max: 1,
  call: (a, ctx) => {
    if (isRef(a[0])) return a[0].areas.length === 1 ? a[0].areas[0].c1 - a[0].areas[0].c0 + 1 : REF
    const m = toMatrix(a[0], ctx)
    return isError(m) ? m : m.cols
  },
})
define("AREAS", {
  min: 1,
  max: 1,
  call: (a) => (isRef(a[0]) ? a[0].areas.length : VALUE),
})
define("ADDRESS", {
  min: 2,
  max: 5,
  params: ["scalar"],
  call: (a) => {
    const row = int(a[0])
    const col = int(a[1])
    const absNum = a.length > 2 && a[2] !== null ? int(a[2]) : 1
    const a1 = a.length > 3 && a[3] !== null ? toBool(a[3] as Scalar) : true
    const sheet = a.length > 4 && a[4] !== null ? toText(a[4] as Scalar) : null
    for (const v of [row, col, absNum, a1]) if (isError(v)) return v
    if (sheet !== null && isError(sheet)) return sheet
    const r = row as number
    const c = col as number
    const mode = absNum as number
    if (r < 1 || c < 1 || r > MAX_ROWS || c > MAX_COLS || mode < 1 || mode > 4) return VALUE
    const rowAbs = mode === 1 || mode === 2
    const colAbs = mode === 1 || mode === 3
    const body = a1
      ? `${colAbs ? "$" : ""}${columnName(c - 1)}${rowAbs ? "$" : ""}${r}`
      : `R${rowAbs ? r : `[${r}]`}C${colAbs ? c : `[${c}]`}`
    return (sheet ? sheetPrefix(sheet as string) : "") + body
  },
})

// ── Dynamic arrays ─────────────────────────────────────────────────────────────────────────────────────────────

const arr = (v: Value, ctx: EvalContext) => toMatrix(v, ctx)
const at = (m: Matrix, r: number, c: number) => m.data[r * m.cols + c]
function build(rows: number, cols: number, f: (r: number, c: number) => Scalar): Matrix {
  const data: Scalar[] = new Array(rows * cols)
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) data[r * cols + c] = f(r, c)
  return matrix(rows, cols, data)
}
const blankZero = (v: Scalar) => (v === null ? 0 : v)

define("TRANSPOSE", {
  min: 1,
  max: 1,
  call: (a, ctx) => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    return build(m.cols, m.rows, (r, c) => blankZero(at(m, c, r)))
  },
})

define("FILTER", {
  min: 2,
  max: 3,
  call: (a, ctx) => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const inc = arr(a[1], ctx)
    if (isError(inc)) return inc
    const byRow = inc.cols === 1 && inc.rows === m.rows
    const byCol = inc.rows === 1 && inc.cols === m.cols
    if (!byRow && !byCol) return VALUE
    const keep: number[] = []
    for (let i = 0; i < inc.data.length; i++) {
      const v = inc.data[i]
      if (isError(v)) return v
      const b = toBool(v)
      if (isError(b)) return VALUE
      if (b) keep.push(i)
    }
    if (keep.length === 0) return a.length > 2 && a[2] !== null ? a[2] : CALC
    if (byRow) return build(keep.length, m.cols, (r, c) => blankZero(at(m, keep[r], c)))
    return build(m.rows, keep.length, (r, c) => blankZero(at(m, r, keep[c])))
  },
})

function sortKeys(m: Matrix, byCol: boolean) {
  return byCol ? m.cols : m.rows
}
function compareForSort(x: Scalar, y: Scalar): number {
  // Blanks and errors sort last; otherwise numbers < text < logicals.
  const rank = (v: Scalar) =>
    v === null ? 4 : isError(v) ? 3 : typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2
  const rx = rank(x)
  const ry = rank(y)
  if (rx !== ry) return rx - ry
  if (typeof x === "number" && typeof y === "number") return x - y
  if (typeof x === "string" && typeof y === "string") return compareText(x, y)
  if (typeof x === "boolean" && typeof y === "boolean") return Number(x) - Number(y)
  return 0
}

define("SORT", {
  min: 1,
  max: 4,
  call: (a, ctx) => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const index = a.length > 1 && a[1] !== null ? int(toScalar(a[1], ctx)) : 1
    const order = a.length > 2 && a[2] !== null ? int(toScalar(a[2], ctx)) : 1
    const byCol = a.length > 3 && a[3] !== null ? toBool(toScalar(a[3], ctx)) : false
    for (const v of [index, order, byCol]) if (isError(v)) return v
    if (order !== 1 && order !== -1) return VALUE
    const lines = sortKeys(m, byCol as boolean)
    const width = byCol ? m.rows : m.cols
    if ((index as number) < 1 || (index as number) > width) return VALUE
    const k = (index as number) - 1
    const ids = Array.from({ length: lines }, (_, i) => i)
    const keyOf = (i: number) => (byCol ? at(m, k, i) : at(m, i, k))
    ids.sort((x, y) => compareForSort(keyOf(x), keyOf(y)) * (order as number) || x - y)
    return byCol
      ? build(m.rows, m.cols, (r, c) => blankZero(at(m, r, ids[c])))
      : build(m.rows, m.cols, (r, c) => blankZero(at(m, ids[r], c)))
  },
})

define("SORTBY", {
  min: 2,
  max: 254,
  call: (a, ctx) => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const keys: { m: Matrix; order: number }[] = []
    for (let i = 1; i < a.length; i += 2) {
      const km = arr(a[i], ctx)
      if (isError(km)) return km
      const order = i + 1 < a.length && a[i + 1] !== null ? int(toScalar(a[i + 1], ctx)) : 1
      if (isError(order)) return order
      if (order !== 1 && order !== -1) return VALUE
      keys.push({ m: km, order })
    }
    const byCol = keys[0].m.rows === 1 && keys[0].m.cols === m.cols && m.cols > 1
    const lines = byCol ? m.cols : m.rows
    if (keys.some((k) => k.m.data.length !== lines)) return VALUE
    const ids = Array.from({ length: lines }, (_, i) => i)
    ids.sort((x, y) => {
      for (const k of keys) {
        const c = compareForSort(k.m.data[x], k.m.data[y]) * k.order
        if (c) return c
      }
      return x - y
    })
    return byCol
      ? build(m.rows, m.cols, (r, c) => blankZero(at(m, r, ids[c])))
      : build(m.rows, m.cols, (r, c) => blankZero(at(m, ids[r], c)))
  },
})

define("UNIQUE", {
  min: 1,
  max: 3,
  call: (a, ctx) => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const byCol = a.length > 1 && a[1] !== null ? toBool(toScalar(a[1], ctx)) : false
    const once = a.length > 2 && a[2] !== null ? toBool(toScalar(a[2], ctx)) : false
    if (isError(byCol)) return byCol
    if (isError(once)) return once
    const lines = byCol ? m.cols : m.rows
    const width = byCol ? m.rows : m.cols
    const lineKey = (i: number) =>
      Array.from({ length: width }, (_, j) => {
        const v = byCol ? at(m, j, i) : at(m, i, j)
        return v === null ? "_" : isError(v) ? `e${v.error}` : key(v)
      }).join("\u0000")
    const counts = new Map<string, number>()
    const first: number[] = []
    for (let i = 0; i < lines; i++) {
      const k = lineKey(i)
      if (!counts.has(k)) first.push(i)
      counts.set(k, (counts.get(k) ?? 0) + 1)
    }
    const picked = once ? first.filter((i) => counts.get(lineKey(i)) === 1) : first
    if (!picked.length) return CALC
    return byCol
      ? build(m.rows, picked.length, (r, c) => blankZero(at(m, r, picked[c])))
      : build(picked.length, m.cols, (r, c) => blankZero(at(m, picked[r], c)))
  },
})

function takeDrop(drop: boolean) {
  return (a: Value[], ctx: EvalContext): Value => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const rArg = a.length > 1 && a[1] !== null ? int(toScalar(a[1], ctx)) : null
    const cArg = a.length > 2 && a[2] !== null ? int(toScalar(a[2], ctx)) : null
    if (rArg !== null && isError(rArg)) return rArg
    if (cArg !== null && isError(cArg)) return cArg
    const span = (n: number | null, total: number): [number, number] => {
      if (n === null) return [0, total]
      if (drop) return n >= 0 ? [Math.min(n, total), total] : [0, Math.max(0, total + n)]
      return n >= 0 ? [0, Math.min(n, total)] : [Math.max(0, total + n), total]
    }
    const [r0, r1] = span(rArg as number | null, m.rows)
    const [c0, c1] = span(cArg as number | null, m.cols)
    if (r1 <= r0 || c1 <= c0) return CALC
    return build(r1 - r0, c1 - c0, (r, c) => blankZero(at(m, r0 + r, c0 + c)))
  }
}
define("TAKE", { min: 2, max: 3, call: takeDrop(false) })
define("DROP", { min: 2, max: 3, call: takeDrop(true) })

function stack(horizontal: boolean) {
  return (a: Value[], ctx: EvalContext): Value => {
    const ms: Matrix[] = []
    for (const v of a) {
      const m = arr(v, ctx)
      if (isError(m)) return m
      ms.push(m)
    }
    if (horizontal) {
      const rows = Math.max(...ms.map((m) => m.rows))
      const cols = ms.reduce((s, m) => s + m.cols, 0)
      return build(rows, cols, (r, c) => {
        for (const m of ms) {
          if (c < m.cols) return r < m.rows ? blankZero(at(m, r, c)) : NA
          c -= m.cols
        }
        return NA
      })
    }
    const cols = Math.max(...ms.map((m) => m.cols))
    const rows = ms.reduce((s, m) => s + m.rows, 0)
    return build(rows, cols, (r, c) => {
      for (const m of ms) {
        if (r < m.rows) return c < m.cols ? blankZero(at(m, r, c)) : NA
        r -= m.rows
      }
      return NA
    })
  }
}
define("HSTACK", { min: 1, max: 254, call: stack(true) })
define("VSTACK", { min: 1, max: 254, call: stack(false) })

function choose(rowsMode: boolean) {
  return (a: Value[], ctx: EvalContext): Value => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const picks: number[] = []
    const total = rowsMode ? m.rows : m.cols
    for (const v of a.slice(1)) {
      const list = arr(v, ctx)
      if (isError(list)) return list
      for (const x of list.data) {
        const n = int(x)
        if (isError(n)) return n
        if (n === 0 || Math.abs(n) > total) return VALUE
        picks.push(n > 0 ? n - 1 : total + n)
      }
    }
    return rowsMode
      ? build(picks.length, m.cols, (r, c) => blankZero(at(m, picks[r], c)))
      : build(m.rows, picks.length, (r, c) => blankZero(at(m, r, picks[c])))
  }
}
define("CHOOSEROWS", { min: 2, max: 254, call: choose(true) })
define("CHOOSECOLS", { min: 2, max: 254, call: choose(false) })

function flatten(byColumn: boolean, toRow: boolean) {
  return (a: Value[], ctx: EvalContext): Value => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const ignore = a.length > 1 && a[1] !== null ? int(toScalar(a[1], ctx)) : 0
    const scanCols = a.length > 2 && a[2] !== null ? toBool(toScalar(a[2], ctx)) : byColumn
    if (isError(ignore)) return ignore
    if (isError(scanCols)) return scanCols
    const out: Scalar[] = []
    const visit = (v: Scalar) => {
      if ((ignore === 1 || ignore === 3) && v === null) return
      if ((ignore === 2 || ignore === 3) && isError(v)) return
      out.push(blankZero(v))
    }
    if (scanCols) for (let c = 0; c < m.cols; c++) for (let r = 0; r < m.rows; r++) visit(at(m, r, c))
    else for (const v of m.data) visit(v)
    if (!out.length) return CALC
    return toRow ? matrix(1, out.length, out) : matrix(out.length, 1, out)
  }
}
define("TOCOL", { min: 1, max: 3, call: flatten(false, false) })
define("TOROW", { min: 1, max: 3, call: flatten(false, true) })

function wrap(rowsFirst: boolean) {
  return (a: Value[], ctx: EvalContext): Value => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    if (m.rows !== 1 && m.cols !== 1) return VALUE
    const n = int(toScalar(a[1], ctx))
    if (isError(n)) return n
    if (n < 1) return NUM
    const pad = a.length > 2 && a[2] !== null ? toScalar(a[2], ctx) : NA
    const len = m.data.length
    const other = Math.ceil(len / n)
    return rowsFirst
      ? build(other, n, (r, c) => (r * n + c < len ? blankZero(m.data[r * n + c]) : pad))
      : build(n, other, (r, c) => (c * n + r < len ? blankZero(m.data[c * n + r]) : pad))
  }
}
define("WRAPROWS", { min: 2, max: 3, call: wrap(true) })
define("WRAPCOLS", { min: 2, max: 3, call: wrap(false) })

define("EXPAND", {
  min: 2,
  max: 4,
  call: (a, ctx) => {
    const m = arr(a[0], ctx)
    if (isError(m)) return m
    const rows = a[1] === null ? m.rows : int(toScalar(a[1], ctx))
    const cols = a.length > 2 && a[2] !== null ? int(toScalar(a[2], ctx)) : m.cols
    if (isError(rows)) return rows
    if (isError(cols)) return cols
    if (rows < m.rows || cols < m.cols) return VALUE
    const pad = a.length > 3 ? toScalar(a[3], ctx) : NA
    return build(rows, cols, (r, c) => (r < m.rows && c < m.cols ? blankZero(at(m, r, c)) : pad))
  },
})

// ── LET ──────────────────────────────────────────────────────────────────────────────────────────────────────────

define("LET", {
  min: 3,
  max: 253,
  lazy: (_args, ctx, nodes) => {
    if (nodes.length % 2 === 0) return VALUE
    const bound = new Map(ctx.bound ?? [])
    const inner: EvalContext = Object.create(ctx, { bound: { value: bound } })
    for (let i = 0; i + 1 < nodes.length; i += 2) {
      const n = nodes[i]
      if (n.type !== "name" || n.sheet !== null) return VALUE
      bound.set(n.name.toUpperCase(), evaluate(nodes[i + 1], inner))
    }
    return evaluate(nodes[nodes.length - 1], inner)
  },
})

// ── GETPIVOTDATA ─────────────────────────────────────────────────────────────────────────────────────────────

define("GETPIVOTDATA", {
  min: 2,
  max: 254,
  params: ["scalar", "any", "scalar"],
  call: (a, ctx) => {
    const name = toText(a[0] as Scalar)
    if (isError(name)) return name
    const ref = a[1]
    if (!isRef(ref)) return REF
    const area = ref.areas[0]
    const sheet = area.sheet
    const found = ctx.pivotAt?.(sheet, area.r0, area.c0)
    if (!found || a.length % 2 !== 0) return REF
    const pairs: [string, string | number | boolean][] = []
    for (let i = 2; i < a.length; i += 2) {
      const field = toText(a[i] as Scalar)
      if (isError(field)) return field
      const item = a[i + 1] as Scalar
      if (isError(item)) return item
      pairs.push([field, item === null ? "" : item])
    }
    const cell = pivotDataCell(found.table, found.cache, name, pairs)
    // The formula follows the whole PivotTable: a change to it can move the cell it reads.
    const range = pivotTableRange(found.table, found.cache)
    ctx.dynamicReference?.({ sheet, ...range })
    if (!cell) return REF
    return ctx.cell(sheet, cell.row, cell.col)
  },
})
