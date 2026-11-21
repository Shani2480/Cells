/**
 * The database functions (DSUM, DAVERAGE, DCOUNT, DGET…): a list with headings in its first row, a field (a heading or
 * a column number) and a criteria range laid out as for Advanced Filter: headings naming the list's columns, each row
 * under them one set of conditions that must all hold, a record counting when any set holds. A formula under a heading
 * that names no column is worked out for each record, written for the list's first record.
 */
import { isError, type CellError } from "../../values"
import { copyFormula } from "../adjust"
import { DIV0, NUM, VALUE, toText } from "../coerce"
import type { EvalContext } from "../context"
import { evaluateFormula } from "../evaluate"
import { parseFormula } from "../parser"
import { isLambda, isMatrix, isRef, type Matrix, type Scalar, type Value } from "../types"
import { advancedCriterion } from "./criteria"
import { arrayArg } from "./helpers"
import { define } from "./registry"
import { variance } from "./stats"

type Condition = { col: number; test: (v: Scalar) => boolean } | { formula: string; sheet: number }

function headingKey(v: Scalar): string {
  if (v === null) return ""
  const t = isError(v) ? v.error : toText(v)
  return (isError(t) ? "" : t).toLowerCase()
}

/** The field's column in the database (0-based), or null for an omitted field. */
function fieldColumn(field: Value, db: Matrix, ctx: EvalContext): number | null | CellError {
  if (field === null) return null
  if (isRef(field)) {
    const a = field.areas[0]
    if (field.areas.length !== 1 || a.r0 !== a.r1 || a.c0 !== a.c1) return VALUE
    field = ctx.cell(a.sheet, a.r0, a.c0)
  }
  if (isMatrix(field) || isLambda(field)) return VALUE
  if (isError(field)) return field
  // TRUE is column 1.
  if (typeof field === "boolean") field = field ? 1 : 0
  if (typeof field === "number") {
    const i = Math.trunc(field)
    return i >= 1 && i <= db.cols ? i - 1 : VALUE
  }
  if (field === null) return VALUE
  const key = field.toLowerCase()
  for (let c = 0; c < db.cols; c++) if (headingKey(db.data[c]) === key) return c
  return VALUE
}

/** The values of the field in the records that meet the criteria (every record's for an omitted field: `true`). */
function matching(args: Value[], ctx: EvalContext, needField: boolean): (Scalar | true)[] | CellError {
  const db = arrayArg(args[0], ctx)
  if (isError(db)) return db
  const field = fieldColumn(args[1] ?? null, db, ctx)
  if (isError(field)) return field
  if (field === null && needField) return VALUE
  const critArg = args[2]
  if (!isRef(critArg) || critArg.areas.length !== 1) return VALUE
  const area = critArg.areas[0]
  const crit = arrayArg(critArg, ctx)
  if (isError(crit)) return crit
  if (crit.rows < 2) return VALUE
  const dbArea = isRef(args[0]) && args[0].areas.length === 1 ? args[0].areas[0] : null

  const columns = new Map<string, number>()
  for (let c = 0; c < db.cols; c++) {
    const key = headingKey(db.data[c])
    if (key && !columns.has(key)) columns.set(key, c)
  }
  const sets: Condition[][] = []
  for (let r = 1; r < crit.rows; r++) {
    const set: Condition[] = []
    for (let c = 0; c < crit.cols; c++) {
      const col = columns.get(headingKey(crit.data[c]))
      const value = crit.data[r * crit.cols + c]
      if (col === undefined) {
        const text = ctx.formulaText(area.sheet, area.r0 + r, area.c0 + c)
        if (text && dbArea) set.push({ formula: text.startsWith("=") ? text : "=" + text, sheet: area.sheet })
        continue
      }
      const test = value === null ? null : advancedCriterion(value)
      if (test) set.push({ col, test })
    }
    sets.push(set)
  }

  const holds = (cond: Condition, record: number): boolean => {
    if ("col" in cond) return cond.test(db.data[record * db.cols + cond.col])
    const at = dbArea!
    const parsed = parseFormula(copyFormula(cond.formula, record - 1, 0).slice(1))
    if (!parsed.ok) return false
    // Worked out as if written in the database's first record row, moved down to this record.
    const here: EvalContext = Object.create(ctx, {
      sheet: { value: at.sheet },
      row: { value: at.r0 + record },
      col: { value: at.c0 },
    })
    let v = evaluateFormula(parsed.ast, here)
    if (isRef(v)) v = ctx.cell(v.areas[0].sheet, v.areas[0].r0, v.areas[0].c0)
    if (isMatrix(v)) v = v.data[0] ?? null
    return v === true || (typeof v === "number" && v !== 0)
  }

  const out: (Scalar | true)[] = []
  for (let record = 1; record < db.rows; record++) {
    if (!sets.some((set) => set.every((cond) => holds(cond, record)))) continue
    out.push(field === null ? true : db.data[record * db.cols + field])
  }
  return out
}

function numbersOf(values: (Scalar | true)[]): number[] {
  return values.filter((v): v is number => typeof v === "number")
}

function databaseFn(name: string, f: (values: (Scalar | true)[]) => Value, needField = true) {
  define(name, {
    min: 3,
    max: 3,
    call: (args, ctx) => {
      // Criteria must be a range: an array gives #VALUE! for each of its elements, as Excel lifts the call over it.
      const crit = args[2]
      if (isMatrix(crit)) return { ...crit, data: crit.data.map(() => VALUE) }
      const values = matching(args, ctx, needField)
      return isError(values) ? values : f(values)
    },
  })
}

const sum = (xs: number[]) => xs.reduce((s, x) => s + x, 0)

databaseFn("DSUM", (v) => sum(numbersOf(v)))
databaseFn("DAVERAGE", (v) => {
  const xs = numbersOf(v)
  return xs.length ? sum(xs) / xs.length : DIV0
})
databaseFn("DCOUNT", (v) => v.filter((x) => x === true || typeof x === "number").length, false)
databaseFn("DCOUNTA", (v) => v.filter((x) => x === true || (x !== null && x !== "")).length, false)
databaseFn("DMAX", (v) => {
  const xs = numbersOf(v)
  return xs.length ? Math.max(...xs) : 0
})
databaseFn("DMIN", (v) => {
  const xs = numbersOf(v)
  return xs.length ? Math.min(...xs) : 0
})
databaseFn("DPRODUCT", (v) => {
  const xs = numbersOf(v)
  return xs.length ? xs.reduce((p, x) => p * x, 1) : 0
})
databaseFn("DGET", (v) => {
  if (v.length === 0) return VALUE
  if (v.length > 1) return NUM
  return v[0] === true ? VALUE : v[0]
})
function deviations(xs: number[]): Value {
  if (!xs.length) return DIV0
  const mean = sum(xs) / xs.length
  return xs.reduce((s, x) => s + (x - mean) * (x - mean), 0) / xs.length
}

for (const [n, population, root] of [
  ["DVAR", false, false],
  ["DVARP", true, false],
  ["DSTDEV", false, true],
  ["DSTDEVP", true, true],
] as const) {
  databaseFn(n, (v) => {
    const xs = numbersOf(v)
    // Excel's DVARP (only) takes the squares about the mean; the others use its Σx² − (Σx)²/n, like VAR.
    const r = population && !root ? deviations(xs) : variance(xs, !population)
    return typeof r === "number" && root ? Math.sqrt(r) : r
  })
}
