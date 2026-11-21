// Calculated fields (Analyze › Fields, Items & Sets › Calculated Field): a cache field with a formula over the other
// fields. Excel 365 works it out on the sums of the fields it names, in every value cell, subtotal and grand total,
// whatever the value field's function; a cell with no records still works it out, on zeros (Amount/Units there is
// #DIV/0!). Checked against io's fixtures/excel/pivot-calc.xlsx.
import type { Area, Scalar, Value } from "../formula/types"
import type { EvalContext } from "../formula/context"
import type { Node } from "../formula/ast"
import { evaluateFormula } from "../formula/evaluate"
// The functions a formula may call (they register themselves).
import "../formula/functions"
import { parseFormula } from "../formula/parser"
import { errorValue, type CellValue } from "../values"
import type { PivotCache } from "./model"

/** The placeholder name a field reference becomes before the formula is parsed (field names needn't be names). */
const placeholder = (f: number) => `_PF${f}_`

/**
 * A calculated field's formula with every field it names (bare, or quoted as 'Unit Cost') swapped for a
 * placeholder, and the fields it names. Text in double quotes and function names are left alone.
 */
export function calcFieldSource(cache: PivotCache, formula: string): { text: string; fields: number[] } {
  const byName = new Map(cache.fields.map((f, i) => [f.name.toLowerCase(), i]))
  const fields = new Set<number>()
  let out = ""
  let i = formula.startsWith("=") ? 1 : 0
  while (i < formula.length) {
    const ch = formula[i]
    if (ch === '"') {
      let j = i + 1
      while (j < formula.length && !(formula[j] === '"' && formula[j + 1] !== '"')) j += formula[j] === '"' ? 2 : 1
      out += formula.slice(i, j + 1)
      i = j + 1
    } else if (ch === "'") {
      let j = i + 1
      let name = ""
      while (j < formula.length && !(formula[j] === "'" && formula[j + 1] !== "'")) {
        name += formula[j]
        j += formula[j] === "'" ? 2 : 1
      }
      const f = byName.get(name.toLowerCase())
      if (f !== undefined) fields.add(f)
      out += f !== undefined ? placeholder(f) : formula.slice(i, j + 1)
      i = j + 1
    } else if (/[A-Za-z_\u00a0-\uffff]/.test(ch)) {
      const m = /^[\w.\u00a0-\uffff]+/.exec(formula.slice(i))!
      const word = m[0]
      const call = /^\s*\(/.test(formula.slice(i + word.length))
      const f = call ? undefined : byName.get(word.toLowerCase())
      if (f !== undefined) fields.add(f)
      out += f !== undefined ? placeholder(f) : word
      i += word.length
    } else {
      out += ch
      i++
    }
  }
  return { text: out, fields: [...fields] }
}

type Parsed = { node: Node | null; fields: number[] }
const parsedCache = new WeakMap<PivotCache, Map<number, Parsed>>()

function parsed(cache: PivotCache, f: number): Parsed {
  let map = parsedCache.get(cache)
  if (!map) parsedCache.set(cache, (map = new Map()))
  let p = map.get(f)
  if (!p) {
    const src = calcFieldSource(cache, cache.fields[f]?.formula ?? "")
    const result = parseFormula(src.text)
    p = { node: result.ok ? result.ast : null, fields: src.fields }
    map.set(f, p)
  }
  return p
}

/** The source fields a value field reads: itself, or every field its formula names (through other formulas too). */
export function calcFieldInputs(cache: PivotCache, f: number, seen = new Set<number>()): number[] {
  if (!cache.fields[f]?.formula) return [f]
  if (seen.has(f)) return []
  seen.add(f)
  return [...new Set(parsed(cache, f).fields.flatMap((g) => calcFieldInputs(cache, g, seen)))]
}

/** Only the names bound to the sums: a calculated field's formula can't read cells. */
class SumsContext implements EvalContext {
  sheet = 0
  row = 0
  col = 0
  constructor(readonly bound: Map<string, Value>) {}
  cell(): Scalar {
    return null
  }
  values(_area: Area, _visit: (value: Scalar, row: number, col: number) => void) {}
  usedArea() {
    return null
  }
  sheetIndex() {
    return -1
  }
  sheetName() {
    return ""
  }
  sheetCount() {
    return 0
  }
  defined() {
    return null
  }
  spillArea() {
    return null
  }
  isFormula() {
    return false
  }
  formulaText() {
    return null
  }
  now() {
    return 0
  }
  random() {
    return 0
  }
}

/**
 * Calculated field `f` for one group of records: its formula with each field it names replaced by `sumOf` that
 * field (0 for none), another calculated field by its own value.
 */
export function calcFieldValue(cache: PivotCache, f: number, sumOf: (field: number) => number, depth = 0): CellValue {
  const p = parsed(cache, f)
  if (!p.node || depth > 32) return errorValue("#NAME?")
  const bound = new Map<string, Value>()
  for (const g of p.fields) {
    const v = cache.fields[g]?.formula ? calcFieldValue(cache, g, sumOf, depth + 1) : sumOf(g)
    bound.set(placeholder(g).toUpperCase(), v as Scalar)
  }
  const v = evaluateFormula(p.node, new SumsContext(bound))
  if (v !== null && typeof v === "object" && "kind" in v) {
    if (v.kind === "matrix") return (v.data[0] ?? 0) as CellValue
    return errorValue("#VALUE!")
  }
  return (v ?? 0) as CellValue
}

/** Excel's messages for a calculated field it won't take. */
export const CALC_FORMULA_ERROR = "There's a problem with this formula."
export const CALC_FORMULA_REFS = "References, names, and arrays are not supported in PivotTable formulas."
export const CALC_NAME_TAKEN = "PivotTable field name already exists."

/** What a calculated field's formula may not hold: cells, names other than fields, arrays, table references. */
function outsideRef(node: Node): boolean {
  switch (node.type) {
    case "ref":
    case "array":
    case "spill":
    case "structured":
      return true
    case "name":
      return node.sheet !== null || !/^_PF\d+_$/.test(node.name)
    case "func":
      return node.args.some(outsideRef)
    case "call":
      return true
    case "unary":
    case "percent":
    case "paren":
      return outsideRef(node.arg)
    case "binary":
      return outsideRef(node.left) || outsideRef(node.right)
    default:
      return false
  }
}

/** Why a formula can't be a calculated field of `cache`, or null when it can. */
export function calcFormulaProblem(cache: PivotCache, formula: string): string | null {
  const src = calcFieldSource(cache, formula)
  if (!src.text.trim()) return CALC_FORMULA_ERROR
  const result = parseFormula(src.text)
  if (!result.ok) return CALC_FORMULA_ERROR
  return outsideRef(result.ast) ? CALC_FORMULA_REFS : null
}

/**
 * The cache with calculated field `name` set to `formula` (Calculated Field's Add, or Modify for a name it has), or
 * Excel's message. The formula is kept without its "=", as a file keeps it.
 */
export function withCalcField(cache: PivotCache, name: string, formula: string): PivotCache | string {
  const clean = name.trim()
  if (!clean) return CALC_FORMULA_ERROR
  const problem = calcFormulaProblem(cache, formula)
  if (problem) return problem
  const text = formula.trim().replace(/^=/, "").trim()
  const at = cache.fields.findIndex((f) => f.name.toLowerCase() === clean.toLowerCase())
  if (at >= 0 && cache.fields[at].formula === undefined) return CALC_NAME_TAKEN
  const fields = [...cache.fields]
  if (at >= 0) fields[at] = { ...fields[at], formula: text }
  else fields.push({ name: clean, items: null, formula: text })
  return { ...cache, fields }
}

/** The cache without calculated field `name`. */
export function withoutCalcField(cache: PivotCache, name: string): PivotCache {
  const lower = name.toLowerCase()
  return { ...cache, fields: cache.fields.filter((f) => f.formula === undefined || f.name.toLowerCase() !== lower) }
}
