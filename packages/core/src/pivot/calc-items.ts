// Calculated items (PivotTable Analyze › Fields, Items & Sets › Calculated Item): an item of a field worked out by a
// formula over the field's other items (=Pens+Ink*2). Measured in Excel 365: the formula works on each item's value
// in every cell (its sum, count, max… as the value field summarizes), an item missing from a cell counting as 0; a
// total adds the calculated items in with the value field's own function (Sum and Count add them up, Max and Min take
// them into account), so a grand total counts a calculated item's records again. Where calculated items of two
// fields meet, the one last in solve order works it out. A file keeps the formula with each item named by its field
// (Product[Pens]+Product[Ink]*2); Cells shows it as typed (=Pens+Ink*2).
import type { Scalar, Value } from "../formula/types"
import type { Node } from "../formula/ast"
import { evaluateFormula } from "../formula/evaluate"
import "../formula/functions"
import { parseFormula } from "../formula/parser"
import { errorValue, type CellValue } from "../values"
import { CALC_FORMULA_ERROR, CALC_FORMULA_REFS, SumsContext } from "./calculated"
import { fieldShared } from "./group"
import type { PivotCache, PivotCalcItem, PivotFn, PivotTable, PivotValue } from "./model"

/** What a formula names: an item (its shared item), the nth item, or the item n places from the calculated one. */
export type CalcItemRef = { kind: "item"; x: number } | { kind: "at"; n: number } | { kind: "rel"; n: number }

export const CALC_ITEM_OTHER_FIELD =
  "PivotTable formulas can only refer to items in the same field as the calculated item."
export const CALC_ITEM_NAME_TAKEN = "The item name already exists in this field."
export const CALC_ITEM_GROUPED = "You cannot add calculated items to grouped fields."
export const CALC_ITEM_NOT_FIELD = "Select an item of a row or column field to add a calculated item to its field."
export const CALC_ITEM_FUNCTIONS =
  "Averages, standard deviations, and variances are not supported when a PivotTable report has calculated items."
export const CALC_ITEM_SAME_FIELD =
  "Multiple data fields of the same field are not supported when a PivotTable report has calculated items."

const placeholder = (k: number) => `_PI${k}_`
const label = (v: PivotValue) => (v === null || v === "" ? "(blank)" : typeof v === "object" ? v.error : String(v))
/** An item named in a formula: bare when it can be, quoted with ' otherwise. */
const named = (text: string) =>
  /^[A-Za-z_\u00a0-\uffff][\w.\u00a0-\uffff]*$/.test(text) ? text : `'${text.replace(/'/g, "''")}'`

/** Reads a quoted name ('Big pens', with '' for a quote) at `i`: the name and where it ends. */
function quoted(formula: string, i: number): { text: string; end: number } {
  let j = i + 1
  let text = ""
  while (j < formula.length && !(formula[j] === "'" && formula[j + 1] !== "'")) {
    text += formula[j]
    j += formula[j] === "'" ? 2 : 1
  }
  return { text, end: j + 1 }
}

/**
 * A calculated item's formula with each item it names swapped for a placeholder, and what each placeholder stands
 * for. Items are named bare (Pens), quoted ('Big pens'), with their field (Product[Pens], Product['Big pens']) or by
 * place (Product[2], Product[-1]). `problem`: Excel's message when the formula names another field.
 */
export function calcItemSource(
  cache: PivotCache,
  f: number,
  formula: string,
): { text: string; refs: CalcItemRef[]; problem: string | null } {
  const items = fieldShared(cache, f)
  const byLabel = new Map<string, number>()
  items.forEach((v, x) => byLabel.has(label(v).toLowerCase()) || byLabel.set(label(v).toLowerCase(), x))
  const fieldName = cache.fields[f]?.name.toLowerCase() ?? ""
  const refs: CalcItemRef[] = []
  const keys = new Map<string, number>()
  let problem: string | null = null
  const ref = (r: CalcItemRef) => {
    const key = JSON.stringify(r)
    let k = keys.get(key)
    if (k === undefined) keys.set(key, (k = refs.push(r) - 1))
    return placeholder(k)
  }
  /** The bracket after a field's name: an item, or a place. */
  const bracket = (start: number): { out: string; end: number } => {
    const close = formula.indexOf("]", start)
    if (close < 0) return { out: formula.slice(start - 1), end: formula.length }
    let inner = formula.slice(start, close).trim()
    if (inner.startsWith("'")) inner = quoted(inner, 0).text
    else if (/^[+-]?\d+$/.test(inner))
      return {
        out: ref(/^[+-]/.test(inner) ? { kind: "rel", n: Number(inner) } : { kind: "at", n: Number(inner) }),
        end: close + 1,
      }
    const x = byLabel.get(inner.toLowerCase())
    return { out: x === undefined ? "_NOITEM_" : ref({ kind: "item", x }), end: close + 1 }
  }
  let out = ""
  let i = formula.trimStart().startsWith("=") ? formula.indexOf("=") + 1 : 0
  while (i < formula.length) {
    const ch = formula[i]
    if (ch === '"') {
      let j = i + 1
      while (j < formula.length && !(formula[j] === '"' && formula[j + 1] !== '"')) j += formula[j] === '"' ? 2 : 1
      out += formula.slice(i, j + 1)
      i = j + 1
    } else if (ch === "'") {
      const q = quoted(formula, i)
      if (formula[q.end] === "[") {
        if (q.text.toLowerCase() !== fieldName) problem ??= CALC_ITEM_OTHER_FIELD
        const b = bracket(q.end + 1)
        out += b.out
        i = b.end
      } else {
        const x = byLabel.get(q.text.toLowerCase())
        out += x === undefined ? "_NOITEM_" : ref({ kind: "item", x })
        i = q.end
      }
    } else if (/[A-Za-z_\u00a0-\uffff]/.test(ch)) {
      const word = /^[\w.\u00a0-\uffff]+/.exec(formula.slice(i))![0]
      const after = i + word.length
      if (/^\s*\(/.test(formula.slice(after))) {
        out += word
        i = after
      } else if (formula[after] === "[") {
        if (word.toLowerCase() !== fieldName) problem ??= CALC_ITEM_OTHER_FIELD
        const b = bracket(after + 1)
        out += b.out
        i = b.end
      } else {
        const x = byLabel.get(word.toLowerCase())
        // TRUE and FALSE stay themselves unless an item has that name.
        out += x !== undefined ? ref({ kind: "item", x }) : word
        i = after
      }
    } else {
      out += ch
      i++
    }
  }
  return { text: out, refs, problem }
}

type Parsed = { node: Node | null; refs: CalcItemRef[] }
const parsedCache = new WeakMap<PivotCache, Map<number, Parsed>>()

/** Calculated item `k` of the cache (its index in `calcItems`), parsed once. */
function parsedItem(cache: PivotCache, k: number): Parsed {
  let map = parsedCache.get(cache)
  if (!map) parsedCache.set(cache, (map = new Map()))
  let p = map.get(k)
  if (!p) {
    const ci = cache.calcItems?.[k]
    const src = ci ? calcItemSource(cache, ci.field, ci.formula) : null
    const result = src && !src.problem ? parseFormula(src.text) : null
    p = { node: result?.ok ? result.ast : null, refs: src?.refs ?? [] }
    map.set(k, p)
  }
  return p
}

/**
 * Calculated item `k` in one cell: its formula with each item it names replaced by `valueOf` that item (a number, 0
 * for none, or an error); a reference it can't follow is #REF!.
 */
export function calcItemValue(
  cache: PivotCache,
  k: number,
  valueOf: (ref: CalcItemRef) => CellValue | undefined,
): CellValue {
  const p = parsedItem(cache, k)
  if (!p.node) return errorValue("#NAME?")
  const bound = new Map<string, Value>()
  p.refs.forEach((r, i) => {
    const v = valueOf(r)
    bound.set(placeholder(i).toUpperCase(), (v === undefined || v === null ? 0 : v) as Scalar)
  })
  const v = evaluateFormula(p.node, new SumsContext(bound))
  if (v !== null && typeof v === "object" && "kind" in v) {
    if (v.kind === "matrix") return (v.data[0] ?? 0) as CellValue
    return errorValue("#VALUE!")
  }
  return (v ?? 0) as CellValue
}

/** A calculated item's formula as Cells shows it (=Pens+Ink*2): items without their field. */
export function calcItemDisplay(cache: PivotCache, ci: PivotCalcItem): string {
  return `=${shorten(cache, ci.field, ci.formula)}`
}

/** Rewrites each item reference with `write`, leaving strings, function names and the rest as they are. */
function rewrite(cache: PivotCache, f: number, formula: string, write: (r: CalcItemRef) => string): string {
  const src = calcItemSource(cache, f, formula)
  return src.text.replace(/_PI(\d+)_/g, (_, k: string) => write(src.refs[Number(k)]))
}

function shorten(cache: PivotCache, f: number, formula: string): string {
  const items = fieldShared(cache, f)
  const field = named(cache.fields[f]?.name ?? "")
  return rewrite(cache, f, formula, (r) =>
    r.kind === "item" ? named(label(items[r.x])) : `${field}[${r.kind === "rel" && r.n >= 0 ? "+" : ""}${r.n}]`,
  )
}

/** A formula as typed (=Pens+Ink*2) in the form a file keeps (Product[Pens]+Product[Ink]*2, no "="). */
export function calcItemStandard(cache: PivotCache, f: number, typed: string): string {
  const items = fieldShared(cache, f)
  const field = named(cache.fields[f]?.name ?? "")
  return rewrite(cache, f, typed, (r) =>
    r.kind === "item"
      ? `${field}[${named(label(items[r.x]))}]`
      : `${field}[${r.kind === "rel" && r.n >= 0 ? "+" : ""}${r.n}]`,
  ).trim()
}

/** Why `formula` can't be a calculated item of field `f`, or null when it can. */
export function calcItemProblem(cache: PivotCache, f: number, formula: string): string | null {
  const src = calcItemSource(cache, f, formula)
  if (src.problem) return src.problem
  if (!src.text.trim() || src.text.includes("_NOITEM_")) return CALC_FORMULA_ERROR
  const result = parseFormula(src.text)
  if (!result.ok) return CALC_FORMULA_ERROR
  return outside(result.ast) ? CALC_FORMULA_REFS : null
}

/** What a calculated item's formula may not hold: cells, names that aren't items, arrays, table references. */
function outside(node: Node): boolean {
  switch (node.type) {
    case "ref":
    case "array":
    case "spill":
    case "structured":
    case "call":
      return true
    case "name":
      return node.sheet !== null || !/^_PI\d+_$/.test(node.name)
    case "func":
      return node.args.some(outside)
    case "unary":
    case "percent":
    case "paren":
      return outside(node.arg)
    case "binary":
      return outside(node.left) || outside(node.right)
    default:
      return false
  }
}

/** Whether field `f` can take calculated items: a plain field (not grouped, not a calculated field). */
export function canHaveCalcItems(cache: PivotCache, f: number): boolean {
  const field = cache.fields[f]
  return Boolean(field && !field.group && field.formula === undefined && !cache.fields.some((g) => g.group?.base === f))
}

/**
 * The cache with calculated item `name` of field `f` set to `formula` (Calculated Item's Add, or Modify for one it
 * has), or Excel's message. A new item goes after the field's other items and last in solve order.
 */
export function withCalcItem(cache: PivotCache, f: number, name: string, formula: string): PivotCache | string {
  const clean = name.trim()
  const field = cache.fields[f]
  if (!field || !clean) return CALC_FORMULA_ERROR
  if (!canHaveCalcItems(cache, f)) return CALC_ITEM_GROUPED
  const problem = calcItemProblem(cache, f, formula)
  if (problem) return problem
  const standard = calcItemStandard(cache, f, formula)
  const items = fieldShared(cache, f)
  const x = items.findIndex((v) => label(v).toLowerCase() === clean.toLowerCase())
  const list = cache.calcItems ?? []
  const own = list.findIndex((ci) => ci.field === f && ci.x === x)
  if (x >= 0 && own < 0) return CALC_ITEM_NAME_TAKEN
  if (own >= 0) return { ...cache, calcItems: list.map((ci, i) => (i === own ? { ...ci, formula: standard } : ci)) }
  const fields = [...cache.fields]
  fields[f] = { ...field, items: [...items, clean] }
  return { ...cache, fields, calcItems: [...list, { field: f, x: items.length, formula: standard }] }
}

/**
 * The cache without calculated item `x` of field `f`: its shared item goes too (the items after it move up one,
 * as `shift` tells the PivotTables).
 */
export function withoutCalcItem(cache: PivotCache, f: number, x: number): PivotCache {
  const field = cache.fields[f]
  if (!field?.items || !cache.calcItems?.some((ci) => ci.field === f && ci.x === x)) return cache
  const fields = [...cache.fields]
  fields[f] = { ...field, items: field.items.filter((_, i) => i !== x) }
  const calcItems = cache.calcItems
    .filter((ci) => !(ci.field === f && ci.x === x))
    .map((ci) => (ci.field === f && ci.x > x ? { ...ci, x: ci.x - 1 } : ci))
  return { ...cache, fields, ...(calcItems.length ? { calcItems } : { calcItems: undefined }) }
}

/** A PivotTable after shared item `x` of field `f` left its cache (withoutCalcItem): its items follow. */
export function tableWithoutItem(table: PivotTable, f: number, x: number): PivotTable {
  const field = table.fields[f]
  if (!field) return table
  const pos = field.items.findIndex((i) => i.x === x)
  const items = field.items.filter((i) => i.x !== x).map((i) => (i.x > x ? { ...i, x: i.x - 1 } : i))
  const fields = [...table.fields]
  fields[f] = { ...field, items }
  const pages = table.pages.map((p) =>
    p.field !== f || p.item === null || pos < 0
      ? p
      : { ...p, item: p.item === pos ? null : p.item > pos ? p.item - 1 : p.item },
  )
  return { ...table, fields, pages }
}

/** The functions a value field may summarize by in a PivotTable with calculated items (Excel's). */
export const CALC_ITEM_FNS: readonly PivotFn[] = ["sum", "count", "max", "min", "product", "countNums"]

/** Excel's message when the PivotTables on a cache can't have calculated items (an Average, a field summed twice). */
export function calcItemsBlocked(tables: readonly PivotTable[]): string | null {
  for (const t of tables) {
    if (t.values.some((v) => !CALC_ITEM_FNS.includes(v.fn))) return CALC_ITEM_FUNCTIONS
    const fields = t.values.map((v) => v.field)
    if (new Set(fields).size !== fields.length) return CALC_ITEM_SAME_FIELD
  }
  return null
}
