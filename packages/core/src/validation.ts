import type { Range } from "./address"
import { shiftRange, subtract, ValueContext, type CellIsOperator } from "./conditional"
import { copyFormula } from "./formula/adjust"
import { compareText } from "./formula/coerce"
import { evaluateFormula } from "./formula/evaluate"
import { parseFormula } from "./formula/parser"
import { isMatrix, isRef, type Area, type Scalar, type Value } from "./formula/types"
import { displayText } from "./format"
import type { Sheet } from "./sheet"
import { completeFormula } from "./formula/editing"
import { isError, numberToEditText, parseInput, type CellValue } from "./values"

/**
 * Data Validation, as Excel models it: a rule per cell (over ranges), checked when a value is typed, with an input
 * message shown when the cell is selected and an error alert when the value breaks the rule.
 */

export type ValidationType = "any" | "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom"
export type ValidationOperator = CellIsOperator
export type ErrorStyle = "stop" | "warning" | "information"

export type Validation = {
  type: ValidationType
  operator: ValidationOperator
  /** Formulas as the file holds them, without "=": "1", "$H$1:$H$3", "\"a,b,c\"" (a typed list), "46082". */
  formula1: string
  formula2: string
  /** Ignore blank. */
  allowBlank: boolean
  /** In-cell dropdown (lists). */
  dropDown: boolean
  showInput: boolean
  promptTitle: string
  prompt: string
  showError: boolean
  errorStyle: ErrorStyle
  errorTitle: string
  error: string
}

export type ValidationEntry = { ranges: Range[]; rule: Validation }

export const DEFAULT_VALIDATION: Validation = {
  type: "any",
  operator: "between",
  formula1: "",
  formula2: "",
  allowBlank: true,
  dropDown: true,
  showInput: true,
  promptTitle: "",
  prompt: "",
  showError: true,
  errorStyle: "stop",
  errorTitle: "",
  error: "",
}

/** Excel's message when an error alert has none of its own. */
export const DEFAULT_VALIDATION_ERROR =
  "This value doesn't match the data validation restrictions defined for this cell."

/** Whether the operator takes two values (Minimum and Maximum). */
export const twoValues = (op: ValidationOperator) => op === "between" || op === "notBetween"

/** Whether the rule's type compares with an operator (not Any value, List or Custom). */
export const usesOperator = (type: ValidationType) =>
  type === "whole" || type === "decimal" || type === "date" || type === "time" || type === "textLength"

// ── Editing ───────────────────────────────────────────────────────────────────────────────────────────────────

export function copyValidations(entries: readonly ValidationEntry[]): ValidationEntry[] {
  return entries.map((e) => ({ ranges: e.ranges.map((r) => ({ ...r })), rule: { ...e.rule } }))
}

const inside = (r: Range, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1
const overlaps = (a: Range, b: Range) => a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/** The rule of a cell, or null. */
export function validationAt(sheet: Sheet, row: number, col: number): ValidationEntry | null {
  for (let i = sheet.validations.length - 1; i >= 0; i--) {
    const e = sheet.validations[i]
    if (e.ranges.some((r) => inside(r, row, col))) return e
  }
  return null
}

/** The rules that touch any of `ranges`. */
export function validationsIn(sheet: Sheet, ranges: readonly Range[]): ValidationEntry[] {
  return sheet.validations.filter((e) => e.ranges.some((r) => ranges.some((x) => overlaps(r, x))))
}

/**
 * The sheet's rules after `ranges` get `rule` (null: Clear All). The cells leave any rule they had; a rule keeps the
 * rest of its cells.
 */
export function withValidation(
  entries: readonly ValidationEntry[],
  ranges: readonly Range[],
  rule: Validation | null,
): ValidationEntry[] {
  const out: ValidationEntry[] = []
  for (const e of entries) {
    let parts = e.ranges.map((r) => ({ ...r }))
    for (const cut of ranges) parts = parts.flatMap((p) => subtract(p, cut))
    if (parts.length) out.push({ ranges: parts, rule: { ...e.rule } })
  }
  if (rule) out.push({ ranges: ranges.map((r) => ({ ...r })), rule: { ...rule } })
  return out
}

/** Rules follow inserted and deleted rows and columns; a rule whose cells all go is dropped. */
export function shiftValidations(sheet: Sheet, axis: "rows" | "cols", at: number, count: number) {
  if (!sheet.validations.length) return
  sheet.validations = sheet.validations
    .map((e) => ({ ...e, ranges: e.ranges.map((r) => shiftRange(r, axis, at, count)).filter((r) => r !== null) }))
    .filter((e) => e.ranges.length > 0)
}

/** Two rules alike (Apply these changes to all other cells with the same settings). */
export function sameValidation(a: Validation, b: Validation): boolean {
  return (Object.keys(a) as (keyof Validation)[]).every((k) => a[k] === b[k])
}

// ── Checking ──────────────────────────────────────────────────────────────────────────────────────────────────

/** Reads the workbook, with the value being typed in its cell (the rule checks it before it is entered). */
class PendingContext extends ValueContext {
  pending: { sheet: number; row: number; col: number; value: Scalar } | null = null
  override cell(sheet: number, row: number, col: number): Scalar {
    const p = this.pending
    if (p && p.sheet === sheet && p.row === row && p.col === col) return p.value
    return super.cell(sheet, row, col)
  }
  override values(area: Area, visit: (value: Scalar, row: number, col: number) => void) {
    const p = this.pending
    if (!p || p.sheet !== area.sheet || p.row < area.r0 || p.row > area.r1 || p.col < area.c0 || p.col > area.c1) {
      super.values(area, visit)
      return
    }
    super.values(area, (v, r, c) => (r === p.row && c === p.col ? undefined : visit(v, r, c)))
    if (p.value !== null) visit(p.value, p.row, p.col)
  }
}

export class ValidationChecker {
  private readonly ctx: PendingContext

  constructor(
    private readonly sheets: readonly Sheet[],
    private readonly sheetIndex: number,
    today: number,
  ) {
    this.ctx = new PendingContext(sheets, today)
  }

  private evaluate(formula: string, e: ValidationEntry, row: number, col: number): Value {
    const origin = e.ranges[0]
    const text = copyFormula("=" + formula, row - origin.r0, col - origin.c0)
    const parsed = parseFormula(text.slice(1))
    if (!parsed.ok) return { error: "#NAME?" }
    this.ctx.sheet = this.sheetIndex
    this.ctx.row = row
    this.ctx.col = col
    return evaluateFormula(parsed.ast, this.ctx)
  }

  private scalar(formula: string, e: ValidationEntry, row: number, col: number): Scalar {
    const v = this.evaluate(formula, e, row, col)
    if (isRef(v)) {
      const a = v.areas[0]
      return this.ctx.cell(a.sheet, a.r0, a.c0)
    }
    if (isMatrix(v)) return v.data[0] ?? null
    return v
  }

  /**
   * A list's items: a typed list ("a, b,c", trimmed) or the cells of a range or name. `typed` says which, as the two
   * match differently.
   */
  listItems(e: ValidationEntry, row: number, col: number): { items: string[]; values: CellValue[]; typed: boolean } {
    const f = e.rule.formula1.trim()
    if (f.startsWith('"')) {
      const items = f
        .slice(1, f.endsWith('"') ? -1 : undefined)
        .replace(/""/g, '"')
        .split(",")
        .map((s) => s.trim())
      return { items, values: items, typed: true }
    }
    const v = this.evaluate(f, e, row, col)
    const values: CellValue[] = []
    const items: string[] = []
    const push = (value: CellValue | undefined, text: string) => {
      values.push(value ?? "")
      items.push(text)
    }
    if (isRef(v)) {
      const a = v.areas[0]
      const sheet = this.sheets[a.sheet]
      if (sheet && (a.r0 === a.r1 || a.c0 === a.c1)) {
        // Only the used part of a whole column (A:A) gives items.
        const used = sheet.cells.usedBounds()
        const r1 = used ? Math.min(a.r1, used.r1) : a.r0 - 1
        const c1 = used ? Math.min(a.c1, used.c1) : a.c0 - 1
        for (let r = a.r0; r <= r1; r++)
          for (let c = a.c0; c <= c1; c++) {
            const d = sheet.cells.get(r, c)
            const fmt = sheet.styleAt(r, c)
            push(d?.value, displayText(d?.value, this.formatOf?.(fmt)))
          }
      }
    } else if (isMatrix(v)) for (const x of v.data) push(x ?? "", x === null ? "" : scalarText(x))
    return { items, values, typed: false }
  }

  /** The number format of a style id (set by the caller, for the dropdown's texts). */
  formatOf?: (style: number) => string | undefined

  /** The value typed text gives a cell: the text as entered, or what its formula works out to there. */
  entryValue(text: string, row: number, col: number): CellValue | undefined {
    const parsed = parseInput(completeFormula(text))
    if (!parsed) return undefined
    if (!parsed.formula) return parsed.value
    const ast = parseFormula(parsed.input!.slice(1))
    if (!ast.ok) return { error: "#NAME?" }
    this.ctx.sheet = this.sheetIndex
    this.ctx.row = row
    this.ctx.col = col
    let v = evaluateFormula(ast.ast, this.ctx)
    if (isRef(v)) {
      const a = v.areas[0]
      v = this.ctx.cell(a.sheet, a.r0, a.c0) ?? 0
    } else if (isMatrix(v)) v = v.data[0] ?? 0
    return v === null ? 0 : (v as CellValue)
  }

  /** The rule an entry breaks at a cell (null when it keeps it, or the rule shows no error alert). */
  rejects(text: string, row: number, col: number): ValidationEntry | null {
    const e = validationAt(this.sheets[this.sheetIndex], row, col)
    if (!e || !e.rule.showError || e.rule.type === "any") return null
    return this.valid(e, row, col, this.entryValue(text, row, col)) ? null : e
  }

  /** Whether `value` keeps the rule of entry `e` at a cell. */
  valid(e: ValidationEntry, row: number, col: number, value: CellValue | undefined): boolean {
    const rule = e.rule
    const blank = value === undefined || value === null || value === ""
    if (rule.type === "any") return true
    if (blank && rule.allowBlank) return true
    if (rule.type === "custom") {
      this.ctx.pending = { sheet: this.sheetIndex, row, col, value: blank ? null : (value as Scalar) }
      try {
        const v = this.scalar(rule.formula1, e, row, col)
        return v === true || (typeof v === "number" && v !== 0)
      } finally {
        this.ctx.pending = null
      }
    }
    if (blank) return false
    if (isError(value)) return false
    if (rule.type === "list") {
      const { values, typed } = this.listItems(e, row, col)
      const text = scalarText(value).trim()
      return values.some((item) => {
        if (typeof item === "number" || typeof value === "number") {
          const a = typeof value === "number" ? value : Number(text)
          const b = typeof item === "number" ? item : Number(String(item).trim())
          if (text !== "" && String(item).trim() !== "" && a === b) return true
        }
        const t = typeof item === "string" ? item.trim() : scalarText(item as CellValue)
        return typed ? t === text : compareText(t, text) === 0
      })
    }
    let n: number
    if (rule.type === "textLength") {
      n = scalarText(value).length
    } else {
      if (typeof value !== "number") return false
      if (rule.type === "whole" && !Number.isInteger(value)) return false
      n = value
    }
    const bound = (f: string): number | null => {
      const v = this.scalar(f, e, row, col)
      if (v === null) return 0
      if (typeof v === "number") return v
      if (typeof v === "boolean") return v ? 1 : 0
      return null
    }
    const a = bound(rule.formula1)
    if (a === null) return false
    if (twoValues(rule.operator)) {
      const b = bound(rule.formula2)
      if (b === null) return false
      const within = n >= Math.min(a, b) && n <= Math.max(a, b)
      return rule.operator === "between" ? within : !within
    }
    switch (rule.operator) {
      case "equal":
        return n === a
      case "notEqual":
        return n !== a
      case "greaterThan":
        return n > a
      case "lessThan":
        return n < a
      case "greaterThanOrEqual":
        return n >= a
      case "lessThanOrEqual":
        return n <= a
      default:
        return true
    }
  }
}

function scalarText(v: CellValue | Scalar): string {
  if (v === null || v === undefined) return ""
  if (typeof v === "number") return numberToEditText(v)
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE"
  if (typeof v === "object" && "error" in v) return v.error
  return String(v)
}

/** Circle Invalid Data: the cells with a value (or a blank that must not be) that breaks their rule. */
export function invalidCells(sheets: readonly Sheet[], index: number, today: number, limit = 255): Range[] {
  const sheet = sheets[index]
  const check = new ValidationChecker(sheets, index, today)
  const out: Range[] = []
  const used = sheet.cells.usedBounds()
  for (const e of sheet.validations) {
    if (e.rule.type === "any") continue
    for (const r of e.ranges) {
      // Blank cells only count where blanks aren't allowed; then the whole range is looked at (up to the used area).
      const r1 = used ? Math.min(r.r1, Math.max(used.r1, r.r0)) : r.r0
      const c1 = used ? Math.min(r.c1, Math.max(used.c1, r.c0)) : r.c0
      for (let row = r.r0; row <= r1; row++)
        for (let col = r.c0; col <= c1; col++) {
          if (validationAt(sheet, row, col) !== e) continue
          if (!check.valid(e, row, col, sheet.cells.get(row, col)?.value)) {
            out.push({ r0: row, c0: col, r1: row, c1: col })
            if (out.length >= limit) return out
          }
        }
    }
  }
  return out
}
