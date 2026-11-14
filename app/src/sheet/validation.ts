import {
  DEFAULT_VALIDATION,
  DEFAULT_VALIDATION_ERROR,
  displayText,
  invalidCells,
  numberToEditText,
  parseInput,
  sameValidation,
  selectCell,
  subtract,
  ValidationChecker,
  validationAt,
  validationsIn,
  withValidation,
  writeInputOp,
  type ErrorStyle,
  type Range,
  type Sheet,
  type Validation,
  type ValidationEntry,
  type ValidationOperator,
  type ValidationType,
} from "@cells/core"
import { todaySerial, type SheetController } from "./controller"

/** What the Data Validation dialog opens with. */
export type ValidationDialogInit = { rule: Validation; ranges: Range[] }

/** The error alert after an entry breaks a rule (Excel's Stop, Warning and Information messages). */
export type ValidationAlert = { style: ErrorStyle; title: string; message: string }

/** How the alert was answered: Stop gives Retry or Cancel, Warning Yes, No or Cancel, Information OK or Cancel. */
export type AlertAnswer = "retry" | "yes" | "no" | "ok" | "cancel"

export const MIXED_RULES =
  "The selection contains more than one type of validation. Erase current settings and continue?"
export const PARTLY_RULED =
  "The selection contains some cells without Data Validation settings. Do you want to extend Data Validation to these cells?"
/** Excel stops at 255 circles. */
export const CIRCLE_LIMIT = 255

/**
 * Data › Data Validation: the dialog's rule for the selection, the check on entry with its error alert, the input
 * message, the in-cell list, and Circle Invalid Data.
 */
export class ValidationCommands {
  constructor(private readonly c: SheetController) {}

  /** Circle Invalid Data: the circled cells of one sheet, until Clear Validation Circles. */
  circles: { sheet: Sheet; cells: Range[] } | null = null

  private checker(): ValidationChecker {
    const checker = new ValidationChecker(this.c.wb.sheets, this.c.sheetIndex, todaySerial())
    checker.formatOf = (style) => this.c.wb.styles.get(style).numberFormat
    return checker
  }

  /** The rule of a cell of the sheet shown. */
  at(row: number, col: number): ValidationEntry | null {
    return validationAt(this.c.sheet, row, col)
  }

  // ── The dialog ──

  /**
   * Data › Data Validation…: asks first when the selection mixes rules (Excel erases them) or has cells without the
   * rule the others have (Excel extends it), then opens the dialog.
   */
  open() {
    if (this.c.edit) this.c.commitEdit()
    const ranges = this.c.selection.ranges.map((r) => ({ ...r }))
    const touched = validationsIn(this.c.sheet, ranges)
    const show = (rule: Validation) => this.c.shell.dataValidation?.({ rule: { ...rule }, ranges })
    if (!touched.length) return show(DEFAULT_VALIDATION)
    if (touched.some((e) => !sameValidation(e.rule, touched[0].rule)))
      return this.c.shell.confirm(MIXED_RULES, () => show(DEFAULT_VALIDATION))
    const rule = touched[0].rule
    if (everyCellRuled(this.c.sheet, ranges)) return show(rule)
    this.c.shell.confirm(PARTLY_RULED, () => show(rule))
  }

  /**
   * The dialog's OK: the rule for `ranges` (Any value with no messages clears it). `sameToo` also gives it to every
   * other cell that had the rule the dialog opened with ("Apply these changes to all other cells with the same
   * settings").
   */
  apply(ranges: Range[], rule: Validation | null, sameToo: Validation | null = null): boolean {
    const sheet = this.c.sheet
    const targets = ranges.map((r) => ({ ...r }))
    if (sameToo) for (const e of sheet.validations) if (sameValidation(e.rule, sameToo)) targets.push(...e.ranges)
    const empty = rule && rule.type === "any" && !(rule.showInput && (rule.prompt || rule.promptTitle))
    const entries = withValidation(sheet.validations, targets, empty ? null : rule)
    const before = this.c.selection
    const ok = this.c.transact([{ kind: "validations", sheet: this.c.sheetIndex, entries }], {
      label: "Data Validation",
      before,
      after: before,
    })
    return Boolean(ok)
  }

  // ── Checking an entry ──

  /** The rule a typed entry breaks at its cell, or null. */
  rejects(text: string, row: number, col: number): ValidationEntry | null {
    return this.checker().rejects(text, row, col)
  }

  /** The error alert for a rule: its own title and message, or Excel's. */
  alertFor(e: ValidationEntry): ValidationAlert {
    return {
      style: e.rule.errorStyle,
      title: e.rule.errorTitle || "Cells",
      message: e.rule.error || DEFAULT_VALIDATION_ERROR,
    }
  }

  // ── Input message ──

  /** The input message of the active cell (null when it has none, or two or more cells are picked in a drag). */
  inputMessage(): { title: string; prompt: string } | null {
    const { row, col } = this.c.selection.active
    const e = this.at(row, col)
    if (!e || !e.rule.showInput || !(e.rule.prompt || e.rule.promptTitle)) return null
    return { title: e.rule.promptTitle, prompt: e.rule.prompt }
  }

  // ── In-cell list ──

  /** The active cell's list, when its rule is a list with the in-cell dropdown. */
  dropdownAt(row: number, col: number): string[] | null {
    const e = this.at(row, col)
    if (!e || e.rule.type !== "list" || !e.rule.dropDown) return null
    return this.checker().listItems(e, row, col).items
  }

  /** Picks an item of the in-cell list: it goes into the cell as if typed (always valid, so no check). */
  pick(row: number, col: number, item: string) {
    if (this.c.edit) this.c.cancelEdit()
    const sel = selectCell({ row, col })
    // An item such as "=x" or "1/2" stays text, as Excel keeps list items as they are.
    const text = /^[=+-]/.test(item) ? "'" + item : item
    const op = writeInputOp(this.c.sheetIndex, this.c.sheet, this.c.wb.styles, row, col, text)
    this.c.transact([op], { label: `Typing '${item}'`, before: this.c.selection, after: sel })
    this.c.setSelection(sel)
  }

  // ── Circle Invalid Data ──

  circleInvalid() {
    if (this.c.edit) this.c.commitEdit()
    const cells = invalidCells(this.c.wb.sheets, this.c.sheetIndex, todaySerial(), CIRCLE_LIMIT + 1)
    if (cells.length > CIRCLE_LIMIT)
      this.c.shell.alert(
        `This worksheet contains more than ${CIRCLE_LIMIT} invalid cells. Only ${CIRCLE_LIMIT} invalid cells will be circled.`,
      )
    this.circles = { sheet: this.c.sheet, cells: cells.slice(0, CIRCLE_LIMIT) }
    this.checkedAt = this.c.wb.version
    this.c.changed()
  }

  clearCircles() {
    if (!this.circles) return
    this.circles = null
    this.c.changed()
  }

  /** Circled cells of the sheet shown. */
  circled(): Range[] {
    return this.circles?.sheet === this.c.sheet ? this.circles.cells : []
  }

  private checkedAt = -1

  /** After a change: a circled cell whose value is now valid loses its circle (Excel adds none by itself). */
  refreshCircles() {
    const circles = this.circles
    if (!circles?.cells.length || this.checkedAt === this.c.wb.version) return
    this.checkedAt = this.c.wb.version
    const index = this.c.wb.sheets.indexOf(circles.sheet)
    if (index < 0) {
      this.circles = null
      return
    }
    const checker = new ValidationChecker(this.c.wb.sheets, index, todaySerial())
    const kept = circles.cells.filter((r) => {
      const e = validationAt(circles.sheet, r.r0, r.c0)
      return e !== null && !checker.valid(e, r.r0, r.c0, circles.sheet.cells.get(r.r0, r.c0)?.value)
    })
    if (kept.length !== circles.cells.length) this.circles = { sheet: circles.sheet, cells: kept }
  }
}

/** Whether every cell of `ranges` has some rule. */
function everyCellRuled(sheet: Sheet, ranges: readonly Range[]): boolean {
  let left = ranges.map((r) => ({ ...r }))
  for (const e of sheet.validations) for (const r of e.ranges) left = left.flatMap((p) => subtract(p, r))
  return left.length === 0
}

// ── The dialog's fields ──

const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/

/** A rule formula as the dialog shows it: "=A1" for formulas, 1/1/2026 for dates, the bare items of a typed list. */
export function fieldText(formula: string, type: ValidationType): string {
  if (!formula) return ""
  if (type === "list") {
    if (formula.startsWith('"') && formula.endsWith('"') && formula.length > 1)
      return formula.slice(1, -1).replace(/""/g, '"')
    return "=" + formula
  }
  if (type !== "custom" && NUMBER.test(formula)) {
    const n = Number(formula)
    if (type === "date") return displayText(n, "m/d/yyyy")
    if (type === "time") return displayText(n, "h:mm:ss AM/PM")
    return numberToEditText(n)
  }
  return "=" + formula
}

/** What a field gives the rule (null: not a value the rule can take). */
export function fieldFormula(text: string, type: ValidationType): string | null {
  const t = text.trim()
  if (!t) return null
  if (t.startsWith("=")) return t.length > 1 ? t.slice(1) : null
  if (type === "list") return '"' + t.replace(/"/g, '""') + '"'
  if (type === "custom") return t
  const parsed = parseInput(t)
  if (!parsed || typeof parsed.value !== "number") return null
  return numberToEditText(parsed.value)
}

/** The labels of the value fields, by type and operator (Excel's Minimum / Start date / Value …). */
export function fieldLabels(type: ValidationType, op: ValidationOperator): [string, string | null] {
  if (type === "list") return ["Source", null]
  if (type === "custom") return ["Formula", null]
  const [low, high, one] =
    type === "date"
      ? ["Start date", "End date", "Date"]
      : type === "time"
        ? ["Start time", "End time", "Time"]
        : type === "textLength"
          ? ["Minimum", "Maximum", "Length"]
          : ["Minimum", "Maximum", "Value"]
  if (op === "between" || op === "notBetween") return [low, high]
  if (op === "greaterThan" || op === "greaterThanOrEqual") return [low, null]
  if (op === "lessThan" || op === "lessThanOrEqual") return [high, null]
  return [one, null]
}

/** Excel's message for a field that holds no value the rule can use. */
export const LIST_SOURCE_ERROR = "The list source must be a delimited list, or a reference to single row or column."
export const fieldError = (label: string) =>
  `The ${label} you entered is not valid. Enter a number, a date or a formula.`
export const blankError = (label: string) => `You must enter a value in the ${label} box.`
export const ORDER_ERROR = "The Maximum must be greater than or equal to the Minimum."
