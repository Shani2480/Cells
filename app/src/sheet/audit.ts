import { formulaRefs, isError, type ErrorCode, type Range, type Sheet } from "@cells/core"
import type { SheetController } from "./controller"

type At = { row: number; col: number }

/** One tracer arrow on the sheet: from a precedent range (or another sheet) to the formula that uses it. */
export type TraceArrow = {
  /** The range the value comes from; null when it comes from another sheet (drawn from a sheet icon). */
  from: Range | null
  to: At
  /** Red when the value it carries is an error, as Excel draws them. */
  error: boolean
}

const key = (at: At) => `${at.row},${at.col}`
const unkey = (k: string): At => {
  const [row, col] = k.split(",").map(Number)
  return { row, col }
}
const inside = (r: Range, at: At) => at.row >= r.r0 && at.row <= r.r1 && at.col >= r.c0 && at.col <= r.c1

/**
 * Formulas › Formula Auditing: Trace Precedents, Trace Dependents and Remove Arrows. Each sheet remembers which cells
 * were traced, each way; the arrows themselves are worked out from the formulas as they are now, so they follow
 * edits. Each Trace click goes one level further from the active cell, as in Excel.
 */
export class AuditCommands {
  /** By sheet: the cells whose precedents (dependents) are drawn. */
  private traced = new Map<number, { pre: Set<string>; dep: Set<string> }>()

  constructor(private readonly c: SheetController) {}

  private sets() {
    let s = this.traced.get(this.c.sheetIndex)
    if (!s) this.traced.set(this.c.sheetIndex, (s = { pre: new Set(), dep: new Set() }))
    return s
  }

  private input(at: At) {
    return this.c.sheet.cells.input(at.row, at.col)
  }

  private refsOf(at: At) {
    const input = this.input(at)
    return input ? formulaRefs(input, this.c.sheet.name) : { local: [], offSheet: false }
  }

  /** The formulas on the sheet that read `at`. */
  private dependentsOf(at: At): At[] {
    const used = this.c.sheet.cells.usedBounds()
    if (!used) return []
    const out: At[] = []
    this.c.sheet.cells.forEach(used, (row, col, data) => {
      if (data.input?.startsWith("=") && formulaRefs(data.input, this.c.sheet.name).local.some((r) => inside(r, at)))
        out.push({ row, col })
    })
    return out
  }

  /** Formula cells inside a range (the next level of precedents goes through them). */
  private formulasIn(r: Range): At[] {
    const out: At[] = []
    this.c.sheet.cells.forEach(r, (row, col, data) => {
      if (data.input?.startsWith("=") && !data.spill) out.push({ row, col })
    })
    return out
  }

  /** Trace Precedents: the active cell's, then on each click one level further back. False when nothing more to draw. */
  tracePrecedents(): boolean {
    const { pre } = this.sets()
    const active = this.c.selection.active
    let next: At[]
    if (!pre.has(key(active))) next = this.input(active)?.startsWith("=") ? [active] : []
    else {
      next = []
      for (const k of pre)
        for (const r of this.refsOf(unkey(k)).local)
          for (const f of this.formulasIn(r)) if (!pre.has(key(f)) && !next.some((n) => key(n) === key(f))) next.push(f)
    }
    if (!next.length) return false
    for (const at of next) pre.add(key(at))
    this.c.changed()
    return true
  }

  /** Trace Dependents: the formulas that read the active cell, then on each click the ones that read those. */
  traceDependents(): boolean {
    const { dep } = this.sets()
    const active = this.c.selection.active
    let next: At[]
    if (!dep.has(key(active))) next = this.dependentsOf(active).length ? [active] : []
    else {
      next = []
      for (const k of dep)
        for (const d of this.dependentsOf(unkey(k)))
          if (!dep.has(key(d)) && !next.some((n) => key(n) === key(d)) && this.dependentsOf(d).length) next.push(d)
    }
    if (!next.length) return false
    for (const at of next) dep.add(key(at))
    this.c.changed()
    return true
  }

  /** Remove Arrows (all), Remove Precedent Arrows, Remove Dependent Arrows. */
  removeArrows(kind: "all" | "precedents" | "dependents" = "all") {
    const s = this.sets()
    if (kind !== "dependents") s.pre.clear()
    if (kind !== "precedents") s.dep.clear()
    this.c.changed()
  }

  get hasArrows(): boolean {
    const s = this.traced.get(this.c.sheetIndex)
    return Boolean(s && (s.pre.size || s.dep.size))
  }

  /** The arrows to draw on the active sheet now. */
  arrows(): TraceArrow[] {
    const s = this.traced.get(this.c.sheetIndex)
    if (!s) return []
    const out: TraceArrow[] = []
    const seen = new Set<string>()
    const push = (a: TraceArrow) => {
      const k = `${a.from ? `${a.from.r0},${a.from.c0},${a.from.r1},${a.from.c1}` : "sheet"}>${key(a.to)}`
      if (seen.has(k)) return
      seen.add(k)
      out.push(a)
    }
    const errorIn = (r: Range) => {
      let bad = false
      this.c.sheet.cells.forEach(r, (_row, _col, d) => {
        if (isError(d.value)) bad = true
      })
      return bad
    }
    for (const k of s.pre) {
      const to = unkey(k)
      const refs = this.refsOf(to)
      for (const r of refs.local) push({ from: r, to, error: errorIn(r) })
      if (refs.offSheet) push({ from: null, to, error: false })
    }
    for (const k of s.dep) {
      const at = unkey(k)
      const from = { r0: at.row, c0: at.col, r1: at.row, c1: at.col }
      for (const d of this.dependentsOf(at))
        push({ from, to: d, error: isError(this.c.sheet.cells.get(at.row, at.col)?.value) })
    }
    return out
  }

  /**
   * Error Checking › Trace Error: from a cell showing an error, back through the precedents that show errors too, to
   * where the error starts. False when the active cell shows no error.
   */
  traceError(): boolean {
    const active = this.c.selection.active
    if (!isError(this.c.sheet.cells.value(active.row, active.col))) return false
    const { pre } = this.sets()
    const queue: At[] = [active]
    const seen = new Set<string>()
    while (queue.length) {
      const at = queue.shift()!
      if (seen.has(key(at)) || !this.input(at)?.startsWith("=")) continue
      seen.add(key(at))
      pre.add(key(at))
      for (const r of this.refsOf(at).local)
        this.c.sheet.cells.forEach(r, (row, col, d) => {
          if (isError(d.value)) queue.push({ row, col })
        })
    }
    this.c.changed()
    return true
  }

  // ── Error Checking ────────────────────────────────────────────────────────────────────────────────────────────

  /** Cells the user chose Ignore Error for, by sheet ("row,col"). */
  private ignored = new WeakMap<object, Set<string>>()

  private ignoredSet() {
    let set = this.ignored.get(this.c.sheet)
    if (!set) this.ignored.set(this.c.sheet, (set = new Set()))
    return set
  }

  /** What Error Checking flags at a cell, or null: a formula showing an error, or a number stored as text. */
  problemAt(row: number, col: number): CellProblem | null {
    if (this.ignoredSet().has(`${row},${col}`)) return null
    const data = this.c.sheet.cells.get(row, col)
    if (!data || data.spill) return null
    if (isError(data.value) && data.input?.startsWith("=")) return { kind: "error", code: data.value.error }
    if (typeof data.value === "string" && /^\s*[-+]?(\d+([.,]\d+)?|\.\d+)(e[-+]?\d+)?\s*$/i.test(data.value))
      return { kind: "numberAsText" }
    return null
  }

  /** The next (or previous) cell to flag after `from`, row by row across the used range, wrapping; null if none. */
  nextProblem(from: At, back = false): At | null {
    const used = this.c.sheet.cells.usedBounds()
    if (!used) return null
    const cells: At[] = []
    this.c.sheet.cells.forEach(used, (row, col) => {
      if (this.problemAt(row, col)) cells.push({ row, col })
    })
    if (!cells.length) return null
    const after = (a: At) => a.row > from.row || (a.row === from.row && a.col > from.col)
    const before = (a: At) => a.row < from.row || (a.row === from.row && a.col < from.col)
    if (back) return [...cells].reverse().find(before) ?? cells[cells.length - 1]
    return cells.find(after) ?? cells[0]
  }

  /** Convert to Number: a number stored as text becomes the number. */
  convertToNumber(at: At) {
    const sheet = this.c.sheet
    const data = sheet.cells.get(at.row, at.col)
    if (typeof data?.value !== "string") return
    const n = Number(data.value.trim().replace(",", "."))
    if (!Number.isFinite(n)) return
    const sel = this.c.selection
    this.c.transact(
      [
        {
          kind: "setCells",
          sheet: this.c.sheetIndex,
          cells: [{ row: at.row, col: at.col, data: { ...data, value: n, input: undefined } }],
        },
      ],
      { label: "Convert to Number", before: sel, after: sel },
    )
  }

  ignoreError(at: At) {
    this.ignoredSet().add(key(at))
    this.c.changed()
  }

  /** Error Checking's Reset Ignored Errors. */
  resetIgnored() {
    this.ignored.delete(this.c.sheet)
    this.c.changed()
  }

  // ── Watch Window ──────────────────────────────────────────────────────────────────────────────────────────────

  /** Cells kept in view in the Watch Window, by their sheet (so they follow a rename). */
  watches: { sheet: Sheet; row: number; col: number }[] = []
  watchOpen = false

  toggleWatchWindow(open = !this.watchOpen) {
    this.watchOpen = open
    this.c.changed()
  }

  /** Add Watch: the selected cells (up to 100 at a time), each once. */
  addWatches() {
    const sheet = this.c.sheet
    let added = 0
    for (const r of this.c.selection.ranges)
      for (let row = r.r0; row <= r.r1 && added < 100; row++)
        for (let col = r.c0; col <= r.c1 && added < 100; col++) {
          if (this.watches.some((w) => w.sheet === sheet && w.row === row && w.col === col)) continue
          this.watches.push({ sheet, row, col })
          added++
        }
    this.c.changed()
  }

  deleteWatch(index: number) {
    this.watches = this.watches.filter((_, i) => i !== index)
    this.c.changed()
  }
}

export type CellProblem = { kind: "error"; code: ErrorCode } | { kind: "numberAsText" }

/** Error Checking's heading and explanation for each error, in Excel's words. */
export const ERROR_HELP: Record<string, [string, string]> = {
  "#DIV/0!": ["Divide by Zero Error", "The formula or function used is dividing by zero or empty cells."],
  "#N/A": ["Value Not Available Error", "A value is not available to the formula or function."],
  "#NAME?": ["Invalid Name Error", "The formula contains unrecognized text."],
  "#NULL!": ["Null Error", "The ranges in the formula don't intersect."],
  "#NUM!": ["Invalid Number Error", "The formula or function contains invalid numeric values."],
  "#REF!": ["Invalid Cell Reference Error", "The formula refers to cells that have been deleted or replaced."],
  "#VALUE!": ["Error in Value", "A value used in the formula is of the wrong data type."],
  "#SPILL!": ["Spill Range Isn't Blank", "The formula's results can't spill into cells that already hold data."],
  "#CALC!": ["Calculation Error", "The formula can't be calculated, for example an empty array."],
}
