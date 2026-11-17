import { MAX_COLS, MAX_ROWS, type Range } from "../address"
import type { Sheet } from "../sheet"
import { errorValue, isError, type CellValue } from "../values"
import type { Node } from "../formula/ast"
import { CALC, setCurrentYear, SPILL } from "../formula/coerce"
import type { EvalContext } from "../formula/context"
import { dateSerial, serialToDate } from "../formula/dates"
import { evaluate, evaluateFormula, refArea, toMatrix } from "../formula/evaluate"
import { FUNCTIONS } from "../formula/functions"
import { allTables, parseStructured, structuredArea, type Table } from "../table/table"
import { parseFormula } from "../formula/parser"
import { isMatrix, isRef, type Area, type Scalar, type Value } from "../formula/types"
import type { Change } from "../ops"

/** A formula cell as the engine knows it. */
type Formula = {
  key: number
  sheet: number
  row: number
  col: number
  text: string
  ast: Node | null
  /** Areas the formula reads, from its references (static) and from INDIRECT/OFFSET/range results (dynamic). */
  precedents: Area[]
  dynamic: Area[]
  volatile: boolean
  state: State
  /** The area it spills into (anchor included), when its result is an array that fits. */
  spill: Area | null
  /** The area it wanted to spill into but could not (#SPILL!). */
  blocked: Area | null
}

const enum State {
  Clean,
  Dirty,
  Visiting,
}

const COL_BITS = 2 ** 14
const ROW_SPAN = 2 ** 20 * COL_BITS

const keyOf = (sheet: number, row: number, col: number) => sheet * ROW_SPAN + row * COL_BITS + col

/** The cell of a change-log key. */
export function cellOfKey(key: number): { sheet: number; row: number; col: number } {
  const sheet = Math.floor(key / ROW_SPAN)
  const rest = key - sheet * ROW_SPAN
  return { sheet, row: Math.floor(rest / COL_BITS), col: rest % COL_BITS }
}

const EMPTY: Formula[] = []

/** Ranges at least this wide go in one list scanned for every change; narrower ones are indexed per column. */
const WIDE = 64

type RangeEntry = { area: Area; dependents: Set<number> }

class NeedsValue {
  constructor(readonly formula: Formula) {}
}

export type CalcMode = "automatic" | "manual"

/** What the engine needs from the workbook. */
export interface CalcHost {
  readonly sheets: readonly Sheet[]
  /** Defined names (Phase 4 adds the Name Manager); the engine asks by upper-case name. */
  definedName?(name: string, sheet: number): Node | null
}

/**
 * The calculation engine: keeps every formula parsed, knows who depends on what, and after each change
 * recalculates exactly the formulas affected, in dependency order, writing their values into the cell store.
 */
export class CalcEngine {
  private readonly formulas = new Map<number, Formula>()
  /** Formula cells per sheet and column (rows), for finding formulas inside an area. */
  private readonly byColumn = new Map<number, Set<number>>()
  /** cell → formulas that read exactly that cell. */
  private readonly cellDependents = new Map<number, Set<number>>()
  /** Distinct ranges read by formulas: per sheet+column for narrow ones, one list for wide ones. */
  private readonly ranges = new Map<string, RangeEntry>()
  private readonly narrow = new Map<number, Set<RangeEntry>>()
  private readonly wide = new Set<RangeEntry>()
  private readonly volatiles = new Set<number>()
  /** Formulas that spill or want to spill (few): checked when cells near them change. */
  private readonly spillers = new Set<Formula>()
  private readonly columnVersions = new Map<number, number>()
  readonly cache = new Map<string, unknown>()
  /**
   * When set, every cell whose value the engine writes (formula results and spilled values) is added here, keyed
   * by sheet, row and column. The worker host turns it into the value patch it sends to the UI.
   */
  changeLog: Set<number> | null = null
  /** Formulas found in a circular reference during the last recalculation. */
  circular: { sheet: number; row: number; col: number } | null = null
  mode: CalcMode = "automatic"
  /** Something changed while calculation was manual (the status bar shows "Calculate"). */
  pending = false
  /** Injected for tests; defaults to the clock and Math.random. */
  clock: () => Date = () => new Date()
  random: () => number = Math.random
  private nowSerial = 0
  private readonly evalStack: Formula[] = []
  /** The workbook's tables, read again on rebuild (a table change is structural). */
  tableList: { sheet: number; table: Table }[] = []

  constructor(private readonly host: CalcHost) {
    this.rebuild()
  }

  // ── Registration ─────────────────────────────────────────────────────────────────────────────────────────────

  /**
   * Re-reads every formula of the workbook (after structural changes or loading). `trusted`: the stored values are
   * already right (a file saved with its calculated values), so the formulas start calculated and only what changes
   * from now on recalculates. Otherwise they all wait for the next calculation.
   */
  rebuild(trusted = false) {
    // Spilled values must go through clearSpill (they live in the cell store); the rest is dropped wholesale.
    for (const f of this.spillers) if (f.spill) this.clearSpill(f, null)
    this.formulas.clear()
    this.byColumn.clear()
    this.cellDependents.clear()
    this.ranges.clear()
    this.narrow.clear()
    this.wide.clear()
    this.volatiles.clear()
    this.spillers.clear()
    this.tableList = allTables(this.host.sheets)
    this.host.sheets.forEach((sheet, s) => {
      sheet.cells.forEachFormula((row, col, input) => {
        const f = this.register(s, row, col, input)
        if (trusted) f.state = State.Clean
      })
    })
  }

  private register(sheet: number, row: number, col: number, input: string): Formula {
    const key = keyOf(sheet, row, col)
    const parsed = parseFormula(input.slice(1))
    const f: Formula = {
      key,
      sheet,
      row,
      col,
      text: input,
      ast: parsed.ok ? parsed.ast : null,
      precedents: [],
      dynamic: [],
      volatile: false,
      state: State.Dirty,
      spill: null,
      blocked: null,
    }
    if (f.ast) {
      const collected = { areas: [] as Area[], volatile: false }
      this.collect(f.ast, sheet, row, col, collected, 0)
      f.precedents = collected.areas
      f.volatile = collected.volatile
    }
    this.formulas.set(key, f)
    const colKey = sheet * MAX_COLS + col
    let rows = this.byColumn.get(colKey)
    if (!rows) this.byColumn.set(colKey, (rows = new Set()))
    rows.add(row)
    for (const area of f.precedents) this.link(f, area)
    if (f.volatile) this.volatiles.add(key)
    return f
  }

  private unregister(f: Formula) {
    for (const area of f.precedents) this.unlink(f, area)
    for (const area of f.dynamic) this.unlink(f, area)
    this.formulas.delete(f.key)
    this.byColumn.get(f.sheet * MAX_COLS + f.col)?.delete(f.row)
    this.volatiles.delete(f.key)
    // Its spilled values go with it.
    if (f.spill) this.clearSpill(f, null)
    this.spillers.delete(f)
  }

  /** Static precedents: every reference in the formula (names expanded), and whether it calls a volatile function. */
  private collect(
    node: Node,
    sheet: number,
    row: number,
    col: number,
    out: { areas: Area[]; volatile: boolean },
    depth: number,
  ) {
    if (depth > 64) return
    switch (node.type) {
      case "ref": {
        const ctx = {
          sheet,
          sheetIndex: (n: string) => this.sheetIndex(n),
          sheetName: (i: number) => this.host.sheets[i]?.name ?? "",
        }
        if (node.sheetTo !== null) {
          const a = this.sheetIndex(node.sheet!)
          const b = this.sheetIndex(node.sheetTo)
          if (a < 0 || b < 0) return
          for (let s = Math.min(a, b); s <= Math.max(a, b); s++) {
            const area = refArea(
              { ...node, sheet: this.host.sheets[s].name, sheetTo: null },
              ctx as unknown as EvalContext,
            )
            if (!isError(area)) out.areas.push(area)
          }
          return
        }
        const area = refArea(node, ctx as unknown as EvalContext)
        if (!isError(area)) out.areas.push(area)
        return
      }
      case "name": {
        const s = node.sheet === null ? sheet : this.sheetIndex(node.sheet)
        const def = s >= 0 ? this.host.definedName?.(node.name.toUpperCase(), s) : null
        if (def) this.collect(def, s, row, col, out, depth + 1)
        return
      }
      case "func":
        if (FUNCTIONS.get(node.name)?.volatile) out.volatile = true
        for (const a of node.args) this.collect(a, sheet, row, col, out, depth)
        return
      case "unary":
      case "percent":
        this.collect(node.arg, sheet, row, col, out, depth)
        return
      case "spill":
        this.collect(node.ref, sheet, row, col, out, depth)
        return
      case "paren":
        this.collect(node.arg, sheet, row, col, out, depth)
        return
      case "binary":
        this.collect(node.left, sheet, row, col, out, depth)
        this.collect(node.right, sheet, row, col, out, depth)
        return
      case "structured": {
        const spec = parseStructured(node.text)
        const area = spec ? structuredArea(spec, this.tableList, sheet, row, col) : null
        if (area && typeof area !== "string") out.areas.push(area)
        return
      }
      case "array":
        return
      default:
        return
    }
  }

  private link(f: Formula, area: Area) {
    if (area.r0 === area.r1 && area.c0 === area.c1) {
      const k = keyOf(area.sheet, area.r0, area.c0)
      let set = this.cellDependents.get(k)
      if (!set) this.cellDependents.set(k, (set = new Set()))
      set.add(f.key)
      return
    }
    const id = `${area.sheet}:${area.r0}:${area.c0}:${area.r1}:${area.c1}`
    let entry = this.ranges.get(id)
    if (!entry) {
      entry = { area, dependents: new Set() }
      this.ranges.set(id, entry)
      if (area.c1 - area.c0 + 1 >= WIDE) this.wide.add(entry)
      else {
        for (let c = area.c0; c <= area.c1; c++) {
          const k = area.sheet * MAX_COLS + c
          let set = this.narrow.get(k)
          if (!set) this.narrow.set(k, (set = new Set()))
          set.add(entry)
        }
      }
    }
    entry.dependents.add(f.key)
  }

  private unlink(f: Formula, area: Area) {
    if (area.r0 === area.r1 && area.c0 === area.c1) {
      const k = keyOf(area.sheet, area.r0, area.c0)
      const set = this.cellDependents.get(k)
      set?.delete(f.key)
      if (set?.size === 0) this.cellDependents.delete(k)
      return
    }
    const id = `${area.sheet}:${area.r0}:${area.c0}:${area.r1}:${area.c1}`
    const entry = this.ranges.get(id)
    if (!entry) return
    entry.dependents.delete(f.key)
    if (entry.dependents.size) return
    this.ranges.delete(id)
    if (this.wide.delete(entry)) return
    for (let c = area.c0; c <= area.c1; c++) this.narrow.get(area.sheet * MAX_COLS + c)?.delete(entry)
  }

  private sheetIndex(name: string): number {
    const upper = name.toUpperCase()
    return this.host.sheets.findIndex((s) => s.name.toUpperCase() === upper)
  }

  // ── Changes ──────────────────────────────────────────────────────────────────────────────────────────────────

  /** Brings formulas up to date after operations. Called by the workbook before it notifies its listeners. */
  afterChanges(all: readonly Change[]) {
    // Formats only (bold, fills, merges): no value or formula moved.
    const changes = all.some((c) => c.format) ? all.filter((c) => !c.format) : all
    if (!changes.length) return
    if (changes.some((c) => c.structural)) {
      // Rows or columns moved: re-read every formula (their text was rewritten by the operation) and recalculate.
      this.clearAllSpills()
      this.rebuild()
      this.cache.clear()
      for (const f of this.formulas.values()) f.state = State.Dirty
      this.recalculate([...this.formulas.values()])
      return
    }
    const dirty: Formula[] = []
    const areas: Area[] = []
    for (const change of changes) {
      if (!change.range) continue
      const area = { sheet: change.sheet, ...change.range }
      areas.push(area)
      this.bumpColumns(area)
      this.syncFormulas(area, dirty)
    }
    this.markDependents(areas, dirty)
    for (const f of this.volatiles) this.markDirty(this.formulas.get(f)!, dirty)
    if (this.mode === "manual") {
      // Manual: only the edited formulas calculate (Excel calculates a formula when you enter it).
      this.pending = this.pending || dirty.length > 0
      this.recalculate(dirty.filter((f) => areas.some((a) => contains(a, f.sheet, f.row, f.col))))
      return
    }
    this.recalculate(dirty)
  }

  /** Registers, re-registers or drops formulas whose cells were written. */
  private syncFormulas(area: Area, dirty: Formula[]) {
    const sheet = this.host.sheets[area.sheet]
    // Formulas that were there.
    for (const f of this.formulasIn(area)) {
      const input = sheet.cells.input(f.row, f.col)
      if (input === f.text) continue
      this.unregister(f)
    }
    // Formulas that are there now.
    sheet.cells.forEach(area, (row, col, data) => {
      if (!data.input?.startsWith("=")) return
      const key = keyOf(area.sheet, row, col)
      if (this.formulas.has(key)) return
      dirty.push(this.register(area.sheet, row, col, data.input))
    })
  }

  private formulasIn(area: Area): Formula[] {
    const out: Formula[] = []
    const cols = area.c1 - area.c0 + 1
    if (cols > 256) {
      for (const f of this.formulas.values())
        if (f.sheet === area.sheet && contains(area, f.sheet, f.row, f.col)) out.push(f)
      return out
    }
    for (let c = area.c0; c <= area.c1; c++) {
      const rows = this.byColumn.get(area.sheet * MAX_COLS + c)
      if (!rows || rows.size === 0) continue
      if (area.r1 - area.r0 + 1 < rows.size) {
        for (let r = area.r0; r <= area.r1; r++) if (rows.has(r)) out.push(this.formulas.get(keyOf(area.sheet, r, c))!)
      } else {
        for (const r of rows) if (r >= area.r0 && r <= area.r1) out.push(this.formulas.get(keyOf(area.sheet, r, c))!)
      }
    }
    return out
  }

  private markDirty(f: Formula, out: Formula[]) {
    if (f.state === State.Dirty) return
    f.state = State.Dirty
    out.push(f)
  }

  /** Every formula that reads any of `areas`, and the formulas that read those, and so on. */
  private markDependents(areas: Area[], out: Formula[]) {
    const mark = (f: Formula) => {
      if (f.state === State.Dirty) return
      f.state = State.Dirty
      out.push(f)
      queue.push(f)
    }
    // Formulas whose own cells changed (entered or re-entered) pass the change on, like any formula marked here.
    const queue: Formula[] = [...out]
    for (const area of areas) for (const f of this.readersOf(area)) mark(f)
    while (queue.length) {
      const f = queue.pop()!
      this.forEachReaderOfCell(f.sheet, f.row, f.col, mark)
      if (f.spill) for (const g of this.readersOf(f.spill)) mark(g)
    }
  }

  /**
   * Calls `visit` for every formula that reads one cell (directly, through a range, or as a spill it may block).
   * A formula can be visited twice; this is the hot path of a recalculation, so it allocates nothing.
   */
  private forEachReaderOfCell(sheet: number, row: number, col: number, visit: (f: Formula) => void) {
    const set = this.cellDependents.get(keyOf(sheet, row, col))
    if (set) {
      for (const k of set) {
        const f = this.formulas.get(k)
        if (f) visit(f)
      }
    }
    if (this.ranges.size) {
      const inside = (a: Area) => a.sheet === sheet && row >= a.r0 && row <= a.r1 && col >= a.c0 && col <= a.c1
      const narrow = this.narrow.get(sheet * MAX_COLS + col)
      if (narrow) {
        for (const entry of narrow) {
          if (!inside(entry.area)) continue
          for (const k of entry.dependents) {
            const f = this.formulas.get(k)
            if (f) visit(f)
          }
        }
      }
      for (const entry of this.wide) {
        if (!inside(entry.area)) continue
        for (const k of entry.dependents) {
          const f = this.formulas.get(k)
          if (f) visit(f)
        }
      }
    }
    for (const f of this.spillers) {
      if (f.blocked && contains(f.blocked, sheet, row, col)) visit(f)
      else if (f.spill && contains(f.spill, sheet, row, col) && (f.row !== row || f.col !== col || f.sheet !== sheet))
        visit(f)
    }
  }

  private readersOf(area: Area): Formula[] {
    if (area.r0 === area.r1 && area.c0 === area.c1 && this.ranges.size === 0 && this.spillers.size === 0) {
      // The common case (a chain of single-cell references): no allocation beyond the result.
      const set = this.cellDependents.get(keyOf(area.sheet, area.r0, area.c0))
      if (!set) return EMPTY
      const out: Formula[] = []
      for (const k of set) {
        const f = this.formulas.get(k)
        if (f) out.push(f)
      }
      return out
    }
    const out: Formula[] = []
    const seen = new Set<number>()
    const add = (k: number) => {
      if (seen.has(k)) return
      seen.add(k)
      const f = this.formulas.get(k)
      if (f) out.push(f)
    }
    const cells = (area.r1 - area.r0 + 1) * (area.c1 - area.c0 + 1)
    if (cells <= 4096) {
      for (let r = area.r0; r <= area.r1; r++) {
        for (let c = area.c0; c <= area.c1; c++) {
          const set = this.cellDependents.get(keyOf(area.sheet, r, c))
          if (set) for (const k of set) add(k)
        }
      }
    } else {
      for (const [cellKey, set] of this.cellDependents) {
        const s = Math.floor(cellKey / ROW_SPAN)
        const rest = cellKey - s * ROW_SPAN
        const r = Math.floor(rest / COL_BITS)
        const c = rest - r * COL_BITS
        if (contains(area, s, r, c)) for (const k of set) add(k)
      }
    }
    const visitEntry = (entry: RangeEntry) => {
      if (overlaps(entry.area, area)) for (const k of entry.dependents) add(k)
    }
    if (area.c1 - area.c0 + 1 > 256) {
      for (const entry of this.ranges.values()) visitEntry(entry)
    } else {
      const visited = new Set<RangeEntry>()
      for (let c = area.c0; c <= area.c1; c++) {
        const set = this.narrow.get(area.sheet * MAX_COLS + c)
        if (!set) continue
        for (const entry of set) {
          if (visited.has(entry)) continue
          visited.add(entry)
          visitEntry(entry)
        }
      }
      for (const entry of this.wide) visitEntry(entry)
    }
    // Spills that the change may block or unblock.
    for (const f of this.spillers) {
      if (f.blocked && overlaps(f.blocked, area)) add(f.key)
      else if (f.spill && overlaps(f.spill, area) && !contains(area, f.sheet, f.row, f.col)) add(f.key)
    }
    return out
  }

  private bumpColumns(area: Area) {
    if (area.c1 - area.c0 > 1024) {
      this.columnVersions.clear()
      this.cache.clear()
      return
    }
    for (let c = area.c0; c <= area.c1; c++) {
      const k = area.sheet * MAX_COLS + c
      this.columnVersions.set(k, (this.columnVersions.get(k) ?? 0) + 1)
    }
  }

  // ── Calculation ──────────────────────────────────────────────────────────────────────────────────────────────

  /** F9: everything volatile and everything dirty; Ctrl+Alt+F9 (full) recalculates every formula. */
  calculateNow(full = false) {
    const dirty: Formula[] = []
    if (full) {
      this.cache.clear()
      // Formulas registered but never calculated (just loaded) are already dirty; they count too.
      for (const f of this.formulas.values()) {
        f.state = State.Dirty
        dirty.push(f)
      }
    } else {
      for (const f of this.formulas.values()) if (f.state === State.Dirty) dirty.push(f)
      for (const k of this.volatiles) this.markDirty(this.formulas.get(k)!, dirty)
      if (this.pending) {
        // In manual mode the dependents of earlier edits were left dirty-but-unmarked: mark all.
        for (const f of this.formulas.values()) this.markDirty(f, dirty)
      }
    }
    this.pending = false
    this.recalculate(dirty)
  }

  private recalculate(dirty: Formula[]) {
    if (dirty.length === 0) return
    const now = this.clock()
    this.nowSerial =
      dateSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()) +
      (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds() + now.getMilliseconds() / 1000) / 86_400
    setCurrentYear(now.getFullYear())
    this.whatIfReaders.clear()
    this.circular = null
    const stack = [...dirty]
    let rounds = 0
    for (;;) {
      if (!stack.length) {
        // Spills that changed size or values re-queue the formulas that read them (bounded, in case two spills
        // keep pushing each other around).
        if (!this.pendingAgain.length || rounds++ > 16) break
        stack.push(...this.pendingAgain)
        this.pendingAgain = []
      }
      const f = stack[stack.length - 1]
      if (!this.formulas.has(f.key) || f.state === State.Clean) {
        stack.pop()
        continue
      }
      if (f.state === State.Dirty) {
        f.state = State.Visiting
        for (const p of this.dirtyPrecedents(f)) {
          if (p.state === State.Dirty) stack.push(p)
          else if (p.state === State.Visiting && p !== f && this.onPath(stack, p))
            this.circular ??= { sheet: f.sheet, row: f.row, col: f.col }
        }
        continue
      }
      // Visiting: all precedents we knew of are done. Evaluate; a value we did not foresee interrupts once.
      const need = this.evaluate(f)
      if (need) {
        if (need.state === State.Dirty) {
          stack.push(need)
          continue
        }
        // It is on the path: a circular reference through a dynamic reference.
        this.circular ??= { sheet: f.sheet, row: f.row, col: f.col }
        this.write(f, 0)
      }
      f.state = State.Clean
      stack.pop()
    }
  }

  private onPath(stack: Formula[], f: Formula) {
    return stack.includes(f)
  }

  /** Dirty formulas inside the areas a formula reads (including the anchors of spills that reach into them). */
  private dirtyPrecedents(f: Formula): Formula[] {
    const out: Formula[] = []
    const visitArea = (area: Area) => {
      if (area.r0 === area.r1 && area.c0 === area.c1) {
        const g = this.formulas.get(keyOf(area.sheet, area.r0, area.c0))
        if (g && g.state !== State.Clean) out.push(g)
      } else {
        for (const g of this.formulasIn(area)) if (g.state !== State.Clean) out.push(g)
      }
      if (!this.spillers.size) return
      for (const g of this.spillers) {
        if (g.state !== State.Clean && g !== f && overlaps(g.spill ?? g.blocked!, area)) out.push(g)
      }
    }
    for (const area of f.precedents) visitArea(area)
    for (const area of f.dynamic) visitArea(area)
    return out
  }

  /** Evaluates one formula and stores its result. Returns a formula that has to be calculated first, if any. */
  private evaluate(f: Formula): Formula | null {
    if (!f.ast) {
      this.write(f, errorValue("#NAME?"))
      return null
    }
    const ctx = this.contextFor(f)
    const newDynamic: Area[] = this.ctx.dynamicAreas
    let result: Value
    this.evalStack.push(f)
    const circularBefore = this.circular
    this.circular = null
    try {
      result = isDataTable(f.ast) ? this.dataTable(f, f.ast, this.ctx) : evaluateFormula(f.ast, ctx)
    } catch (error) {
      if (error instanceof NeedsValue) return error.formula
      throw error
    } finally {
      this.evalStack.pop()
    }
    const hitCircle = this.circular !== null
    this.circular = circularBefore ?? this.circular
    if (newDynamic.length || f.dynamic.length) this.setDynamic(f, [...newDynamic])
    // A formula in a circular reference shows 0 (iterative calculation is off, as in a new Excel workbook).
    this.store(f, hitCircle ? 0 : result)
    return null
  }

  // ── What-If data tables ──────────────────────────────────────────────────────────────────────────────────────

  /** Formulas that depend on a data table's input cells, worked out once per recalculation. */
  private whatIfReaders = new Map<string, Set<Formula>>()

  /**
   * One cell of a What-If data table, {=TABLE(row input, column input)}, as Excel lays them out: the cells holding the
   * same TABLE formula form a block; the row above it and the column left of it hold the values to try and the
   * formulas to work out. The formula is worked out again with the input cell(s) taken as this cell's values; only
   * the formulas that depend on an input are, privately (nothing is written to the sheet).
   */
  private dataTable(f: Formula, ast: Extract<Node, { type: "func" }>, ctx: FormulaContext): Value {
    const inputOf = (node: Node | undefined): Area | null | "bad" => {
      if (!node || node.type === "missing") return null
      if (node.type !== "ref") return "bad"
      const area = refArea(node, ctx)
      return isError(area) || area.r0 !== area.r1 || area.c0 !== area.c1 ? "bad" : area
    }
    const rowIn = inputOf(ast.args[0])
    const colIn = inputOf(ast.args[1])
    if (rowIn === "bad" || colIn === "bad" || ast.args.length !== 2 || (!rowIn && !colIn)) return errorValue("#REF!")
    let top = f.row
    while (top > 0 && this.formulaAt(f.sheet, top - 1, f.col)?.text === f.text) top--
    let left = f.col
    while (left > 0 && this.formulaAt(f.sheet, f.row, left - 1)?.text === f.text) left--
    if (top === 0 || left === 0) return errorValue("#REF!")
    const read = (row: number, col: number) => {
      ctx.dynamicReference({ sheet: f.sheet, r0: row, r1: row, c0: col, c1: col })
      return ctx.cell(f.sheet, row, col)
    }
    const overrides = new Map<number, Scalar>()
    let formula: { row: number; col: number }
    if (rowIn && colIn) {
      formula = { row: top - 1, col: left - 1 }
      overrides.set(keyOf(rowIn.sheet, rowIn.r0, rowIn.c0), read(top - 1, f.col))
      overrides.set(keyOf(colIn.sheet, colIn.r0, colIn.c0), read(f.row, left - 1))
    } else if (colIn) {
      formula = { row: top - 1, col: f.col }
      overrides.set(keyOf(colIn.sheet, colIn.r0, colIn.c0), read(f.row, left - 1))
    } else {
      formula = { row: f.row, col: left - 1 }
      overrides.set(keyOf(rowIn!.sheet, rowIn!.r0, rowIn!.c0), read(top - 1, f.col))
    }
    const plain = read(formula.row, formula.col)
    const g = this.formulaAt(f.sheet, formula.row, formula.col)
    const inputs = [rowIn, colIn].filter((a): a is Area => a !== null)
    const readers = this.readersOfInputs(inputs)
    if (!g || !readers.has(g)) return plain
    const state: WhatIfState = { overrides, readers, memo: new Map(), visiting: new Set(), cache: new Map() }
    return this.whatIfValue(g, state, ctx)
  }

  /** Every formula that depends on the input cells, data tables left out (they hold their own values). */
  private readersOfInputs(inputs: Area[]): Set<Formula> {
    const key = inputs.map((a) => keyOf(a.sheet, a.r0, a.c0)).join(",")
    let out = this.whatIfReaders.get(key)
    if (out) return out
    out = new Set()
    const queue: Formula[] = []
    const add = (g: Formula) => {
      if (out!.has(g) || isDataTable(g.ast)) return
      out!.add(g)
      queue.push(g)
    }
    for (const a of inputs) for (const g of this.readersOf(a)) add(g)
    while (queue.length) {
      const g = queue.pop()!
      this.forEachReaderOfCell(g.sheet, g.row, g.col, add)
    }
    this.whatIfReaders.set(key, out)
    return out
  }

  /** @internal A formula's value with the data table's inputs in place (WhatIfContext reads through it). */
  whatIfValue(g: Formula, state: WhatIfState, base: FormulaContext): Scalar {
    const known = state.memo.get(g)
    if (known !== undefined) return known
    if (state.visiting.has(g) || state.visiting.size > 256) return 0
    if (!g.ast) return errorValue("#NAME?")
    state.visiting.add(g)
    let v: Value
    try {
      v = evaluateFormula(g.ast, new WhatIfContext(base, this, g, state))
    } finally {
      state.visiting.delete(g)
    }
    const scalar: Scalar = isMatrix(v) ? (v.data[0] ?? 0) : isRef(v) ? CALC : ((v as Scalar) ?? 0)
    state.memo.set(g, scalar)
    return scalar
  }

  private setDynamic(f: Formula, areas: Area[]) {
    for (const a of f.dynamic) this.unlink(f, a)
    f.dynamic = areas
    for (const a of areas) this.link(f, a)
  }

  /** Writes the result: a scalar into the cell, an array into the cells it spills over (or #SPILL!). */
  private store(f: Formula, result: Value) {
    if (!isMatrix(result)) {
      if (f.spill) this.clearSpill(f, null)
      f.blocked = null
      this.spillers.delete(f)
      this.write(f, isRef(result) ? CALC : ((result as Scalar) ?? 0))
      return
    }
    this.spillers.add(f)
    const area: Area = {
      sheet: f.sheet,
      r0: f.row,
      c0: f.col,
      r1: f.row + result.rows - 1,
      c1: f.col + result.cols - 1,
    }
    if (area.r1 >= MAX_ROWS || area.c1 >= MAX_COLS || this.blockedBy(f, area)) {
      if (f.spill) this.clearSpill(f, null)
      f.blocked = area
      this.write(f, SPILL)
      return
    }
    f.blocked = null
    const old = f.spill
    if (old) this.clearSpill(f, area)
    f.spill = area
    const cells = this.host.sheets[f.sheet].cells
    let changed = !old || !sameArea(old, area)
    for (let r = 0; r < result.rows; r++) {
      for (let c = 0; c < result.cols; c++) {
        const v = result.data[r * result.cols + c] ?? 0
        if (r === 0 && c === 0) {
          this.write(f, v)
          continue
        }
        const row = f.row + r
        const col = f.col + c
        const before = cells.value(row, col)
        if (!sameValue(before, v as CellValue)) changed = true
        cells.setSpill(row, col, v as CellValue)
        this.changeLog?.add(keyOf(f.sheet, row, col))
      }
    }
    if (changed) {
      this.bumpColumns(area)
      // Formulas already calculated in this pass that read the spilled cells (old or new) must run again.
      for (const a of old ? [area, old] : [area]) {
        for (const g of this.readersOf(a)) {
          if (g === f || g.state !== State.Clean) continue
          g.state = State.Dirty
          this.pendingAgain.push(g)
        }
      }
    }
  }

  private pendingAgain: Formula[] = []

  /** Another value is in the way when a cell of the area holds content of its own or another formula's spill. */
  private blockedBy(f: Formula, area: Area): boolean {
    const cells = this.host.sheets[f.sheet].cells
    let blocked = false
    cells.forEach(
      area,
      (row, col, data) => {
        if (blocked || (row === f.row && col === f.col)) return
        if (data.value === undefined && data.input === undefined) return
        if (data.spill) {
          const owner = this.spillOwner(f.sheet, row, col)
          if (owner && owner !== f) blocked = true
          return
        }
        blocked = true
      },
      true,
    )
    return blocked
  }

  private spillOwner(sheet: number, row: number, col: number): Formula | null {
    for (const g of this.spillers)
      if (g.spill && contains(g.spill, sheet, row, col) && !(g.row === row && g.col === col)) return g
    return null
  }

  /** Removes a formula's spilled values, except those inside `keep`. */
  private clearSpill(f: Formula, keep: Area | null) {
    const area = f.spill
    if (!area) return
    const cells = this.host.sheets[f.sheet]?.cells
    if (cells) {
      for (let r = area.r0; r <= area.r1; r++) {
        for (let c = area.c0; c <= area.c1; c++) {
          if (r === f.row && c === f.col) continue
          if (keep && contains(keep, f.sheet, r, c)) continue
          if (cells.isSpill(r, c)) {
            cells.setSpill(r, c, undefined)
            this.changeLog?.add(keyOf(f.sheet, r, c))
          }
        }
      }
    }
    f.spill = null
    this.bumpColumns(area)
  }

  private clearAllSpills() {
    for (const f of [...this.spillers]) {
      if (f.spill) this.clearSpill(f, null)
      f.blocked = null
      this.spillers.delete(f)
    }
  }

  private write(f: Formula, value: Scalar) {
    const v: CellValue = value === null ? 0 : value
    const cells = this.host.sheets[f.sheet].cells
    if (sameValue(cells.value(f.row, f.col), v)) return
    cells.setComputed(f.row, f.col, v)
    this.changeLog?.add(keyOf(f.sheet, f.row, f.col))
    const k = f.sheet * MAX_COLS + f.col
    this.columnVersions.set(k, (this.columnVersions.get(k) ?? 0) + 1)
  }

  // ── Evaluation context ───────────────────────────────────────────────────────────────────────────────────────

  private readonly ctx = new FormulaContext(this)

  private contextFor(f: Formula): EvalContext {
    this.ctx.reset(f)
    return this.ctx
  }

  /** @internal Used by FormulaContext. */
  checkRead(f: Formula, sheet: number, row: number, col: number) {
    const g = this.formulas.get(keyOf(sheet, row, col))
    if (g && g !== f && g.state !== State.Clean) {
      if (g.state === State.Dirty) throw new NeedsValue(g)
      // Visiting: a circular reference.
      this.circular ??= { sheet: f.sheet, row: f.row, col: f.col }
    }
  }

  /** @internal */
  /**
   * Evaluate Formula: one part of the formula at a cell, worked out as calculating that cell would, with ranges read
   * as arrays. Nothing is stored. Null when the cell has no formula.
   */
  evaluatePart(part: Node, sheet: number, row: number, col: number): Value | null {
    const f = this.formulaAt(sheet, row, col)
    if (!f) return null
    const ctx = this.contextFor(f)
    this.evalStack.push(f)
    try {
      const v = evaluate(part, ctx)
      return isRef(v) ? toMatrix(v, ctx) : v
    } catch (error) {
      if (error instanceof NeedsValue) return errorValue("#CALC!")
      throw error
    } finally {
      this.evalStack.pop()
    }
  }

  formulaAt(sheet: number, row: number, col: number) {
    return this.formulas.get(keyOf(sheet, row, col))
  }

  /** @internal */
  spillOwnerFor(sheet: number, row: number, col: number) {
    return this.spillOwner(sheet, row, col)
  }

  /** @internal */
  formulasInArea(area: Area) {
    return this.formulasIn(area)
  }

  /** @internal */
  get hostSheets() {
    return this.host.sheets
  }

  /** @internal */
  get now() {
    return this.nowSerial
  }

  /** @internal */
  versionOf(sheet: number, col: number) {
    return this.columnVersions.get(sheet * MAX_COLS + col) ?? 0
  }

  /** @internal */
  sheetIndexOf(name: string) {
    return this.sheetIndex(name)
  }

  /** @internal */
  definedName(name: string, sheet: number) {
    return this.host.definedName?.(name.toUpperCase(), sheet) ?? null
  }

  // ── Queries for the shell ────────────────────────────────────────────────────────────────────────────────────

  /** The spill area of the formula at a cell, or of the formula whose spill covers it. */
  spillAt(sheet: number, row: number, col: number): Range | null {
    const own = this.formulas.get(keyOf(sheet, row, col))
    if (own?.spill) return areaRange(own.spill)
    const owner = this.spillOwner(sheet, row, col)
    return owner?.spill ? areaRange(owner.spill) : null
  }

  /** The anchor (formula cell) of a spill covering the cell, for the greyed formula bar text in spilled cells. */
  spillAnchor(sheet: number, row: number, col: number): { row: number; col: number; text: string } | null {
    const owner = this.spillOwner(sheet, row, col)
    return owner ? { row: owner.row, col: owner.col, text: owner.text } : null
  }

  /** Every spill area, anchor first (for drawing spill borders without the engine). */
  spillAreas(): { sheet: number; row: number; col: number; range: Range }[] {
    const out: { sheet: number; row: number; col: number; range: Range }[] = []
    for (const f of this.spillers)
      if (f.spill) out.push({ sheet: f.sheet, row: f.row, col: f.col, range: areaRange(f.spill) })
    return out
  }

  get formulaCount() {
    return this.formulas.size
  }

  /** The current date serial (for tests). */
  today(): number {
    const d = this.clock()
    return dateSerial(d.getFullYear(), d.getMonth() + 1, d.getDate())
  }
}

function contains(area: Area, sheet: number, row: number, col: number) {
  return area.sheet === sheet && row >= area.r0 && row <= area.r1 && col >= area.c0 && col <= area.c1
}

function overlaps(a: Area, b: Area) {
  return a.sheet === b.sheet && a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
}

function sameArea(a: Area, b: Area) {
  return a.sheet === b.sheet && a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1
}

function areaRange(a: Area): Range {
  return { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }
}

function sameValue(a: CellValue | undefined, b: CellValue | undefined) {
  if (a === b) return true
  if (isError(a) && isError(b)) return a.error === b.error
  return false
}

export { serialToDate }

/** A data table cell: {=TABLE(row input, column input)}. */
function isDataTable(ast: Node | null): ast is Extract<Node, { type: "func" }> {
  return ast?.type === "func" && ast.name === "TABLE"
}

type WhatIfState = {
  /** The input cells, by key, with the values tried. */
  overrides: Map<number, Scalar>
  /** Formulas that depend on the inputs: worked out again. The rest read as they are. */
  readers: Set<Formula>
  memo: Map<Formula, Scalar>
  visiting: Set<Formula>
  /** Lookup indexes for this evaluation only (the inputs change what they'd index). */
  cache: Map<string, unknown>
}

/** Reads cells for one formula inside a data table's evaluation: inputs and the formulas that use them substituted. */
class WhatIfContext implements EvalContext {
  readonly sheet: number
  readonly row: number
  readonly col: number
  readonly cache: Map<string, unknown>

  constructor(
    private readonly base: FormulaContext,
    private readonly engine: CalcEngine,
    at: { sheet: number; row: number; col: number },
    private readonly state: WhatIfState,
  ) {
    this.sheet = at.sheet
    this.row = at.row
    this.col = at.col
    this.cache = state.cache
  }

  private substitute(sheet: number, row: number, col: number): { value: Scalar } | null {
    const k = keyOf(sheet, row, col)
    if (this.state.overrides.has(k)) return { value: this.state.overrides.get(k)! }
    const g = this.engine.formulaAt(sheet, row, col)
    if (g && this.state.readers.has(g)) return { value: this.engine.whatIfValue(g, this.state, this.base) }
    return null
  }

  cell(sheet: number, row: number, col: number): Scalar {
    return this.substitute(sheet, row, col)?.value ?? this.base.cell(sheet, row, col)
  }

  values(area: Area, visit: (value: Scalar, row: number, col: number) => void) {
    const seen = new Set<number>()
    this.base.values(area, (value, row, col) => {
      seen.add(keyOf(area.sheet, row, col))
      const sub = this.substitute(area.sheet, row, col)
      if (!sub) visit(value, row, col)
      else if (sub.value !== null) visit(sub.value, row, col)
    })
    // An input that was blank still counts once it has a value.
    for (const [k, value] of this.state.overrides) {
      if (seen.has(k) || value === null) continue
      const at = cellOfKey(k)
      if (at.sheet === area.sheet && contains(area, at.sheet, at.row, at.col)) visit(value, at.row, at.col)
    }
  }

  usedArea(area: Area) {
    return this.base.usedArea(area)
  }
  sheetIndex(name: string) {
    return this.base.sheetIndex(name)
  }
  sheetName(i: number) {
    return this.base.sheetName(i)
  }
  sheetCount() {
    return this.base.sheetCount()
  }
  defined(name: string, sheet: number) {
    return this.base.defined(name, sheet)
  }
  spillArea(sheet: number, row: number, col: number) {
    return this.base.spillArea(sheet, row, col)
  }
  isFormula(sheet: number, row: number, col: number) {
    return this.base.isFormula(sheet, row, col)
  }
  formulaText(sheet: number, row: number, col: number) {
    return this.base.formulaText(sheet, row, col)
  }
  now() {
    return this.base.now()
  }
  random() {
    return this.base.random()
  }
  isHiddenRow(sheet: number, row: number) {
    return this.base.isHiddenRow(sheet, row)
  }
  tables() {
    return this.base.tables()
  }
}

const toScalar = (v: CellValue | undefined): Scalar => (v === undefined ? null : v)

/** The evaluation context, reused for every formula (one object instead of a set of closures per evaluation). */
class FormulaContext implements EvalContext {
  sheet = 0
  row = 0
  col = 0
  f!: Formula
  dynamicAreas: Area[] = []
  readonly cache: Map<string, unknown>

  constructor(private readonly engine: CalcEngine) {
    this.cache = engine.cache
  }

  reset(f: Formula) {
    this.f = f
    this.sheet = f.sheet
    this.row = f.row
    this.col = f.col
    this.dynamicAreas = []
  }

  dynamicReference(area: Area) {
    this.dynamicAreas.push(area)
  }

  tables() {
    return this.engine.tableList
  }

  cell(sheet: number, row: number, col: number): Scalar {
    const f = this.f
    if (sheet === f.sheet && row === f.row && col === f.col) {
      this.engine.circular ??= { sheet: f.sheet, row: f.row, col: f.col }
      return 0
    }
    this.engine.checkRead(f, sheet, row, col)
    const cells = this.engine.hostSheets[sheet]?.cells
    if (!cells) return null
    if (cells.isSpill(row, col)) {
      const owner = this.engine.spillOwnerFor(sheet, row, col)
      if (owner && owner.state === State.Dirty) throw new NeedsValue(owner)
    }
    return toScalar(cells.value(row, col))
  }

  values(area: Area, visit: (value: Scalar, row: number, col: number) => void) {
    const cells = this.engine.hostSheets[area.sheet]?.cells
    if (!cells) return
    const f = this.f
    for (const g of this.engine.formulasInArea(area)) {
      if (g === f) {
        this.engine.circular ??= { sheet: f.sheet, row: f.row, col: f.col }
        continue
      }
      this.engine.checkRead(f, g.sheet, g.row, g.col)
    }
    cells.forEach(
      area,
      (row, col, data) => {
        if (area.sheet === f.sheet && row === f.row && col === f.col) return
        visit(toScalar(data.value), row, col)
      },
      true,
    )
  }

  usedArea(area: Area): Area | null {
    const bounds = this.engine.hostSheets[area.sheet]?.cells.usedBounds()
    if (!bounds) return null
    const r0 = Math.max(area.r0, bounds.r0)
    const r1 = Math.min(area.r1, bounds.r1)
    const c0 = Math.max(area.c0, bounds.c0)
    const c1 = Math.min(area.c1, bounds.c1)
    return r0 <= r1 && c0 <= c1 ? { sheet: area.sheet, r0, r1, c0, c1 } : null
  }

  sheetIndex(name: string) {
    return this.engine.sheetIndexOf(name)
  }

  sheetName(i: number) {
    return this.engine.hostSheets[i]?.name ?? ""
  }

  sheetCount() {
    return this.engine.hostSheets.length
  }

  defined(name: string, sheet: number) {
    return this.engine.definedName(name, sheet)
  }

  spillArea(sheet: number, row: number, col: number): Area | null {
    const g = this.engine.formulaAt(sheet, row, col)
    if (!g) return null
    if (g !== this.f && g.state === State.Dirty) throw new NeedsValue(g)
    return g.spill ?? (g.blocked ? null : { sheet, r0: row, r1: row, c0: col, c1: col })
  }

  isFormula(sheet: number, row: number, col: number) {
    return this.engine.formulaAt(sheet, row, col) !== undefined
  }

  formulaText(sheet: number, row: number, col: number) {
    return this.engine.formulaAt(sheet, row, col)?.text ?? null
  }

  now() {
    return this.engine.now
  }

  random() {
    return this.engine.random()
  }

  columnVersion(sheet: number, col: number) {
    return this.engine.versionOf(sheet, col)
  }

  isHiddenRow(sheet: number, row: number) {
    return this.engine.hostSheets[sheet]?.rows.isHidden(row) ?? false
  }
}
