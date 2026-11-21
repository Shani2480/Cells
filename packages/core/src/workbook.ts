import type { WorkbookQuery } from "./power-query"
import type { ExternalBook } from "./links"
import { CalcEngine, cellOfKey, type CalcMode } from "./calc/engine"
import type { Range } from "./address"
import type { CustomView } from "./custom-views"
import type { PivotCache } from "./pivot/model"
import { applyOp, type Change, type Op } from "./ops"
import type { CellValue } from "./values"
import { Sheet } from "./sheet"
import { StyleTable, type CellStyle } from "./styles"
import { NameTable, type SheetList } from "./names"
import type { WorkbookProtection } from "./protection"
import type { Node } from "./formula/ast"
import { copyTheme, OFFICE_THEME, type Theme } from "./theme"

/** A module of the workbook's VBA project, read to show (View › Macros); Cells doesn't run VBA. */
export type VbaModule = { name: string; kind: "module" | "class" | "document"; code: string }

/** One undoable step: a label for the Undo tooltip, the ops either way, and the caller's state either side. */
export type UndoEntry<Meta = unknown> = {
  label: string
  undo: Op[]
  redo: Op[]
  before?: Meta
  after?: Meta
}

export type TransactionOptions<Meta> = {
  label: string
  /** Caller state to restore on undo (the selection before the change). */
  before?: Meta
  /** Caller state to restore on redo (the selection after the change). */
  after?: Meta
  /** false: apply without an undo step (and without clearing redo). */
  history?: boolean
}

/** `ops` are the operations just applied, in order (empty for calculated values arriving from a worker). */
export type WorkbookListener = (changes: readonly Change[], ops: readonly Op[]) => void

export type SpillArea = { sheet: number; row: number; col: number; range: Range }

/** What the status bar and the grid need to know about calculation. */
export type CalcStatus = {
  mode: CalcMode
  /** Manual mode with changes not yet calculated ("Calculate" in the status bar). */
  pending: boolean
  circular: { sheet: number; row: number; col: number } | null
  spills: readonly SpillArea[]
}

/**
 * Who calculates. By default the workbook's own engine, right after each change; the app swaps in a calculator
 * that forwards operations to a Web Worker and applies the values it sends back.
 */
export interface Calculator {
  afterOps(ops: readonly Op[], changes: readonly Change[]): void
  calculate(full: boolean): void
  setMode(mode: CalcMode): void
  status(): CalcStatus
}

/** Calculated values from an engine elsewhere: changed cells (formula results and spills) and the new status. */
export type ValuePatch = {
  cells: { sheet: number; row: number; col: number; value: CellValue | undefined; spill: boolean }[]
  status: CalcStatus
}

/** Collects what the engine wrote since the last call (its change log must be on). */
export function takePatch(wb: Workbook<unknown>): ValuePatch {
  const log = wb.calc.changeLog
  const cells: ValuePatch["cells"] = []
  if (log) {
    for (const key of log) {
      const { sheet, row, col } = cellOfKey(key)
      const store = wb.sheets[sheet]?.cells
      if (!store) continue
      cells.push({ sheet, row, col, value: store.value(row, col), spill: store.isSpill(row, col) })
    }
    log.clear()
  }
  return { cells, status: localStatus(wb.calc) }
}

function localStatus(engine: CalcEngine): CalcStatus {
  return { mode: engine.mode, pending: engine.pending, circular: engine.circular, spills: engine.spillAreas() }
}

class LocalCalculator implements Calculator {
  constructor(private readonly engine: CalcEngine) {}
  afterOps(_ops: readonly Op[], changes: readonly Change[]) {
    this.engine.afterChanges(changes)
  }
  calculate(full: boolean) {
    this.engine.calculateNow(full)
  }
  setMode(mode: CalcMode) {
    this.engine.mode = mode
    if (mode === "automatic" && this.engine.pending) this.engine.calculateNow(false)
  }
  status() {
    return localStatus(this.engine)
  }
}

/** Excel keeps 100 undo levels. */
export const DEFAULT_UNDO_LIMIT = 100

export class Workbook<Meta = unknown> {
  readonly sheets: Sheet[]
  readonly styles = new StyleTable()
  /** Defined names, workbook-wide. */
  readonly names = new NameTable()
  /** Protect Workbook (Structure): null when the workbook isn't protected. */
  readonly structure: { protection: WorkbookProtection | null } = { protection: null }
  /** Other workbooks that formulas link to (Data › Workbook Links), with the values last read from them. */
  readonly links: { list: ExternalBook[] } = { list: [] }
  /** View › Custom Views (kept out of undo, as Excel does). */
  customViews: CustomView[] = []
  /**
   * Review › Share Workbook (Legacy): the file keeps a change history for several people editing it at once. Unshare
   * Workbook turns it off (and drops the history when saving), as Excel does; not undoable, like Excel's.
   */
  legacyShared = false
  /** The VBA project's modules as read from an .xlsm or .xls (the project itself stays in the file as it was). */
  vba: VbaModule[] = []
  /** The PivotTables' caches (their data as last refreshed), shared by the sheets' PivotTables by id. */
  readonly pivotData: { caches: PivotCache[] } = { caches: [] }
  get pivotCaches(): PivotCache[] {
    return this.pivotData.caches
  }
  /** Sets the caches as read from a file (use the "pivotCaches" op to change them). */
  set pivotCaches(caches: PivotCache[]) {
    this.pivotData.caches = caches
  }
  /** Data › Get Data: the workbook's queries (Power Query), in a box so operations can change them. */
  readonly mashup: { queries: WorkbookQuery[] } = { queries: [] }
  get queries(): WorkbookQuery[] {
    return this.mashup.queries
  }
  /** Sets the queries as read from a file (use the "queries" op to change them). */
  set queries(list: WorkbookQuery[]) {
    this.mashup.queries = list
  }
  /** Page Layout › Themes, in a box so operations can change it. */
  readonly look: { theme: Theme } = { theme: copyTheme(OFFICE_THEME) }
  /** The colours and fonts that theme colours and fonts in cells, shapes and charts follow. */
  get theme(): Theme {
    return this.look.theme
  }
  /** Sets the theme as read from a file (cell formats already show it; use the "theme" op to change it). */
  set theme(theme: Theme) {
    this.look.theme = theme
  }
  undoLimit: number
  private readonly undoStack: UndoEntry<Meta>[] = []
  private readonly redoStack: UndoEntry<Meta>[] = []
  private readonly listeners = new Set<WorkbookListener>()
  /** Bumped on every change; views compare it to know when to redraw. */
  version = 0
  /** The workbook's own engine. It calculates unless another calculator is set (see setCalculator). */
  readonly calc: CalcEngine
  private calculator: Calculator

  constructor(sheetNames: string[] = ["Sheet1"], undoLimit = DEFAULT_UNDO_LIMIT) {
    this.sheets = sheetNames.map((name) => new Sheet(name))
    ;(this.sheets as SheetList).names = this.names
    ;(this.sheets as SheetList).structure = this.structure
    ;(this.sheets as SheetList).links = this.links
    ;(this.sheets as SheetList).look = this.look
    ;(this.sheets as SheetList).pivotData = this.pivotData
    ;(this.sheets as SheetList).mashup = this.mashup
    this.undoLimit = undoLimit
    this.calc = new CalcEngine(this)
    this.calculator = new LocalCalculator(this.calc)
  }
  /** A cell's format as the calculation engine sees it (CELL("format"), CELL("protect")…). */
  cellStyle(sheet: number, row: number, col: number): CellStyle {
    const s = this.sheets[sheet]
    return s ? this.styles.get(s.styleAt(row, col)) : {}
  }

  /** The engine's hook: a defined name's formula, parsed. */
  definedName(name: string, _sheet: number): Node | null {
    return this.names.node(name)
  }

  /** Hands calculation to another calculator (the worker host). The own engine then stays idle. */
  setCalculator(calculator: Calculator) {
    this.calculator = calculator
  }

  get calcStatus(): CalcStatus {
    return this.calculator.status()
  }

  /** F9 / Shift+F9 / Ctrl+Alt+F9: recalculates and tells listeners. */
  calculate(full = false) {
    this.calculator.calculate(full)
    this.version++
    this.emit([], [])
  }

  setCalcMode(mode: CalcMode) {
    this.calculator.setMode(mode)
    this.version++
    this.emit([], [])
  }

  /** Writes values calculated elsewhere (no undo step: they follow from the operations already applied). */
  applyPatch(patch: ValuePatch) {
    const touched = new Map<number, Range>()
    for (const { sheet, row, col, value, spill } of patch.cells) {
      const cells = this.sheets[sheet]?.cells
      if (!cells) continue
      if (spill || cells.isSpill(row, col)) cells.setSpill(row, col, value)
      else if (cells.input(row, col)?.startsWith("=")) cells.setComputed(row, col, value)
      const r = touched.get(sheet)
      touched.set(
        sheet,
        r
          ? { r0: Math.min(r.r0, row), c0: Math.min(r.c0, col), r1: Math.max(r.r1, row), c1: Math.max(r.c1, col) }
          : { r0: row, c0: col, r1: row, c1: col },
      )
    }
    this.version++
    const changes: Change[] = [...touched].map(([sheet, range]) => ({ sheet, structural: false, range }))
    this.emit(changes, [])
  }

  subscribe(listener: WorkbookListener) {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  /** Applies `ops` in order as one undo step. Returns the step (or null when nothing was applied). */
  transact(ops: Op[], options: TransactionOptions<Meta>): UndoEntry<Meta> | null {
    if (ops.length === 0) {
      return null
    }
    const { inverse, changes } = this.run(ops)
    const applied = ops
    const entry: UndoEntry<Meta> = {
      label: options.label,
      undo: inverse,
      redo: ops,
      before: options.before,
      after: options.after,
    }
    if (options.history !== false) {
      this.undoStack.push(entry)
      if (this.undoStack.length > this.undoLimit) {
        this.undoStack.splice(0, this.undoStack.length - this.undoLimit)
      }
      this.redoStack.length = 0
    }
    this.emit(changes, applied)
    return entry
  }

  /**
   * Adds follow-up ops to a step just taken (row heights fitted to what an edit wrote): they apply now and undo and
   * redo with the step.
   */
  amend(entry: UndoEntry<Meta>, ops: Op[]) {
    if (!ops.length) return
    const { inverse, changes } = this.run(ops)
    entry.undo = [...inverse, ...entry.undo]
    entry.redo = [...entry.redo, ...ops]
    this.emit(changes, ops)
  }

  get canUndo() {
    return this.undoStack.length > 0
  }

  get canRedo() {
    return this.redoStack.length > 0
  }

  /** The label of the step Undo would take back. */
  get undoLabel(): string | null {
    return this.undoStack[this.undoStack.length - 1]?.label ?? null
  }

  get redoLabel(): string | null {
    return this.redoStack[this.redoStack.length - 1]?.label ?? null
  }

  get undoDepth() {
    return this.undoStack.length
  }

  undo(): UndoEntry<Meta> | null {
    const entry = this.undoStack.pop()
    if (!entry) {
      return null
    }
    const applied = entry.undo
    const { inverse, changes } = this.run(applied)
    entry.redo = inverse
    this.redoStack.push(entry)
    this.emit(changes, applied)
    return entry
  }

  redo(): UndoEntry<Meta> | null {
    const entry = this.redoStack.pop()
    if (!entry) {
      return null
    }
    const applied = entry.redo
    const { inverse, changes } = this.run(applied)
    entry.undo = inverse
    this.undoStack.push(entry)
    this.emit(changes, applied)
    return entry
  }

  clearHistory() {
    this.undoStack.length = 0
    this.redoStack.length = 0
  }

  /** Applies ops; the returned inverse undoes all of them (last op's inverse first). */
  private run(ops: Op[]): { inverse: Op[]; changes: Change[] } {
    const inverses: Op[][] = []
    const changes: Change[] = []
    for (const op of ops) {
      const applied = applyOp(this.sheets, op, this.styles)
      inverses.push(applied.inverse)
      changes.push(applied.change)
    }
    this.calculator.afterOps(ops, changes)
    this.version++
    return { inverse: inverses.reverse().flat(), changes }
  }

  private emit(changes: Change[], ops: readonly Op[]) {
    for (const listener of this.listeners) {
      listener(changes, ops)
    }
  }
}
