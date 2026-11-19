import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { Axis } from "./axis"
import type { CellData } from "./cell-store"
import type { AnnotationWrite } from "./annotations"
import { copyEntries, shiftConditional, type CFEntry } from "./conditional"
import { copyValidations, shiftValidations, type ValidationEntry } from "./validation"
import { copyScenarios, shiftScenarios, type ScenarioSet } from "./scenarios"
import { copyAdvancedFilter, shiftAdvancedFilter, type AdvancedFilter } from "./advanced-filter"
import { copyConsolidate, type ConsolidateSettings } from "./consolidate"
import { copyQueryTables, shiftQueryTables, type QueryTable } from "./query-tables"
import { copyHyperlinks, shiftHyperlinks, type Hyperlink } from "./hyperlinks"
import { copyEditRanges, shiftEditRanges, type EditRange } from "./edit-ranges"
import { copySparklineGroups, rewriteSparklineRefs, shiftSparklines, type SparklineGroup } from "./sparklines"
import { copyPivotTables, shiftPivots } from "./pivot/sheet"
import type { PivotTable } from "./pivot/model"
import { copyExternalBooks, type ExternalBook } from "./links"
import { copyTables, shiftTables, syncTableHeaders, type Table } from "./table/table"
import { copyCharts, rewriteChartRefs, type SheetChart } from "./chart"
import { copyObjects, shiftObjects, type SheetObject } from "./shapes"
import { copyPageSetup, shiftPageSetup, type PageSetup } from "./page-setup"
import { copyAutoFilter, shiftAutoFilter, type AutoFilter } from "./filter"
import { rewriteNames, type DefinedName, type SheetList } from "./names"
import type { SheetProtection, WorkbookProtection } from "./protection"
import { formatRanges, recordBounds, restoreStyles, shiftMerges, type StyleRecord } from "./formatting"
import { moveMapper, rewriteFormula, shiftMapper, type RefMapper } from "./formula/adjust"
import type { OutlineSettings, Sheet } from "./sheet"
import { applySheetOp, type SheetOp } from "./sheet-ops"
import { StyleTable, type StylePatch } from "./styles"
import { copyTheme, type Theme } from "./theme"

/** One cell's new content; `data: undefined` clears it. */
export type CellWrite = { row: number; col: number; data: CellData | undefined }

export type AxisName = "rows" | "cols"

/** One row's or column's size and visibility; `size: null` means the default size. */
/** One row's or column's size and visibility; with `level`, its outline level and collapsed mark too. */
export type AxisEntry = {
  index: number
  size: number | null
  hidden: boolean
  auto?: true
  level?: number
  collapsed?: true
}

/**
 * Every change to a workbook is one of these: plain, serialisable data. `applyOp` performs one and returns the
 * operations that undo it, so undo/redo, autosave replay and the worker protocol all share one representation.
 */
export type Op =
  | { kind: "setCells"; sheet: number; cells: CellWrite[] }
  /** `plain`: no formats copied from the line before (the undo of a deletion puts back exactly what was there). */
  | { kind: "insertRows"; sheet: number; at: number; count: number; plain?: true }
  | { kind: "deleteRows"; sheet: number; at: number; count: number }
  | { kind: "insertCols"; sheet: number; at: number; count: number; plain?: true }
  | { kind: "deleteCols"; sheet: number; at: number; count: number }
  /** Insert blank cells in `range`, pushing the cells there down or right (within the range's columns or rows). */
  | { kind: "insertCells"; sheet: number; range: Range; shift: "down" | "right" }
  /** Delete the cells in `range`, pulling the cells below up or the cells to the right left. */
  | { kind: "deleteCells"; sheet: number; range: Range; shift: "up" | "left" }
  | { kind: "axis"; sheet: number; axis: AxisName; entries: AxisEntry[] }
  /**
   * Cut and paste: the cells in `source` move by (dr, dc), overwriting what was there. References anywhere in the
   * workbook that point into the source follow the cells.
   */
  | { kind: "moveCells"; sheet: number; source: Range; dr: number; dc: number }
  /** Formats every cell of `ranges` (whole rows and columns as Excel stores them). */
  | { kind: "format"; sheet: number; ranges: Range[]; patch: StylePatch }
  /** Puts exact style ids back (the undo of a format, Format Painter). Ids belong to the workbook's style table. */
  | { kind: "styles"; sheet: number; record: StyleRecord }
  | { kind: "merge"; sheet: number; ranges: Range[] }
  | { kind: "unmerge"; sheet: number; ranges: Range[] }
  /** The whole set of row formats, column formats and merges (flat [index, id, …]): undo of structural edits. */
  | {
      kind: "layout"
      sheet: number
      rows: number[]
      cols: number[]
      merges: Range[]
      conditional?: CFEntry[]
      validations?: ValidationEntry[]
      tables?: Table[]
      charts?: SheetChart[]
      objects?: SheetObject[]
      pageSetup?: PageSetup
      annotations?: AnnotationWrite[]
      autoFilter?: AutoFilter | null
      scenarios?: ScenarioSet
      advancedFilter?: AdvancedFilter | null
      queryTables?: QueryTable[]
      hyperlinks?: Hyperlink[]
      editRanges?: EditRange[]
      sparklines?: SparklineGroup[]
      pivots?: PivotTable[]
    }
  /** The sheet's whole list of conditional formatting rules (New Rule, Clear Rules, Manage Rules). */
  | { kind: "conditional"; sheet: number; entries: CFEntry[] }
  /** The sheet's whole list of Data Validation rules. */
  | { kind: "validations"; sheet: number; entries: ValidationEntry[] }
  /** The sheet's whole list of tables (create, resize, rename, style, convert to range). */
  | { kind: "tables"; sheet: number; tables: Table[] }
  /** The sheet's whole list of charts (insert, move, resize, change, delete). */
  | { kind: "charts"; sheet: number; charts: SheetChart[] }
  /** The sheet's whole drawing: charts, shapes, pictures and groups in order (insert, arrange, group, delete). */
  | { kind: "objects"; sheet: number; objects: SheetObject[] }
  /** The sheet's page setup (Page Layout, Page Setup, print area and titles, page breaks). */
  | { kind: "pageSetup"; sheet: number; setup: PageSetup }
  /** Notes and comments: set or remove one cell's annotation each. */
  | { kind: "annotations"; sheet: number; writes: AnnotationWrite[] }
  /** Insert › Link, Remove Hyperlinks: the sheet's whole list of links. */
  | { kind: "hyperlinks"; sheet: number; links: Hyperlink[] }
  /** Review › Allow Edit Ranges: the sheet's whole list. */
  | { kind: "editRanges"; sheet: number; ranges: EditRange[] }
  /** Insert › Sparklines and the Sparkline tab: the sheet's whole list of sparkline groups. */
  | { kind: "sparklines"; sheet: number; groups: SparklineGroup[] }
  /** The sheet's whole list of PivotTables (their cells are written by setCells ops alongside). */
  | { kind: "pivots"; sheet: number; tables: PivotTable[] }
  /** Defined names: the whole list after the change (Define Name, the Name Box). */
  | { kind: "names"; names: DefinedName[] }
  /** The linked workbooks and the values read from them (update, change source, break). */
  | { kind: "externalLinks"; links: ExternalBook[] }
  /**
   * Data › Filter: the sheet's AutoFilter (null: none), with the rows it hides or shows (Excel hides filtered rows,
   * so the row entries come from the filter, worked out where today's date is known).
   */
  | { kind: "autoFilter"; sheet: number; filter: AutoFilter | null; rows?: AxisEntry[] }
  /** Data › Advanced in place: the filter and the rows it shows and hides. */
  | { kind: "advancedFilter"; sheet: number; filter: AdvancedFilter | null; rows?: AxisEntry[] }
  /** The sheet's ranges linked to text files (import, refresh, rename, delete). */
  | { kind: "queryTables"; sheet: number; tables: QueryTable[] }
  /** Data › Consolidate: what the dialog last used on the sheet. */
  | { kind: "consolidateSettings"; sheet: number; settings: ConsolidateSettings | null }
  /** Protect Sheet (with the options) or Unprotect Sheet (null). */
  | { kind: "protect"; sheet: number; protection: SheetProtection | null }
  /** Protect Workbook (Structure) or Unprotect Workbook (null). */
  | { kind: "protectWorkbook"; protection: WorkbookProtection | null }
  /** Page Layout › Themes: the workbook's theme; formats with theme colours and fonts follow it. */
  | { kind: "theme"; theme: Theme }
  /** Outline levels alone ([index, level] pairs), leaving sizes and visibility: Subtotal's grouping. */
  | { kind: "outlineLevels"; sheet: number; axis: AxisName; levels: [number, number][] }
  /** Data › Outline › Settings (and Ctrl+8's outline symbols). */
  | { kind: "outlineSettings"; sheet: number; settings: OutlineSettings }
  /** Data › What-If Analysis › Scenario Manager: the sheet's scenarios. */
  | { kind: "scenarios"; sheet: number; scenarios: ScenarioSet }
  /** Sheets: add, delete, rename, move, tab colour and hiding (see sheet-ops). */
  | SheetOp

/** Operations that change formats only: the calculation engine has nothing to do for them. */
export const FORMAT_OPS: ReadonlySet<Op["kind"]> = new Set([
  "format",
  "styles",
  "merge",
  "unmerge",
  "layout",
  "conditional",
  "validations",
  "annotations",
  "protect",
  "protectWorkbook",
  "theme",
  "sheetProps",
  "scenarios",
  "hyperlinks",
  "editRanges",
  "sparklines",
  "pivots",
])

/**
 * What an operation touched, for redraws and (from Phase 2) recalculation. `format` marks a change of formats only
 * (no values), which the engine skips.
 */
export type Change = { sheet: number; structural: boolean; range?: Range; format?: true }

export type Applied = { inverse: Op[]; change: Change }

/** Styles for ops that format; a workbook passes its own table (tests and the worker may leave it out). */
const fallbackStyles = new StyleTable()

export function applyOp(sheets: Sheet[], op: Op, styles: StyleTable = fallbackStyles): Applied {
  switch (op.kind) {
    case "addSheet":
    case "removeSheet":
    case "renameSheet":
    case "moveSheet":
    case "sheetProps":
      return applySheetOp(sheets, op)
    case "names": {
      const table = (sheets as SheetList).names
      if (!table) throw new Error("No name table")
      const before = table.copy()
      table.replace(op.names)
      // Formulas that use a name read it when they are registered: re-read them all.
      return { inverse: [{ kind: "names", names: before }], change: { sheet: 0, structural: true } }
    }
    case "externalLinks": {
      const box = (sheets as SheetList).links
      if (!box) throw new Error("No links")
      const before = copyExternalBooks(box.list)
      box.list = copyExternalBooks(op.links)
      // Every formula that reads a linked workbook calculates again.
      return { inverse: [{ kind: "externalLinks", links: before }], change: { sheet: 0, structural: true } }
    }
    case "theme": {
      const box = (sheets as SheetList).look
      if (!box) throw new Error("No theme")
      const before = copyTheme(box.theme)
      box.theme = copyTheme(op.theme)
      styles.retheme(box.theme)
      return { inverse: [{ kind: "theme", theme: before }], change: { sheet: 0, structural: false, format: true } }
    }
    case "protectWorkbook": {
      const box = (sheets as SheetList).structure
      if (!box) throw new Error("No workbook protection")
      const before = box.protection ? { ...box.protection } : null
      box.protection = op.protection ? { ...op.protection } : null
      return {
        inverse: [{ kind: "protectWorkbook", protection: before }],
        change: { sheet: 0, structural: false, format: true },
      }
    }
  }
  const sheet = sheets[op.sheet]
  if (!sheet) {
    throw new Error(`No sheet ${op.sheet}`)
  }
  switch (op.kind) {
    case "setCells":
      return setCells(sheet, op)
    case "insertRows":
      return insertLines(sheets, op.sheet, "rows", op.at, op.count, op.plain)
    case "deleteRows":
      return deleteLines(sheets, op.sheet, "rows", op.at, op.count)
    case "insertCols":
      return insertLines(sheets, op.sheet, "cols", op.at, op.count, op.plain)
    case "deleteCols":
      return deleteLines(sheets, op.sheet, "cols", op.at, op.count)
    case "insertCells":
      return insertCells(sheets, op.sheet, op.range, op.shift)
    case "deleteCells":
      return deleteCells(sheets, op.sheet, op.range, op.shift)
    case "axis":
      return setAxis(sheet, op.sheet, op.axis, op.entries)
    case "moveCells":
      return moveCells(sheets, op)
    case "format": {
      const record = formatRanges(sheet, styles, op.ranges, op.patch)
      return formatApplied(op.sheet, record)
    }
    case "styles":
      return formatApplied(op.sheet, restoreStyles(sheet, op.record))
    case "merge":
    case "unmerge":
      return mergeCells(sheet, op)
    case "layout":
      return setLayout(sheet, op)
    case "conditional": {
      const before = copyEntries(sheet.conditional)
      sheet.conditional = copyEntries(op.entries)
      return {
        inverse: [{ kind: "conditional", sheet: op.sheet, entries: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "validations": {
      const before = copyValidations(sheet.validations)
      sheet.validations = copyValidations(op.entries)
      return {
        inverse: [{ kind: "validations", sheet: op.sheet, entries: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "tables": {
      const before = copyTables(sheet.tables)
      sheet.tables = copyTables(op.tables)
      // Structured references now point elsewhere: formulas are re-read (as after rows move).
      return {
        inverse: [{ kind: "tables", sheet: op.sheet, tables: before }],
        change: { sheet: op.sheet, structural: true },
      }
    }
    case "objects": {
      const before = copyObjects(sheet.objects)
      sheet.objects = copyObjects(op.objects)
      return {
        inverse: [{ kind: "objects", sheet: op.sheet, objects: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "charts": {
      const before = copyCharts(sheet.charts)
      sheet.charts = copyCharts(op.charts)
      return {
        inverse: [{ kind: "charts", sheet: op.sheet, charts: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "pageSetup": {
      const before = copyPageSetup(sheet.pageSetup)
      sheet.pageSetup = copyPageSetup(op.setup)
      return {
        inverse: [{ kind: "pageSetup", sheet: op.sheet, setup: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "editRanges": {
      const before = copyEditRanges(sheet.editRanges)
      sheet.editRanges = copyEditRanges(op.ranges)
      return {
        inverse: [{ kind: "editRanges", sheet: op.sheet, ranges: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "pivots": {
      const before = copyPivotTables(sheet.pivots)
      sheet.pivots = copyPivotTables(op.tables)
      return {
        inverse: [{ kind: "pivots", sheet: op.sheet, tables: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "sparklines": {
      const before = copySparklineGroups(sheet.sparklines)
      sheet.sparklines = copySparklineGroups(op.groups)
      return {
        inverse: [{ kind: "sparklines", sheet: op.sheet, groups: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "hyperlinks": {
      const before = copyHyperlinks(sheet.hyperlinks)
      sheet.hyperlinks = copyHyperlinks(op.links)
      return {
        inverse: [{ kind: "hyperlinks", sheet: op.sheet, links: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "annotations": {
      const before: AnnotationWrite[] = op.writes.map((w) => ({
        row: w.row,
        col: w.col,
        value: sheet.annotations.get(w.row, w.col) ?? null,
      }))
      for (const w of op.writes) sheet.annotations.set(w.row, w.col, w.value)
      return {
        inverse: [{ kind: "annotations", sheet: op.sheet, writes: before.reverse() }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "autoFilter": {
      const before = copyAutoFilter(sheet.autoFilter)
      sheet.autoFilter = copyAutoFilter(op.filter)
      const rows = op.rows?.length ? setAxis(sheet, op.sheet, "rows", op.rows) : null
      const undoRows = rows ? (rows.inverse[0] as Extract<Op, { kind: "axis" }>).entries : undefined
      return {
        inverse: [{ kind: "autoFilter", sheet: op.sheet, filter: before, ...(undoRows ? { rows: undoRows } : {}) }],
        change: rows?.change.structural ? rows.change : { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "advancedFilter": {
      const before = copyAdvancedFilter(sheet.advancedFilter)
      sheet.advancedFilter = copyAdvancedFilter(op.filter)
      const rows = op.rows?.length ? setAxis(sheet, op.sheet, "rows", op.rows) : null
      const undoRows = rows ? (rows.inverse[0] as Extract<Op, { kind: "axis" }>).entries : undefined
      return {
        inverse: [{ kind: "advancedFilter", sheet: op.sheet, filter: before, ...(undoRows ? { rows: undoRows } : {}) }],
        change: rows?.change.structural ? rows.change : { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "queryTables": {
      const before = copyQueryTables(sheet.queryTables)
      sheet.queryTables = copyQueryTables(op.tables)
      return {
        inverse: [{ kind: "queryTables", sheet: op.sheet, tables: before }],
        change: { sheet: op.sheet, structural: false },
      }
    }
    case "consolidateSettings": {
      const before = copyConsolidate(sheet.consolidate)
      sheet.consolidate = copyConsolidate(op.settings)
      return {
        inverse: [{ kind: "consolidateSettings", sheet: op.sheet, settings: before }],
        change: { sheet: op.sheet, structural: false },
      }
    }
    case "outlineLevels": {
      const axis = axisOf(sheet, op.axis)
      const before: [number, number][] = op.levels.map(([i]) => [i, axis.level(i)])
      for (const [i, level] of op.levels) axis.setLevel(i, level)
      return {
        inverse: [{ kind: "outlineLevels", sheet: op.sheet, axis: op.axis, levels: before.reverse() }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "scenarios": {
      const before = copyScenarios(sheet.scenarios)
      sheet.scenarios = copyScenarios(op.scenarios)
      return {
        inverse: [{ kind: "scenarios", sheet: op.sheet, scenarios: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "outlineSettings": {
      const before = { ...sheet.outline }
      sheet.outline = { ...op.settings }
      return {
        inverse: [{ kind: "outlineSettings", sheet: op.sheet, settings: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
    case "protect": {
      const before = sheet.protection ? { ...sheet.protection } : null
      sheet.protection = op.protection ? { ...op.protection } : null
      return {
        inverse: [{ kind: "protect", sheet: op.sheet, protection: before }],
        change: { sheet: op.sheet, structural: false, format: true },
      }
    }
  }
}

function formatApplied(sheet: number, undo: StyleRecord): Applied {
  return {
    inverse: [{ kind: "styles", sheet, record: undo }],
    change: { sheet, structural: false, range: recordBounds(undo), format: true },
  }
}

/**
 * Merge: the ranges become merged cells, absorbing any merges they overlap (as Excel does when a selection contains
 * merged cells). Unmerge: every merge overlapping the ranges goes. The inverse puts back exactly what was there.
 */
function mergeCells(sheet: Sheet, op: Extract<Op, { kind: "merge" | "unmerge" }>): Applied {
  const removed: Range[] = []
  for (const r of op.ranges) {
    const hit = sheet.merges.overlapping(r)
    sheet.merges.remove(hit)
    removed.push(...hit)
  }
  if (op.kind === "merge") for (const r of op.ranges) sheet.merges.add(r)
  const touched = [...op.ranges, ...removed]
  const bounds = touched.reduce<Range | undefined>(
    (b, r) =>
      b
        ? { r0: Math.min(b.r0, r.r0), r1: Math.max(b.r1, r.r1), c0: Math.min(b.c0, r.c0), c1: Math.max(b.c1, r.c1) }
        : { ...r },
    undefined,
  )
  const inverse: Op[] = []
  if (op.kind === "merge") inverse.push({ kind: "unmerge", sheet: op.sheet, ranges: op.ranges })
  if (removed.length) inverse.push({ kind: "merge", sheet: op.sheet, ranges: removed })
  return { inverse, change: { sheet: op.sheet, structural: false, range: bounds, format: true } }
}

/** A snapshot of the sheet's row and column formats and merges. */
function layoutOf(sheet: Sheet, index: number): Extract<Op, { kind: "layout" }> {
  return {
    kind: "layout",
    sheet: index,
    rows: sheet.rowStyles.entries().flat(),
    cols: sheet.colStyles.entries().flat(),
    merges: sheet.merges.all.map((m) => ({ ...m })),
    conditional: copyEntries(sheet.conditional),
    validations: copyValidations(sheet.validations),
    tables: copyTables(sheet.tables),
    objects: copyObjects(sheet.objects),
    pageSetup: copyPageSetup(sheet.pageSetup),
    annotations: sheet.annotations.all().map((a) => ({ row: a.row, col: a.col, value: a.value })),
    autoFilter: copyAutoFilter(sheet.autoFilter),
    scenarios: copyScenarios(sheet.scenarios),
    advancedFilter: copyAdvancedFilter(sheet.advancedFilter),
    queryTables: copyQueryTables(sheet.queryTables),
    hyperlinks: copyHyperlinks(sheet.hyperlinks),
    editRanges: copyEditRanges(sheet.editRanges),
    sparklines: copySparklineGroups(sheet.sparklines),
    pivots: copyPivotTables(sheet.pivots),
  }
}

function setLayout(sheet: Sheet, op: Extract<Op, { kind: "layout" }>): Applied {
  const before = layoutOf(sheet, op.sheet)
  for (const [i] of sheet.rowStyles.entries()) sheet.rowStyles.set(i, 0)
  for (const [i] of sheet.colStyles.entries()) sheet.colStyles.set(i, 0)
  for (let i = 0; i < op.rows.length; i += 2) sheet.rowStyles.set(op.rows[i], op.rows[i + 1])
  for (let i = 0; i < op.cols.length; i += 2) sheet.colStyles.set(op.cols[i], op.cols[i + 1])
  sheet.merges.replace(op.merges.map((m) => ({ ...m })))
  if (op.conditional) sheet.conditional = copyEntries(op.conditional)
  if (op.validations) sheet.validations = copyValidations(op.validations)
  if (op.tables) sheet.tables = copyTables(op.tables)
  if (op.objects) sheet.objects = copyObjects(op.objects)
  else if (op.charts) sheet.charts = copyCharts(op.charts)
  if (op.pageSetup) sheet.pageSetup = copyPageSetup(op.pageSetup)
  if (op.annotations) {
    sheet.annotations.clear()
    for (const a of op.annotations) sheet.annotations.set(a.row, a.col, a.value)
  }
  if (op.autoFilter !== undefined) sheet.autoFilter = copyAutoFilter(op.autoFilter)
  if (op.scenarios) sheet.scenarios = copyScenarios(op.scenarios)
  if (op.advancedFilter !== undefined) sheet.advancedFilter = copyAdvancedFilter(op.advancedFilter)
  if (op.queryTables) sheet.queryTables = copyQueryTables(op.queryTables)
  if (op.hyperlinks) sheet.hyperlinks = copyHyperlinks(op.hyperlinks)
  if (op.editRanges) sheet.editRanges = copyEditRanges(op.editRanges)
  if (op.sparklines) sheet.sparklines = copySparklineGroups(op.sparklines)
  if (op.pivots) sheet.pivots = copyPivotTables(op.pivots)
  return { inverse: [before], change: { sheet: op.sheet, structural: true, format: true } }
}

/** Whether a structural edit has row/column formats or merges to carry along (and to snapshot for undo). */
function hasLayout(sheet: Sheet) {
  return (
    sheet.rowStyles.size > 0 ||
    sheet.colStyles.size > 0 ||
    sheet.merges.size > 0 ||
    sheet.conditional.length > 0 ||
    sheet.validations.length > 0 ||
    sheet.tables.length > 0 ||
    sheet.objects.length > 0 ||
    sheet.pageSetup.printArea !== null ||
    sheet.pageSetup.titleRows !== null ||
    sheet.pageSetup.titleCols !== null ||
    sheet.pageSetup.rowBreaks.length > 0 ||
    sheet.pageSetup.colBreaks.length > 0 ||
    sheet.annotations.size > 0 ||
    sheet.autoFilter !== null ||
    sheet.scenarios.scenarios.length > 0 ||
    sheet.scenarios.resultCells.length > 0 ||
    sheet.advancedFilter !== null ||
    sheet.queryTables.length > 0 ||
    sheet.hyperlinks.length > 0 ||
    sheet.editRanges.length > 0 ||
    sheet.sparklines.length > 0 ||
    sheet.pivots.length > 0
  )
}

/**
 * Rewrites the references of every formula in the workbook that points at sheet `target`, before the cells move.
 * Returns setCells ops that put the old texts back; they run after the structural inverse, which has moved the
 * cells back to where they were when the texts were recorded.
 */
function rewriteReferences(sheets: readonly Sheet[], target: number, map: RefMapper): Op[] {
  const targetName = sheets[target].name
  const undo: Op[] = []
  sheets.forEach((sheet, index) => {
    const edits: { row: number; col: number; text: string }[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      // Formulas on other sheets can only point here through a sheet prefix.
      if (index !== target && !input.includes("!")) return
      const text = rewriteFormula(input, sheet.name, targetName, map)
      if (text !== null) edits.push({ row, col, text })
    })
    if (!edits.length) return
    const previous: CellWrite[] = []
    for (const edit of edits) {
      const data = sheet.cells.get(edit.row, edit.col)!
      previous.push({ row: edit.row, col: edit.col, data })
      sheet.cells.set(edit.row, edit.col, { ...data, input: edit.text })
    }
    undo.push({ kind: "setCells", sheet: index, cells: previous })
  })
  // Charts' series ranges (always with their sheet's name) follow too.
  sheets.forEach((sheet, index) => {
    if (!sheet.charts.length) return
    const before = copyCharts(sheet.charts)
    let changed = false
    for (const chart of sheet.charts)
      if (rewriteChartRefs(chart, (ref) => rewriteFormula(ref, sheet.name, targetName, map))) changed = true
    if (changed) undo.push({ kind: "charts", sheet: index, charts: before })
  })
  // So do sparklines' data and date ranges.
  sheets.forEach((sheet, index) => {
    if (!sheet.sparklines.length) return
    const before = copySparklineGroups(sheet.sparklines)
    if (rewriteSparklineRefs(sheet.sparklines, (ref) => rewriteFormula(ref, sheet.name, targetName, map)))
      undo.push({ kind: "sparklines", sheet: index, groups: before })
  })
  const host = sheets[0]?.name ?? targetName
  const names = rewriteNames(sheets as SheetList, (ref) => rewriteFormula(ref, host, targetName, map))
  if (names) undo.push({ kind: "names", names })
  return undo
}

function setCells(sheet: Sheet, op: Extract<Op, { kind: "setCells" }>): Applied {
  const previous: CellWrite[] = []
  let r0 = MAX_ROWS
  let c0 = MAX_COLS
  let r1 = -1
  let c1 = -1
  for (const write of op.cells) {
    previous.push({ row: write.row, col: write.col, data: sheet.cells.get(write.row, write.col) })
    r0 = Math.min(r0, write.row)
    r1 = Math.max(r1, write.row)
    c0 = Math.min(c0, write.col)
    c1 = Math.max(c1, write.col)
  }
  for (const write of op.cells) {
    sheet.cells.set(write.row, write.col, write.data)
  }
  // Undo writes the old contents in reverse, so a cell written twice ends with its first old content.
  previous.reverse()
  // A header cell typed over renames its table column; structured references then read again.
  const renamed = r1 >= 0 && syncTableHeaders(sheet, { r0, c0, r1, c1 })
  return {
    inverse: [{ kind: "setCells", sheet: op.sheet, cells: previous }],
    change: { sheet: op.sheet, structural: renamed, range: r1 >= 0 ? { r0, c0, r1, c1 } : undefined },
  }
}

function axisOf(sheet: Sheet, axis: AxisName): Axis {
  return axis === "rows" ? sheet.rows : sheet.cols
}

function lineRange(axis: AxisName, from: number, to: number): Range {
  return axis === "rows" ? { r0: from, r1: to, c0: 0, c1: MAX_COLS - 1 } : { r0: 0, r1: MAX_ROWS - 1, c0: from, c1: to }
}

function insertLines(
  sheets: readonly Sheet[],
  index: number,
  axis: AxisName,
  at: number,
  count: number,
  plain = false,
): Applied {
  const sheet = sheets[index]
  const size = axis === "rows" ? MAX_ROWS : MAX_COLS
  // Lines pushed off the end are lost; keep them so undo can restore them.
  const dropped = sheet.cells.entries(lineRange(axis, size - count, size - 1))
  const droppedAxis = axisOf(sheet, axis).overrides(size - count, size - 1)
  const texts = rewriteReferences(sheets, index, shiftMapper(axis, at, count))
  const layout = hasLayout(sheet) ? layoutOf(sheet, index) : null
  if (axis === "rows") {
    sheet.cells.shiftRows(at, count)
  } else {
    sheet.cells.shiftCols(at, count)
  }
  axisOf(sheet, axis).insert(at, count)
  ;(axis === "rows" ? sheet.rowStyles : sheet.colStyles).insert(at, count)
  shiftMerges(sheet, axis, at, count)
  shiftConditional(sheet, axis, at, count)
  shiftValidations(sheet, axis, at, count)
  shiftScenarios(sheet, axis, at, count)
  shiftTables(sheet, axis, at, count)
  shiftObjects(sheet.objects, axis, at, count)
  sheet.pageSetup = shiftPageSetup(sheet.pageSetup, axis, at, count)
  sheet.annotations.shift(axis, at, count)
  if (sheet.autoFilter) sheet.autoFilter = shiftAutoFilter(sheet.autoFilter, axis, at, count)
  if (sheet.advancedFilter) sheet.advancedFilter = shiftAdvancedFilter(sheet.advancedFilter, index, axis, at, count)
  if (sheet.queryTables.length) sheet.queryTables = shiftQueryTables(sheet.queryTables, axis, at, count)
  if (sheet.hyperlinks.length) sheet.hyperlinks = shiftHyperlinks(sheet.hyperlinks, axis, at, count)
  if (sheet.editRanges.length) sheet.editRanges = shiftEditRanges(sheet.editRanges, axis, at, count)
  if (sheet.sparklines.length) sheet.sparklines = shiftSparklines(sheet.sparklines, axis, at, count)
  if (sheet.pivots.length) sheet.pivots = shiftPivots(sheet.pivots, axis, at, count)
  if (!plain) copyFormatsInto(sheet, axis, at, count)
  const inverse: Op[] = [{ kind: axis === "rows" ? "deleteRows" : "deleteCols", sheet: index, at, count }]
  if (layout) inverse.push(layout)
  if (dropped.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: dropped })
  }
  if (droppedAxis.length) {
    inverse.push({ kind: "axis", sheet: index, axis, entries: droppedAxis })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

/**
 * Excel's default for inserted rows and columns ("Format Same As Above / Left"): the new cells take the formats of
 * the cells in the line before them. Only formats; the inverse (deleting the lines) takes them away again.
 */
function copyFormatsInto(sheet: Sheet, axis: AxisName, at: number, count: number) {
  if (at === 0) return
  const source = lineRange(axis, at - 1, at - 1)
  const styled: number[] = []
  sheet.cells.forEachStyle(source, (row, col, style) => {
    if (style) styled.push(axis === "rows" ? col : row, style)
  })
  for (let i = 0; i < styled.length; i += 2) {
    for (let k = at; k < at + count; k++) {
      if (axis === "rows") sheet.cells.setStyle(k, styled[i], styled[i + 1])
      else sheet.cells.setStyle(styled[i], k, styled[i + 1])
    }
  }
}

function deleteLines(sheets: readonly Sheet[], index: number, axis: AxisName, at: number, count: number): Applied {
  const sheet = sheets[index]
  const range = lineRange(axis, at, at + count - 1)
  const removed = sheet.cells.entries(range)
  const removedAxis = axisOf(sheet, axis).overrides(at, at + count - 1)
  const texts = rewriteReferences(sheets, index, shiftMapper(axis, at, -count))
  const layout = hasLayout(sheet) ? layoutOf(sheet, index) : null
  sheet.cells.clear(range)
  if (axis === "rows") {
    sheet.cells.shiftRows(at + count, -count)
  } else {
    sheet.cells.shiftCols(at + count, -count)
  }
  axisOf(sheet, axis).remove(at, count)
  ;(axis === "rows" ? sheet.rowStyles : sheet.colStyles).remove(at, count)
  shiftMerges(sheet, axis, at, -count)
  shiftConditional(sheet, axis, at, -count)
  shiftValidations(sheet, axis, at, -count)
  shiftScenarios(sheet, axis, at, -count)
  shiftTables(sheet, axis, at, -count)
  shiftObjects(sheet.objects, axis, at, -count)
  sheet.pageSetup = shiftPageSetup(sheet.pageSetup, axis, at, -count)
  sheet.annotations.shift(axis, at, -count)
  if (sheet.autoFilter) sheet.autoFilter = shiftAutoFilter(sheet.autoFilter, axis, at, -count)
  if (sheet.advancedFilter) sheet.advancedFilter = shiftAdvancedFilter(sheet.advancedFilter, index, axis, at, -count)
  if (sheet.queryTables.length) sheet.queryTables = shiftQueryTables(sheet.queryTables, axis, at, -count)
  if (sheet.hyperlinks.length) sheet.hyperlinks = shiftHyperlinks(sheet.hyperlinks, axis, at, -count)
  if (sheet.editRanges.length) sheet.editRanges = shiftEditRanges(sheet.editRanges, axis, at, -count)
  if (sheet.sparklines.length) sheet.sparklines = shiftSparklines(sheet.sparklines, axis, at, -count)
  if (sheet.pivots.length) sheet.pivots = shiftPivots(sheet.pivots, axis, at, -count)
  const inverse: Op[] = [{ kind: axis === "rows" ? "insertRows" : "insertCols", sheet: index, at, count, plain: true }]
  if (layout) inverse.push(layout)
  if (removed.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: removed })
  }
  if (removedAxis.length) {
    inverse.push({ kind: "axis", sheet: index, axis, entries: removedAxis })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function insertCells(sheets: readonly Sheet[], index: number, range: Range, shift: "down" | "right"): Applied {
  const sheet = sheets[index]
  let dropped
  let texts: Op[]
  if (shift === "down") {
    const n = range.r1 - range.r0 + 1
    texts = rewriteReferences(sheets, index, shiftMapper("rows", range.r0, n, [range.c0, range.c1]))
    dropped = sheet.cells.entries({ r0: MAX_ROWS - n, r1: MAX_ROWS - 1, c0: range.c0, c1: range.c1 })
    sheet.cells.shiftRows(range.r0, n, range.c0, range.c1)
  } else {
    const n = range.c1 - range.c0 + 1
    texts = rewriteReferences(sheets, index, shiftMapper("cols", range.c0, n, [range.r0, range.r1]))
    dropped = sheet.cells.entries({ r0: range.r0, r1: range.r1, c0: MAX_COLS - n, c1: MAX_COLS - 1 })
    sheet.cells.shiftCols(range.c0, n, range.r0, range.r1)
  }
  const inverse: Op[] = [{ kind: "deleteCells", sheet: index, range, shift: shift === "down" ? "up" : "left" }]
  if (dropped.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: dropped })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function deleteCells(sheets: readonly Sheet[], index: number, range: Range, shift: "up" | "left"): Applied {
  const sheet = sheets[index]
  const removed = sheet.cells.entries(range)
  const texts =
    shift === "up"
      ? rewriteReferences(
          sheets,
          index,
          shiftMapper("rows", range.r0, -(range.r1 - range.r0 + 1), [range.c0, range.c1]),
        )
      : rewriteReferences(
          sheets,
          index,
          shiftMapper("cols", range.c0, -(range.c1 - range.c0 + 1), [range.r0, range.r1]),
        )
  sheet.cells.clear(range)
  if (shift === "up") {
    sheet.cells.shiftRows(range.r1 + 1, -(range.r1 - range.r0 + 1), range.c0, range.c1)
  } else {
    sheet.cells.shiftCols(range.c1 + 1, -(range.c1 - range.c0 + 1), range.r0, range.r1)
  }
  const inverse: Op[] = [{ kind: "insertCells", sheet: index, range, shift: shift === "up" ? "down" : "right" }]
  if (removed.length) {
    inverse.push({ kind: "setCells", sheet: index, cells: removed })
  }
  inverse.push(...texts)
  return { inverse, change: { sheet: index, structural: true } }
}

function setAxis(sheet: Sheet, index: number, name: AxisName, entries: AxisEntry[]): Applied {
  const axis = axisOf(sheet, name)
  let visibility = false
  const previous: AxisEntry[] = entries.map((entry) => ({
    index: entry.index,
    size: axis.hasCustomSize(entry.index) ? axis.storedSize(entry.index) : null,
    hidden: axis.isHidden(entry.index),
    ...(axis.isAuto(entry.index) ? { auto: true as const } : {}),
    ...(entry.level !== undefined
      ? { level: axis.level(entry.index), ...(axis.isCollapsed(entry.index) ? { collapsed: true as const } : {}) }
      : {}),
  }))
  for (const entry of entries) {
    if (axis.isHidden(entry.index) !== entry.hidden) visibility = true
    axis.setSize(entry.index, entry.size, entry.auto)
    axis.setHidden(entry.index, entry.hidden)
    // Entries without a level (sizes, Hide) leave the outline as it is.
    if (entry.level !== undefined) {
      axis.setLevel(entry.index, entry.level)
      axis.setCollapsed(entry.index, Boolean(entry.collapsed))
    }
  }
  previous.reverse()
  return {
    inverse: [{ kind: "axis", sheet: index, axis: name, entries: previous }],
    // Sizes alone change no value; hiding does (SUBTOTAL 101–111 and AGGREGATE skip hidden rows).
    change: visibility ? { sheet: index, structural: true } : { sheet: index, structural: false, format: true },
  }
}

function moveCells(sheets: readonly Sheet[], op: Extract<Op, { kind: "moveCells" }>): Applied {
  const { source, dr, dc } = op
  const sheet = sheets[op.sheet]
  // The texts' undo goes last: it restores formulas at their old places once the cells are back.
  const texts = rewriteReferences(sheets, op.sheet, moveMapper(source, dr, dc))
  const moved = sheet.cells.entries(source)
  const dest = { r0: source.r0 + dr, r1: source.r1 + dr, c0: source.c0 + dc, c1: source.c1 + dc }
  const previous = new Map<string, CellWrite>()
  const remember = (row: number, col: number) => {
    const key = `${row},${col}`
    if (!previous.has(key)) previous.set(key, { row, col, data: sheet.cells.get(row, col) })
  }
  for (const entry of moved) remember(entry.row, entry.col)
  for (const entry of sheet.cells.entries(dest)) remember(entry.row, entry.col)
  for (const entry of moved) remember(entry.row + dr, entry.col + dc)
  sheet.cells.clear(source)
  sheet.cells.clear(dest)
  for (const entry of moved) {
    const data = { ...entry.data }
    delete data.spill
    sheet.cells.set(entry.row + dr, entry.col + dc, data)
  }
  const inverse: Op[] = [{ kind: "setCells", sheet: op.sheet, cells: [...previous.values()] }, ...texts]
  return { inverse, change: { sheet: op.sheet, structural: true } }
}
