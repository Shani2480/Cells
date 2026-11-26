import { GetData } from "./get-data"
import { PictureData } from "./picture-data"
import { SlicerCommands } from "./slicers"
import { TimelineCommands } from "./timelines"
import { QuickAnalysis } from "./quick-analysis"
import { SparklineCommands, type SparklineDialogInit } from "./sparklines"
import { CustomStyleCommands } from "./custom-styles"
import { PivotCommands, type PivotFilterRequest } from "./pivots"
import { ObjectCommands } from "./objects"
import { ChartCommands } from "./charts"
import { PrintCommands } from "./print"
import { HeaderFooterEditor } from "./header-footer"
import { TableCommands, type CreateTableInit } from "./tables"
import { DEFAULT_VIEW as DEFAULT_FILE_VIEW, type SheetView as FileSheetView } from "@cells/io"
import {
  MAX_COLS,
  PROTECTED_MESSAGE,
  breaksDataTable,
  breaksPivot,
  isDataTableInput,
  pivotAt,
  PIVOT_MESSAGE,
  canEditCell,
  protectionBlocks,
  structureBlocks,
  MAX_ROWS,
  addRange,
  blockFromText,
  canPointAt,
  cellName,
  clampCell,
  clearContentsOp,
  currentRegion,
  dataEdge,
  decodeTSV,
  editReferences,
  editTextFor,
  encodeTSV,
  expandToColumns,
  expandToRows,
  extendSelection,
  fillInputOp,
  fillOp,
  formatValue,
  functionsStartingWith,
  growToMerges,
  isSingleCellOn,
  mergeAnchor,
  isWholeColumns,
  isWholeRows,
  lastCell,
  moveBlockOps,
  moveWithinSelection,
  nameAtCaret,
  pasteArea,
  pointedText,
  pasteBlockOp,
  cutLinksOp,
  pasteLinksOp,
  pasteNotesOp,
  readHyperlinks,
  readNotes,
  rangeCols,
  rangeName,
  rangeOf,
  rangeRows,
  readBlock,
  sameRange,
  selectAll,
  selectCell,
  selectCols,
  selectRange,
  selectRows,
  selectionBounds,
  step,
  toggleAbsolute,
  writeInputOp,
  type EditReference,
  type FunctionInfo,
  type CellAddress,
  type DrawingObject,
  type CellBlock,
  type CellWrite,
  type CFEntry,
  type CFResult,
  ConditionalEval,
  StyleTable,
  dateSerial,
  type Direction,
  type AxisEntry,
  type Op,
  type UndoEntry,
  type Range,
  type Selection,
  type Sheet,
  Workbook,
  quoteSheet,
  rangeHtml,
  type Pagination,
  canSelect,
  nextSelectable,
  selectableCells,
} from "@cells/core"
import { openLink } from "@/links/link"
import { FormatCommands, NUMBER_FORMATS } from "./format"
import { AnnotationCommands } from "./annotations"
import { HyperlinkCommands, type LinkTarget } from "./hyperlinks"
import { AuditCommands } from "./audit"
import { OutlineCommands } from "./outline"
import { DataCommands, type FilterDialogRequest } from "./data"
import { ValidationCommands, type AlertAnswer, type ValidationAlert, type ValidationDialogInit } from "./validation"
import { FindCommands } from "./find"
import { NameCommands } from "./names"
import { ConditionalCommands } from "./conditional"
import { GROUP_MESSAGE, SheetGroup } from "./group"
import { CustomViewCommands } from "./custom-views"
import { EditRangeCommands } from "./edit-ranges"
import { PAGE_GAP, pageAxis, plainMetric, type AxisMetric, type PageAxis } from "./page-map"
import { SHEET_MESSAGES, SheetCommands } from "./sheets"

/**
 * Ready: the grid has the keys. Enter: typing a fresh entry (arrows commit). Edit: F2 editing (arrows move the
 * caret). Point: a formula is pointing at cells (arrows and clicks write a reference).
 */
export type CellMode = "ready" | "enter" | "edit" | "point"

/** The reference Point mode is writing: where it sits in the text and the cells it covers. */
export type PointState = { start: number; end: number; anchor: CellAddress; focus: CellAddress }

/** More an entry's commit does, in the same step to undo (live Flash Fill's cells), and that step's name. */
export type CommitExtra = { ops: Op[]; label: string }

export type EditState = {
  row: number
  col: number
  text: string
  mode: "enter" | "edit"
  /** Where typing happens: in the cell or in the formula bar (Excel keeps both in step). */
  surface: "cell" | "bar"
  /** The editor's selection in the text (equal when it is a caret). */
  selStart: number
  selEnd: number
  point: PointState | null
  /** The highlighted row of the function list. */
  pick: number
}

/** The function list under the editor while a function name is typed. */
export type Autocomplete = { start: number; prefix: string; items: FunctionInfo[]; pick: number }

/** The marching-ants range after Copy or Cut, until Esc, a new entry or a paste of a cut. */
/** `sheet`: where the copy came from (a cut pasted on another sheet moves the cells there). */
export type ClipState = { range: Range; cut: boolean; text: string; block: CellBlock; sheet: Sheet }

/**
 * Frozen panes: rows above `row` and columns left of `col` stay put. They show from `top` / `left` (the first row and
 * column in view when the panes were frozen). `row === top` freezes no rows, `col === left` no columns.
 */
export type Freeze = { row: number; col: number; top: number; left: number }

/**
 * View › Split: the window in two or four panes that scroll on their own. `width` and `height` are the left and top
 * panes' size in sheet pixels (0: no split that way); `scrollX` / `scrollY` are where those panes are scrolled to.
 * The bottom-right pane is the controller's own scroll.
 */
export type Split = { width: number; height: number; scrollX: number; scrollY: number }

/** View › Workbook Views: Normal, Page Break Preview or Page Layout (per sheet, as Excel). */
export type SheetViewMode = "normal" | "pageBreakPreview" | "pageLayout"

/** The zoom a view starts at the first time a sheet is shown in it (Excel: 60% for Page Break Preview). */
const VIEW_ZOOM: Record<SheetViewMode, number> = { normal: 1, pageBreakPreview: 0.6, pageLayout: 1 }

/** Excel's zoom range. */
export const MIN_ZOOM = 0.1
export const MAX_ZOOM = 4

/** What each sheet remembers while another one is shown (Excel keeps these per sheet). */
type SheetView = {
  selection: Selection
  scrollX: number
  scrollY: number
  freeze: Freeze | null
  split: Split | null
  zoom: number
  /** View › Workbook Views, with the zoom each of the other views had last. */
  view: SheetViewMode
  viewZooms: Partial<Record<SheetViewMode, number>>
  showGridlines: boolean
  showHeadings: boolean
  showRuler: boolean
}

/** Undo state: the selection, and the sheet it was on. */
export type ViewMeta = Selection & { sheet?: number }

export type KeyInput = {
  key: string
  /** Ctrl on Linux/Windows, ⌘ on macOS. */
  primary: boolean
  shift: boolean
  alt: boolean
}

export type InsertDeleteRequest = { kind: "insert" | "delete" }

/** Things the controller asks the shell to do (dialogs and the like). */
export type ShellRequests = {
  /** The shown workbook's full path (for copies that say where they came from); null before it is saved. */
  documentPath?(): string | null
  goTo(): void
  insertDelete(request: InsertDeleteRequest): void
  /** Unlock Range: the password of the sheet's edit range `name`; `then` runs once it's open. */
  unlockRange?(name: string, then: () => void): void
  /** Ask before going on (Excel's OK/Cancel warnings); `ok` runs on OK. */
  confirm(message: string, ok: () => void): void
  /** A message with only OK (Excel's refusals); `then` runs when it closes (instead of focusing the sheet). */
  alert(message: string, then?: () => void): void
  /** Shift+F3: Insert Function. */
  insertFunction?(): void
  /** One of the shell's dialogs by name, as the ribbon opens them (Ctrl+F3: "nameManager"). */
  openDialog?(name: string): void
  /** Excel's Sort Warning: expand the selection to the data beside it, or sort only the selection. */
  sortWarning?(then: (expand: boolean) => void, kind?: "sort" | "duplicates"): void
  /** Data › Sort: the Sort dialog, on a range. */
  sortDialog?(range: Range): void
  /** Remove Duplicates' dialog for a range. */
  removeDuplicates?(range: Range): void
  /** Data › Text to Columns: the Convert Text to Columns Wizard for a column of cells and their texts. */
  textToColumns?(range: Range, texts: string[]): void
  /** Find and Replace, on its Find or Replace tab. */
  findReplace?(tab: "find" | "replace"): void
  /** Go To Special. */
  goToSpecial?(): void
  /** Custom AutoFilter or Top 10 AutoFilter, from a filter dropdown. */
  filterDialog?(request: FilterDialogRequest): void
  /** A PivotTable filter dropdown's Label Filter, Value Filter or Top 10 dialog. */
  pivotFilterDialog?(request: PivotFilterRequest): void
  /** Data › Data Validation: the dialog, with the rule it opens on. */
  dataValidation?(init: ValidationDialogInit): void
  /** The error alert after an entry breaks a rule; `answer` gets the button pressed. */
  validationAlert?(alert: ValidationAlert, answer: (a: AlertAnswer) => void): void
  /** Alt+Down on a cell with an in-cell list: open it. */
  /** Opens the active cell's in-cell list; `pick`: the column's own entries (Pick From Drop-down List). */
  openCellList?(pick?: boolean): void
  /** Insert › Table (Ctrl+T): the Create Table dialog. */
  createTable?(init: CreateTableInit): void
  /** Selects a ribbon tab (Chart Design after a chart is inserted). */
  showTab?(id: string): void
  /** Chart Design › Select Data: the Select Data Source dialog for the selected chart. */
  chartData?(): void
  /** Table Design › Resize Table: the dialog, on the table's range. */
  resizeTable?(range: Range): void
  /** Insert › Sparklines and the Sparkline tab's Edit Data. */
  sparklineDialog?(init: SparklineDialogInit): void
  /** A link to follow outside this sheet: open a workbook in Cells, or hand a web page, mail or file to the system. */
  followLink?(target: LinkTarget): void
  /** Ctrl+1: the Format Cells dialog, on a tab. */
  formatCells(tab?: "number" | "alignment" | "font" | "border" | "fill" | "protection"): void
}

export const HEADER_HEIGHT = 20
/** The headings on a phone (the phone shell), big enough to tap a column or row. */
export const LARGE_HEADER_HEIGHT = 30

/** Format changes that can change a row's fitted height. */
const FIT_KEYS = new Set(["fontSize", "fontName", "wrap", "rotation", "indent", "numberFormat"])
/** Edits touching more rows than this don't refit rows (see fitRows). */
const FIT_LIMIT = 20_000
/** The smallest pane a split leaves (sheet pixels); a split bar dragged closer to an edge goes away. */
export const SPLIT_MIN = 24

/** Today as a date serial (local time), for "A Date Occurring" rules. */
export function todaySerial(): number {
  const now = new Date()
  return dateSerial(now.getFullYear(), now.getMonth() + 1, now.getDate())
}

/** Clears every cell of a range (contents and formats): the source of a cut pasted on another sheet. */
function clearCellsOp(sheetIndex: number, sheet: Sheet, range: Range): Op {
  const cells: CellWrite[] = sheet.cells.entries(range).map((e) => ({ row: e.row, col: e.col, data: undefined }))
  return { kind: "setCells", sheet: sheetIndex, cells }
}

/** The notes and comments of a range removed: the source of a cut pasted on another sheet. */
function clearNotesOp(sheetIndex: number, sheet: Sheet, range: Range): Op | null {
  const writes = sheet.annotations.within(range).map((a) => ({ row: a.row, col: a.col, value: null }))
  return writes.length ? { kind: "annotations", sheet: sheetIndex, writes } : null
}

const optional = (op: Op | null): Op[] => (op ? [op] : [])

/** Excel's message when an edit would change some of a What-If data table's cells. */
const DATA_TABLE_MESSAGE = "Cannot change part of a data table."

/**
 * The state and behaviour of one sheet view: selection, in-cell editing, scrolling, clipboard and every keyboard
 * command, on top of the core workbook. No React and no DOM: views subscribe and draw.
 */
export class SheetController {
  readonly wb: Workbook<Selection>
  sheetIndex = 0
  selection: Selection = selectCell({ row: 0, col: 0 })
  edit: EditState | null = null
  clip: ClipState | null = null
  endMode = false
  /** Ctrl+`: cells show their formulas instead of their values. */
  showFormulas = false
  /** View › Freeze Panes (per sheet). */
  freeze: Freeze | null = null
  /** View › Split (per sheet); never together with frozen panes. */
  split: Split | null = null
  /** View › Zoom: 1 is 100% (per sheet, and per view). */
  zoom = 1
  /** View › Workbook Views (per sheet). */
  view: SheetViewMode = "normal"
  /** The zoom of the views not shown, for when they come back. */
  viewZooms: Partial<Record<SheetViewMode, number>> = {}
  /** View › Gridlines. */
  showGridlines = true
  /** View › Headings (row numbers and column letters). */
  showHeadings = true
  /** View › Formula Bar. */
  showFormulaBar = true
  /** The formula bar shows several lines (Ctrl+Shift+U, or its chevron); how many is set by dragging its edge. */
  formulaBarExpanded = false
  formulaBarLines = 3
  /** Preview of a fill-handle drag (the area it will fill). */
  fillPreview: Range | null = null
  /**
   * The last pointer was a finger: the selection shows round grippers instead of the fill handle, and the on-screen
   * keyboard waits for editing to start. Before any pointer, a touch screen's (a tablet's) is assumed, so opening a
   * workbook there doesn't bring the keyboard up.
   */
  touchHandles = primaryPointerIsTouch()
  /** Finger-sized column and row headings (the phone shell sets it while shown). */
  largeHeaders = false
  /** Scroll position in sheet pixels at 100% (top-left of the cell area). */
  scrollX = 0
  scrollY = 0
  /** Size of the cell area (without headers), set by the view, in sheet pixels (at 100%; the zoom is applied after). */
  viewWidth = 800
  viewHeight = 600
  rowHeaderWidth = 40
  shell: ShellRequests = { goTo() {}, insertDelete() {}, confirm: (_m, ok) => ok(), alert() {}, formatCells() {} }
  /** Home-tab formatting on the selection. */
  readonly format: FormatCommands = new FormatCommands(this)
  /** Home › Conditional Formatting. */
  readonly conditional: ConditionalCommands = new ConditionalCommands(this)
  /** Sheet tabs: add, delete, rename, move, copy, colour, hide. */
  readonly sheets: SheetCommands = new SheetCommands(this)
  /** View › Ruler (Page Layout only). */
  showRuler = true

  /** The file's name and folder, for the header and footer codes &F and &Z (Page Layout shows them). */
  printFile = { fileName: "Book1.xlsx", folder: "" }

  private pageMapCache: { pagination: Pagination; rows: PageAxis; cols: PageAxis } | null = null

  /** View › Page Layout: where the rows and columns sit on the pages (null in the other views). */
  pageMap(): { rows: PageAxis; cols: PageAxis } | null {
    if (this.view !== "pageLayout") return null
    const pagination = this.print.current()
    if (this.pageMapCache?.pagination !== pagination)
      this.pageMapCache = { pagination, ...pagesOf(this.sheet, pagination) }
    return this.pageMapCache
  }

  /** How a sheet's view places its rows and columns (any sheet, any view: for its stored scroll position). */
  private metricsFor(sheet: Sheet, view: SheetViewMode): { rows: AxisMetric; cols: AxisMetric } {
    if (view !== "pageLayout") return { rows: plainMetric(sheet.rows), cols: plainMetric(sheet.cols) }
    if (sheet === this.sheet && this.view === "pageLayout") return this.pageMap()!
    return pagesOf(sheet, this.print.pages(sheet))
  }

  /** Rows and columns as the view places them: on pages in Page Layout, one after another otherwise. */
  get rowMetric(): AxisMetric {
    return this.pageMap()?.rows ?? plainMetric(this.sheet.rows)
  }

  get colMetric(): AxisMetric {
    return this.pageMap()?.cols ?? plainMetric(this.sheet.cols)
  }

  /** A change that isn't an undo step but still needs saving (the open document counts it). */
  unsavedChange: () => void = () => this.changed()
  /** Review › Check Accessibility: the Accessibility pane is open. */
  accessibilityOpen = false

  setAccessibilityPane(open = !this.accessibilityOpen) {
    this.accessibilityOpen = open
    this.changed()
  }

  /** A picture of objects for the clipboard (a PNG data URL), when the view can make one. */
  objectPicture: ((objects: DrawingObject[]) => string | null) | null = null

  /** Review › Thesaurus: the Thesaurus pane is open. */
  thesaurusOpen = false

  setThesaurusPane(open = !this.thesaurusOpen) {
    this.thesaurusOpen = open
    this.changed()
  }

  /** Help › Help (F1) or Show Training: the Help pane is open, on its help or on the tour. */
  helpPane: "help" | "training" | null = null

  setHelpPane(open: "help" | "training" | null) {
    this.helpPane = open
    this.changed()
  }

  /** Puts text into a cell as if typed there (Thesaurus › Insert), as one undo step. */
  typeInto(at: CellAddress, text: string, label: string) {
    const before = this.selection
    this.transact([writeInputOp(this.sheetIndex, this.sheet, this.wb.styles, at.row, at.col, text)], {
      label,
      before,
      after: before,
    })
  }

  /** Review › Allow Edit Ranges. */
  readonly editRanges: EditRangeCommands = new EditRangeCommands(this)
  /** View › Custom Views. */
  readonly customViews: CustomViewCommands = new CustomViewCommands(this)
  /** Grouped sheets: edits on the active sheet repeat on each of them. */
  readonly group: SheetGroup = new SheetGroup(this)
  /** Notes and comments. */
  readonly annotations: AnnotationCommands = new AnnotationCommands(this)
  readonly data: DataCommands = new DataCommands(this)
  /** Insert › Link and following links. */
  readonly hyperlinks: HyperlinkCommands = new HyperlinkCommands(this)
  /** Insert › Sparklines and the Sparkline tab. */
  readonly sparklines: SparklineCommands = new SparklineCommands(this)
  readonly pivots: PivotCommands = new PivotCommands(this)
  readonly pivotStyles: CustomStyleCommands = new CustomStyleCommands(this, "pivot")
  readonly tableStyles: CustomStyleCommands = new CustomStyleCommands(this, "table")
  /** Quick Analysis (Ctrl+Q): the gallery's choices, their previews and the changes they make. */
  readonly quick: QuickAnalysis = new QuickAnalysis(this)
  readonly slicers: SlicerCommands = new SlicerCommands(this)
  readonly timelines: TimelineCommands = new TimelineCommands(this)
  readonly audit: AuditCommands = new AuditCommands(this)
  readonly outline: OutlineCommands = new OutlineCommands(this)
  /** Data › Data Validation. */
  readonly validation: ValidationCommands = new ValidationCommands(this)
  /** Insert › Table, Format as Table and the Table Design tab. */
  readonly tables: TableCommands = new TableCommands(this)
  /** Data › Get & Transform Data: queries (Power Query) and the Query Editor's state. */
  readonly getData: GetData = new GetData(this)
  readonly pictureData: PictureData = new PictureData(this)
  /** Charts: inserting, selecting, moving and the Chart Design tab. */
  readonly charts: ChartCommands = new ChartCommands(this)
  /** Shapes, text boxes, pictures and groups (and charts, as objects): Insert › Illustrations, Shape Format. */
  readonly objects: ObjectCommands = new ObjectCommands(this)
  /** Printing: Page Layout, Page Setup, print area, page breaks and the pages File › Print shows. */
  readonly print: PrintCommands = new PrintCommands(this)
  /** Page Layout: the header or footer being edited on the page. */
  readonly headerFooter: HeaderFooterEditor = new HeaderFooterEditor(this)
  /** Find, Replace and Go To Special. */
  readonly find: FindCommands = new FindCommands(this)
  /** Defined names (Name Box, Define Name, Go To). */
  readonly names: NameCommands = new NameCommands(this)
  /** What a cell shows in its column (the view knows the fonts and widths); used by Increase/Decrease Decimal. */
  shownText: ((row: number, col: number) => string) | null = null
  /** AutoFit Row Height / Column Width on the selection (the view measures text). */
  autofit: ((axis: "rows" | "cols") => void) | null = null
  /** The grid's row fitter (it can measure text): the height a row's content needs, null for the default. */
  rowFit: ((row: number) => number | null) | null = null
  /** The grid's column fitter: the width a column's content needs (within `rows` when given), null for the default. */
  columnFit: ((col: number, rows?: { r0: number; r1: number }, sheet?: number) => number | null) | null = null
  /** Styles conditional formatting draws with (kept apart from the workbook's own table). */
  readonly cfStyles = new StyleTable()
  private cf: {
    sheet: Sheet
    version: number
    day: number
    preview: CFEntry | null
    ev: ConditionalEval
  } | null = null
  /** Bumped on every change the view may need to draw; the React hook compares it. */
  version = 0

  private readonly listeners = new Set<() => void>()
  private readonly views = new WeakMap<Sheet, SheetView>()
  /** The sheet whose selection and scroll are live (its index can change when sheets move). */
  private shown: Sheet
  /** Tab moves remember where they started, so Enter returns to that column on the next row (Excel). */
  private tabStartCol: number | null = null
  /** Ctrl+A in data selects the region first, the whole sheet on the next press. */
  private lastCtrlARegion: Range | null = null

  constructor(wb: Workbook<Selection> = new Workbook<Selection>()) {
    this.wb = wb
    this.shown = wb.sheets[0]
    this.stopListening = wb.subscribe(() => this.changed())
  }

  private readonly stopListening: () => void

  /** A window closed (View › New Window's copies): stops following the workbook. */
  dispose() {
    this.stopListening()
    this.listeners.clear()
    this.scrollListeners.clear()
  }

  get sheet(): Sheet {
    return this.wb.sheets[this.sheetIndex]
  }

  /**
   * Records a change for undo, with the sheet it happened on: undo and redo show that sheet again. `sheetAfter` is
   * the sheet shown after the change, when the change itself moves to another one (adding or deleting a sheet).
   */
  transact(
    ops: Op[],
    o: { label: string; before: Selection; after?: Selection; sheetAfter?: number; perSheet?: true },
  ) {
    const at = this.sheetIndex
    // With grouped sheets the edit repeats on each (`perSheet`: the ops were already worked out for each sheet).
    const grouped = o.perSheet ? ops : this.group.expand(ops)
    if (!grouped) {
      this.shell.alert(GROUP_MESSAGE)
      return null
    }
    ops = grouped
    if (this.refuses(ops)) return null
    const entry = this.wb.transact(ops, {
      label: o.label,
      before: { ...o.before, sheet: at } as ViewMeta,
      // Without `after`, redo shows the sheet with the selection it had.
      after: (o.after ? { ...o.after, sheet: o.sheetAfter ?? at } : { sheet: o.sheetAfter ?? at }) as ViewMeta,
    })
    if (entry) this.fitRows(entry, ops)
    return entry
  }

  /**
   * Excel heightens (and lowers again) rows whose content needs it: wrapped text, line breaks, bigger fonts. Rows
   * the user sized by hand keep their height. The fitted heights join the edit's undo step. Very large edits (more
   * than FIT_LIMIT rows) are left as they are, to keep formatting whole columns fast.
   */
  private fitRows(entry: UndoEntry<ViewMeta>, ops: Op[]) {
    const fit = this.rowFit
    if (!fit) return
    const rows = new Set<number>()
    const sheetIndex = this.sheetIndex
    const used = this.sheet.cells.usedBounds()
    const addRange = (r: Range) => {
      const r0 = Math.max(r.r0, used ? 0 : r.r1 + 1)
      const r1 = Math.min(r.r1, used ? used.r1 : -1)
      for (let row = r0; row <= r1 && rows.size <= FIT_LIMIT; row++) rows.add(row)
    }
    for (const op of ops) {
      if (!("sheet" in op) || op.sheet !== sheetIndex) continue
      if (op.kind === "setCells") for (const w of op.cells) rows.add(w.row)
      else if (op.kind === "format" && Object.keys(op.patch).some((k) => FIT_KEYS.has(k))) op.ranges.forEach(addRange)
      if (rows.size > FIT_LIMIT) return
    }
    if (!rows.size) return
    const axis = this.sheet.rows
    const entries: AxisEntry[] = []
    for (const row of rows) {
      if (axis.hasCustomSize(row) && !axis.isAuto(row)) continue
      const size = fit(row)
      const current = axis.hasCustomSize(row) ? axis.storedSize(row) : null
      if (size === current) continue
      entries.push({ index: row, size, hidden: axis.isHidden(row), ...(size !== null ? { auto: true as const } : {}) })
    }
    if (entries.length) this.wb.amend(entry, [{ kind: "axis", sheet: sheetIndex, axis: "rows", entries }])
  }

  /** A protected sheet refuses the change: Excel's message, and nothing happens. */
  refuses(ops: Op[]): boolean {
    if (this.wb.structure.protection && ops.some(structureBlocks)) {
      this.shell.alert(SHEET_MESSAGES.structureLocked)
      return true
    }
    // The sheet each op works on: one added earlier in the same edit (Forecast Sheet, Show Details) is new and empty.
    const now: (Sheet | undefined)[] = [...this.wb.sheets]
    const targets = ops.map((op) => {
      if (op.kind === "addSheet") {
        now.splice(op.sheet, 0, undefined)
        return undefined
      }
      return "sheet" in op ? now[op.sheet] : undefined
    })
    if (ops.some((op, i) => targets[i] !== undefined && breaksDataTable(targets[i]!, op))) {
      this.shell.alert(DATA_TABLE_MESSAGE)
      return true
    }
    // A PivotTable's own changes (a "pivots" op rewrites its cells) are fine.
    const pivotEdit = ops.some((op) => op.kind === "pivots")
    if (!pivotEdit && ops.some((op, i) => targets[i] !== undefined && breaksPivot(targets[i]!, op))) {
      this.shell.alert(PIVOT_MESSAGE)
      return true
    }
    const blocked = ops.some((op, i) => {
      const sheet = targets[i]
      return sheet !== undefined && sheet.protection !== null && protectionBlocks(sheet, this.wb.styles, op)
    })
    if (blocked) {
      // Cells of an edit range with a password: ask for it (the change is made again once it's open).
      const shut = ops.flatMap((op) =>
        op.kind === "setCells" && op.sheet === this.sheetIndex
          ? op.cells.flatMap((w) => this.editRanges.shutRangeAt(w.row, w.col) ?? [])
          : [],
      )[0]
      if (shut && this.shell.unlockRange) this.shell.unlockRange(shut.name, () => undefined)
      else this.shell.alert(PROTECTED_MESSAGE)
    }
    return blocked
  }

  /** Whether the active cell can be edited (not a locked cell on a protected sheet). */
  canEditActive(): boolean {
    const { row, col } = this.selection.active
    return canEditCell(this.sheet, this.wb.styles, row, col)
  }

  /** Shows another sheet; the one left keeps its selection and scroll for when it comes back. */
  activate(index: number) {
    const sheets = this.wb.sheets
    const target = sheets[index]
    if (!target) return
    if (this.edit) this.commitEdit()
    if (this.sheet !== target) this.headerFooter.commit()
    this.annotations.editing = null
    this.annotations.open = null
    // A sheet outside the group ungroups (Excel); one in it becomes the group's active sheet.
    if (this.group.keeps(index)) this.group.keepWhileActivating(index)
    else if (index !== this.sheetIndex || target !== this.shown) this.group.clear(false)
    this.sheetIndex = index
    if (this.shown === target) {
      this.changed()
      return
    }
    if (sheets.includes(this.shown))
      this.views.set(this.shown, {
        selection: this.selection,
        scrollX: this.scrollX,
        scrollY: this.scrollY,
        freeze: this.freeze,
        split: this.split,
        zoom: this.zoom,
        view: this.view,
        viewZooms: this.viewZooms,
        showGridlines: this.showGridlines,
        showHeadings: this.showHeadings,
        showRuler: this.showRuler,
      })
    this.shown = target
    const view = this.views.get(target)
    this.selection = view?.selection ?? selectCell({ row: 0, col: 0 })
    this.scrollX = view?.scrollX ?? 0
    this.scrollY = view?.scrollY ?? 0
    this.freeze = view?.freeze ?? null
    this.split = view?.split ?? null
    this.zoom = view?.zoom ?? 1
    this.view = view?.view ?? "normal"
    this.viewZooms = view?.viewZooms ?? {}
    this.showGridlines = view?.showGridlines ?? true
    this.showHeadings = view?.showHeadings ?? true
    this.showRuler = view?.showRuler ?? true
    this.lastCtrlARegion = null
    this.fillPreview = null
    this.endMode = false
    this.changed()
    this.keepSelectable()
  }

  // ── Views as files keep them ─────────────────────────────────────────────────────────────────────────────

  /** Takes each sheet's view from an opened file (selection, scroll, frozen panes, zoom, gridlines) and shows `active`. */
  loadViews(views: readonly FileSheetView[], active: number) {
    this.wb.sheets.forEach((sheet, i) => {
      const v = views[i]
      if (!v) return
      const ranges = v.selection.length
        ? v.selection
        : [{ r0: v.active.row, c0: v.active.col, r1: v.active.row, c1: v.active.col }]
      const activeRange = Math.max(
        0,
        ranges.findIndex(
          (r) => v.active.row >= r.r0 && v.active.row <= r.r1 && v.active.col >= r.c0 && v.active.col <= r.c1,
        ),
      )
      const m = this.metricsFor(sheet, v.mode ?? "normal")
      this.views.set(sheet, {
        selection: { ranges, active: v.active, activeRange, anchor: v.active, focus: v.active },
        scrollX: m.cols.pos(v.topLeft.col),
        scrollY: m.rows.pos(v.topLeft.row),
        freeze: v.freeze,
        split: v.split
          ? {
              width: v.split.width,
              height: v.split.height,
              scrollX: sheet.cols.offset(v.split.topLeft.col),
              scrollY: sheet.rows.offset(v.split.topLeft.row),
            }
          : null,
        zoom: Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, v.zoom)),
        view: v.mode ?? "normal",
        viewZooms: { ...(v.zooms ?? {}) },
        showGridlines: v.showGridlines,
        showHeadings: v.showHeadings,
        showRuler: v.showRuler ?? true,
      })
    })
    // Make the target load its view: activate() reads the stored view only when the shown sheet changes.
    this.shown = null as unknown as Sheet
    this.activate(Math.min(Math.max(0, active), this.wb.sheets.length - 1))
    // Sheets saved as selected besides the active one open grouped (Excel keeps the group in the file).
    const selected = views.flatMap((v, i) => (v?.selected ? [i] : []))
    if (selected.length > 1 && selected.includes(this.sheetIndex)) this.group.set(selected)
  }

  /** Each sheet's view, for saving. */
  saveViews(): FileSheetView[] {
    return this.wb.sheets.map((sheet, i) => {
      const live = sheet === this.shown
      const v = live
        ? {
            selection: this.selection,
            scrollX: this.scrollX,
            scrollY: this.scrollY,
            freeze: this.freeze,
            split: this.split,
            zoom: this.zoom,
            view: this.view,
            viewZooms: this.viewZooms,
            showGridlines: this.showGridlines,
            showHeadings: this.showHeadings,
            showRuler: this.showRuler,
          }
        : this.views.get(sheet)
      if (!v) return { ...DEFAULT_FILE_VIEW, selected: this.group.has(i) }
      const m = this.metricsFor(sheet, v.view)
      return {
        freeze: v.freeze,
        split: v.split
          ? {
              width: v.split.width,
              height: v.split.height,
              topLeft: { row: sheet.rows.indexAt(v.split.scrollY), col: sheet.cols.indexAt(v.split.scrollX) },
            }
          : null,
        zoom: v.zoom,
        mode: v.view,
        zooms: { ...v.viewZooms, [v.view]: v.zoom },
        showGridlines: v.showGridlines,
        showHeadings: v.showHeadings,
        ...(v.showRuler ? {} : { showRuler: false }),
        topLeft: { row: m.rows.indexAt(v.scrollY), col: m.cols.indexAt(v.scrollX) },
        active: v.selection.active,
        selection: [...v.selection.ranges],
        selected: this.group.has(i),
      }
    })
  }

  /** What conditional formatting does to a cell of the sheet shown (null: nothing). */
  cfAt(row: number, col: number): CFResult | null {
    const preview = this.conditional.preview
    if (!this.sheet.conditional.length && !preview) return null
    const day = todaySerial()
    let cf = this.cf
    if (!cf || cf.sheet !== this.sheet || cf.version !== this.wb.version || cf.day !== day || cf.preview !== preview) {
      const entries = preview ? [preview, ...this.sheet.conditional] : undefined
      cf = this.cf = {
        sheet: this.sheet,
        version: this.wb.version,
        day,
        preview,
        ev: new ConditionalEval(this.wb.sheets, this.sheetIndex, day, entries),
      }
    }
    return cf.ev.at(row, col)
  }

  get mode(): CellMode {
    if (!this.edit) return "ready"
    return this.edit.point ? "point" : this.edit.mode
  }

  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  getVersion = () => this.version

  private scrollListeners = new Set<() => void>()
  private scrollVersion = 0

  /**
   * Scrolling alone: the grid listens here as well as to `subscribe`. Kept apart so a wheel or a scrollbar drag
   * doesn't re-render the ribbon, formula bar, tabs and panes on every step (slow in WebKitGTK and on phones).
   */
  subscribeScroll = (listener: () => void) => {
    this.scrollListeners.add(listener)
    return () => {
      this.scrollListeners.delete(listener)
    }
  }

  getScrollVersion = () => this.scrollVersion

  changed() {
    if (this.validation?.circles) this.validation.refreshCircles()
    this.version++
    for (const listener of this.listeners) listener()
  }

  // ── Selection ──────────────────────────────────────────────────────────────────────────────────────────────

  setSelection(selection: Selection, reveal: CellAddress | null = selection.focus) {
    // Selecting cells leaves the selected objects (and a pending shape).
    this.objects.selected = []
    this.objects.editing = null
    this.objects.pending = null
    const next = this.mergeAware(selection)
    // Protect Sheet with "Select locked cells" off: the cursor stays off locked cells (Excel).
    if (!canSelect(this.sheet, this.wb.styles, next.ranges)) return
    this.selection = next
    this.lastCtrlARegion = null
    if (reveal) this.scrollIntoView(reveal)
    this.changed()
  }

  /** Merged cells are selected whole, and the active cell of a merge is its top-left cell. */
  private mergeAware(selection: Selection): Selection {
    if (!this.sheet.merges.size) return selection
    const ranges = selection.ranges.map((r) => growToMerges(this.sheet, r))
    const active = mergeAnchor(this.sheet, selection.active)
    if (ranges.every((r, i) => r === selection.ranges[i]) && active === selection.active) return selection
    return { ...selection, ranges, active }
  }

  /** The merged area the active cell belongs to, or the cell itself (the active outline and the editor cover it). */
  activeArea(): Range {
    const { row, col } = this.selection.active
    return this.sheet.merges.at(row, col) ?? { r0: row, r1: row, c0: col, c1: col }
  }

  selectCellAt(address: CellAddress) {
    this.tabStartCol = null
    this.setSelection(selectCell(clampCell(address)))
  }

  /** The Name Box text: the active cell, or "3R x 2C" while a drag is in progress. */
  nameBoxText(dragging = false): string {
    if (dragging) {
      const range = this.selection.ranges[this.selection.activeRange]
      if (!sameRange(range, rangeOf(this.selection.active))) {
        if (isWholeColumns(range) && isWholeRows(range)) return `${MAX_ROWS}R x ${MAX_COLS}C`
        if (isWholeColumns(range)) return `${rangeCols(range)}C`
        if (isWholeRows(range)) return `${rangeRows(range)}R`
        return `${rangeRows(range)}R x ${rangeCols(range)}C`
      }
    }
    return this.names.nameOfSelection() ?? cellName(this.selection.active)
  }

  /** The formula bar's text: what is being typed, or the active cell's input. */
  formulaBarText(): string {
    if (this.edit) return this.edit.text
    const { row, col } = this.selection.active
    const anchor = this.spillAnchorOf(row, col)
    if (anchor) return this.sheet.cells.input(anchor.row, anchor.col) ?? ""
    const input = this.sheet.cells.input(row, col)
    // A data table cell shows its TABLE formula in braces, as Excel does.
    if (isDataTableInput(input)) return `{${input}}`
    return (
      this.sheet.cells.input(row, col) ??
      editTextFor(this.sheet.cells.value(row, col), this.wb.styles.get(this.sheet.styleAt(row, col)).numberFormat)
    )
  }

  /** The active cell holds a spilled value: the formula bar shows its formula greyed out, as Excel does. */
  get formulaBarGhost(): boolean {
    if (this.edit) return false
    const { row, col } = this.selection.active
    return this.spillAnchorOf(row, col) !== null
  }

  /** The spill (anchor and area) covering a cell of this sheet, from the calculation status. */
  spillAt(row: number, col: number) {
    for (const spill of this.wb.calcStatus.spills) {
      const r = spill.range
      if (spill.sheet === this.sheetIndex && row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1) return spill
    }
    return null
  }

  private spillAnchorOf(row: number, col: number): CellAddress | null {
    if (!this.sheet.cells.isSpill(row, col)) return null
    const spill = this.spillAt(row, col)
    return spill && (spill.row !== row || spill.col !== col) ? { row: spill.row, col: spill.col } : null
  }

  // ── Editing ────────────────────────────────────────────────────────────────────────────────────────────────

  /** Starts an entry: typing replaces (enter mode) or F2/double-click edits the content (edit mode). */
  beginEdit(mode: "enter" | "edit", text?: string, surface: "cell" | "bar" = "cell") {
    if (!this.canEditActive()) {
      // A cell in an edit range with a password: Unlock Range, then the entry starts as asked.
      const { row, col } = this.selection.active
      const shut = this.editRanges.shutRangeAt(row, col)
      if (shut && this.shell.unlockRange) this.shell.unlockRange(shut.name, () => this.beginEdit(mode, text, surface))
      else this.shell.alert(PROTECTED_MESSAGE)
      return
    }
    const { row, col } = this.selection.active
    if (isDataTableInput(this.sheet.cells.input(row, col))) {
      this.shell.alert(DATA_TABLE_MESSAGE)
      return
    }
    if (pivotAt(this.sheet, row, col)) {
      this.shell.alert(PIVOT_MESSAGE)
      return
    }
    const initial = text ?? (mode === "edit" ? this.formulaBarText() : "")
    this.edit = {
      row,
      col,
      mode,
      surface,
      text: initial,
      selStart: initial.length,
      selEnd: initial.length,
      point: null,
      pick: 0,
    }
    this.endMode = false
    this.scrollIntoView(this.selection.active)
    this.changed()
  }

  /** Typing changed the text: any reference being pointed at is finished. */
  setEditText(text: string, selStart = text.length, selEnd = selStart) {
    if (!this.edit) return
    this.edit = { ...this.edit, text, selStart, selEnd, point: null, pick: 0 }
    this.changed()
  }

  /** The caret moved (click, Home/End, arrows in Edit mode). */
  setEditSelection(selStart: number, selEnd = selStart) {
    const edit = this.edit
    if (!edit || (edit.selStart === selStart && edit.selEnd === selEnd)) return
    const keepPoint = edit.point && selEnd === edit.point.end && selStart === selEnd
    this.edit = { ...edit, selStart, selEnd, point: keepPoint ? edit.point : null, pick: 0 }
    this.changed()
  }

  /** The references in the formula being typed (on this sheet), with their colours, for the editor and the grid. */
  editReferences(): (EditReference & { pointing: boolean })[] {
    const edit = this.edit
    if (!edit || !edit.text.startsWith("=")) return []
    const name = this.sheet.name.toLowerCase()
    return editReferences(edit.text)
      .filter((ref) => ref.sheet === null || ref.sheet.toLowerCase() === name)
      .map((ref) => ({ ...ref, pointing: edit.point !== null && ref.start === edit.point.start }))
  }

  /** Whether a click on a cell (or an arrow key in Enter mode) writes a reference into the formula. */
  get canPoint(): boolean {
    const edit = this.edit
    if (!edit || !edit.text.startsWith("=")) return false
    return edit.point !== null || (edit.selStart === edit.selEnd && canPointAt(edit.text, edit.selEnd))
  }

  /** Point mode: writes (or rewrites) the reference to anchor…focus at the caret. */
  pointAt(anchor: CellAddress, focus: CellAddress = anchor) {
    const edit = this.edit
    if (!edit) return
    const range = {
      r0: Math.min(anchor.row, focus.row),
      r1: Math.max(anchor.row, focus.row),
      c0: Math.min(anchor.col, focus.col),
      c1: Math.max(anchor.col, focus.col),
    }
    // A single value cell of a PivotTable: GETPIVOTDATA (Generate GetPivotData), as Excel writes.
    const single = range.r0 === range.r1 && range.c0 === range.c1
    const ref = (single && this.pivots.pointedFormula(range.r0, range.c0)) || pointedText(range)
    const start = edit.point ? edit.point.start : edit.selStart
    const end = edit.point ? edit.point.end : edit.selEnd
    const text = edit.text.slice(0, start) + ref + edit.text.slice(end)
    const caret = start + ref.length
    this.edit = {
      ...edit,
      text,
      selStart: caret,
      selEnd: caret,
      point: { start, end: caret, anchor, focus },
      pick: 0,
    }
    this.scrollIntoView(focus)
    this.changed()
  }

  /** Arrow keys in Point mode move (or with Shift extend) the pointed reference. */
  pointArrow(dir: Direction, ctrl: boolean, shift: boolean) {
    const edit = this.edit
    if (!edit) return
    const from = edit.point?.focus ?? { row: edit.row, col: edit.col }
    const next = ctrl ? dataEdge(this.sheet, from, dir) : step(this.sheet, from, dir)
    if (shift && edit.point) this.pointAt(edit.point.anchor, next)
    else this.pointAt(next)
  }

  /** F4: cycles the $ signs of the reference at the caret and selects it. */
  toggleAbsoluteReference() {
    const edit = this.edit
    if (!edit) return
    const result = toggleAbsolute(edit.text, edit.selEnd)
    if (!result) return
    this.edit = { ...edit, text: result.text, selStart: result.start, selEnd: result.end, point: null }
    this.changed()
  }

  /** The function list for the name being typed, if any. */
  get autocomplete(): Autocomplete | null {
    const edit = this.edit
    if (!edit || !edit.text.startsWith("=") || edit.selStart !== edit.selEnd) return null
    const name = nameAtCaret(edit.text, edit.selEnd)
    if (!name) return null
    const items = functionsStartingWith(name.prefix)
    if (!items.length) return null
    return { ...name, items, pick: Math.min(edit.pick, items.length - 1) }
  }

  /** Tab (or a click) in the function list: writes the name and an opening parenthesis. */
  acceptAutocomplete(index?: number) {
    const edit = this.edit
    const list = this.autocomplete
    if (!edit || !list) return false
    const info = list.items[index ?? list.pick]
    const before = edit.text.slice(0, list.start)
    const after = edit.text.slice(edit.selEnd)
    const insert = `${info.name}(`
    const caret = before.length + insert.length
    this.edit = { ...edit, text: before + insert + after, selStart: caret, selEnd: caret, point: null, pick: 0 }
    this.changed()
    return true
  }

  /**
   * Insert Function (fx, Shift+F3) and the Function Library menus: `NAME()` at the caret with the caret inside the
   * parentheses; on a cell that isn't being edited the entry starts as `=NAME()`. Excel opens Function Arguments
   * next; here the argument tooltip guides the typing.
   */
  insertFunction(name: string) {
    const call = `${name}(`
    if (!this.edit) {
      this.beginEdit("enter", `=${call})`)
      const edit = this.edit as EditState | null
      if (edit) {
        const caret = call.length + 1
        this.edit = { ...edit, selStart: caret, selEnd: caret }
        this.changed()
      }
      return
    }
    const edit = this.edit
    const before = edit.text.slice(0, edit.selStart)
    const after = edit.text.slice(edit.selEnd)
    const lead = edit.text.startsWith("=") ? "" : before.length === 0 ? "=" : ""
    const insert = `${lead}${call})`
    const caret = before.length + lead.length + call.length
    this.edit = { ...edit, text: before + insert + after, selStart: caret, selEnd: caret, point: null, pick: 0 }
    this.changed()
  }

  /** Insert › Symbol: text at the caret; with no entry open, the active cell opens for editing with it at the end. */
  insertText(text: string) {
    if (!this.edit) this.beginEdit("edit")
    const edit = this.edit
    if (!edit) return
    const before = edit.text.slice(0, edit.selStart)
    const after = edit.text.slice(edit.selEnd)
    const caret = before.length + text.length
    this.edit = { ...edit, text: before + text + after, selStart: caret, selEnd: caret, point: null, pick: 0 }
    this.changed()
  }

  /** Use in Formula: a defined name at the caret (an entry starts with "=" when none is open). */
  insertName(name: string) {
    if (!this.edit) {
      this.beginEdit("enter", `=${name}`)
      return
    }
    const edit = this.edit
    const before = edit.text.slice(0, edit.selStart)
    const after = edit.text.slice(edit.selEnd)
    const insert = (before.length === 0 && !edit.text.startsWith("=") ? "=" : "") + name
    const caret = before.length + insert.length
    this.edit = { ...edit, text: before + insert + after, selStart: caret, selEnd: caret, point: null, pick: 0 }
    this.changed()
  }

  private movePick(delta: number) {
    const list = this.autocomplete
    if (!this.edit || !list) return
    const pick = (list.pick + delta + list.items.length) % list.items.length
    this.edit = { ...this.edit, pick }
    this.changed()
  }

  setEditSurface(surface: "cell" | "bar") {
    if (this.edit && this.edit.surface !== surface) {
      this.edit = { ...this.edit, surface }
      this.changed()
    }
  }

  toggleEditMode() {
    if (!this.edit) return
    this.edit = { ...this.edit, mode: this.edit.mode === "enter" ? "edit" : "enter" }
    this.changed()
  }

  cancelEdit() {
    if (!this.edit) return
    this.edit = null
    this.changed()
  }

  /**
   * Commits the entry. `fill` (Ctrl+Enter) writes it to every selected cell. Returns false when nothing was being
   * edited. A new entry clears the marching ants, as in Excel.
   */
  commitEdit(fill = false, extra?: CommitExtra): boolean {
    const edit = this.edit
    if (!edit) return false
    this.edit = null
    const before = this.selection
    let op: Op
    if (fill && !isSingleCellOn(this.sheet, this.selection)) {
      op = fillInputOp(
        this.sheetIndex,
        this.sheet,
        this.wb.styles,
        this.selection.ranges,
        edit.text,
        this.selection.active,
      )
    } else {
      op = writeInputOp(this.sheetIndex, this.sheet, this.wb.styles, edit.row, edit.col, edit.text)
    }
    const label = edit.text ? `Typing '${truncate(edit.text)}' in ${cellName(edit)}` : `Clear ${cellName(edit)}`
    this.clip = null
    // A new header renames its column, and formulas that name the column follow (Excel).
    const ops: Op[] = [...this.tables.headerRenameOps(op), op]
    // A line break typed with Alt+Enter turns on Wrap Text for the cell (Excel), so the lines show.
    if (!fill || isSingleCellOn(this.sheet, this.selection))
      ops.push(...this.tables.entryOps(edit.row, edit.col, edit.text))
    if (edit.text.includes("\n") && !edit.text.startsWith("=")) {
      const ranges = fill && !isSingleCellOn(this.sheet, this.selection) ? this.selection.ranges : [rangeOf(edit)]
      if (ranges.some((r) => !this.wb.styles.get(this.sheet.styleAt(r.r0, r.c0)).wrap))
        ops.push({
          kind: "format",
          sheet: this.sheetIndex,
          ranges: ranges.map((r) => ({ ...r })),
          patch: { wrap: true },
        })
    }
    if (!fill || isSingleCellOn(this.sheet, this.selection))
      ops.push(...this.hyperlinks.autoLinkOps(edit.row, edit.col, edit.text))
    if (extra) ops.push(...extra.ops)
    this.transact(ops, { label: extra?.label ?? label, before, after: before })
    return true
  }

  /**
   * Commits the entry once Data Validation lets it in; `then` runs after (a move). An entry that breaks its cell's
   * rule shows the error alert and stays in the cell: Retry (or No) selects it to type again, Cancel throws it away,
   * Yes (Warning) and OK (Information) enter it. Returns whether it went in now.
   */
  commitChecked(fill = false, then?: () => void, extra?: CommitExtra): boolean {
    const edit = this.edit
    if (!edit) return false
    const bad = this.validation.rejects(edit.text, edit.row, edit.col)
    if (!bad || !this.shell.validationAlert) {
      this.commitEdit(fill, extra)
      then?.()
      return true
    }
    this.shell.validationAlert(this.validation.alertFor(bad), (answer) => {
      if (!this.edit) return
      if (answer === "yes" || answer === "ok") {
        this.commitEdit(fill, extra)
        then?.()
      } else if (answer === "cancel") this.cancelEdit()
      else this.setEditSelection(0, this.edit.text.length)
    })
    return false
  }

  // ── Movement ───────────────────────────────────────────────────────────────────────────────────────────────

  /** Arrow keys (with Ctrl for data edges, Shift to extend). */
  arrow(dir: Direction, ctrl: boolean, shift: boolean) {
    this.tabStartCol = null
    const jump = ctrl || this.endMode
    this.endMode = false
    if (shift) {
      const focus = jump ? dataEdge(this.sheet, this.selection.focus, dir) : step(this.sheet, this.selection.focus, dir)
      this.setSelection(extendSelection(this.selection, keepWholeLines(this.selection, focus)), focus)
    } else if (selectableCells(this.sheet) === "unlocked") {
      // Only unlocked cells can be selected: the arrow goes to the nearest one that way.
      const target = nextSelectable(this.sheet, this.wb.styles, this.selection.active, dir, "line")
      if (target) this.setSelection(selectCell(target))
    } else {
      const target = jump
        ? dataEdge(this.sheet, this.selection.active, dir)
        : step(this.sheet, this.selection.active, dir)
      this.setSelection(selectCell(target))
    }
  }

  /**
   * A protected sheet that lets only unlocked cells be selected: a cursor on a locked cell goes to the first unlocked
   * one (as Excel does when the sheet is protected or shown).
   */
  keepSelectable() {
    if (selectableCells(this.sheet) !== "unlocked" || canSelect(this.sheet, this.wb.styles, this.selection.ranges))
      return
    const first = nextSelectable(this.sheet, this.wb.styles, { row: -1, col: -1 }, "right", "order")
    if (first) this.setSelection(selectCell(first))
  }

  /** Enter / Shift+Enter / Tab / Shift+Tab: move within a multi-cell selection, or step (with Excel's Tab return). */
  enterMove(dir: Direction) {
    const single = isSingleCellOn(this.sheet, this.selection)
    if (single && selectableCells(this.sheet) === "unlocked") {
      // Tab and Enter go round the unlocked cells, like the fields of a form (Excel).
      this.tabStartCol = null
      const target = nextSelectable(this.sheet, this.wb.styles, this.selection.active, dir, "order")
      if (target) this.setSelection(selectCell(target))
      return
    }
    if (single && dir === "down" && this.tabStartCol !== null) {
      const row = step(this.sheet, this.selection.active, "down").row
      const target = { row, col: this.tabStartCol }
      this.tabStartCol = null
      this.setSelection(selectCell(target))
      return
    }
    if (single && (dir === "right" || dir === "left")) {
      // Tab in a table wraps to the next row's first column; on the last cell it adds a row (Excel).
      const wrap = this.tables.tabTarget(dir)
      if (wrap) {
        this.tabStartCol = null
        this.setSelection(selectCell(wrap))
        return
      }
      this.tabStartCol ??= this.selection.active.col
    } else if (single) {
      this.tabStartCol = null
    }
    const next = moveWithinSelection(this.sheet, this.selection, dir)
    this.selection = next
    this.lastCtrlARegion = null
    this.scrollIntoView(next.active)
    this.changed()
  }

  home(ctrl: boolean, shift: boolean) {
    this.tabStartCol = null
    this.endMode = false
    const from = shift ? this.selection.focus : this.selection.active
    // With frozen panes Home and Ctrl+Home go to the first unfrozen column and cell (Excel).
    const f = this.freeze
    const col = f && f.col > f.left ? f.col : 0
    const row = f && f.row > f.top ? f.row : 0
    const target = ctrl ? { row, col } : { row: from.row, col }
    if (shift) this.setSelection(extendSelection(this.selection, target), target)
    else this.setSelection(selectCell(target))
  }

  ctrlEnd(shift: boolean) {
    this.tabStartCol = null
    this.endMode = false
    const target = lastCell(this.sheet)
    if (shift) this.setSelection(extendSelection(this.selection, target), target)
    else this.setSelection(selectCell(target))
  }

  /** Page Down/Up (Alt: a screen right/left): moves the active cell and the view by one screen. */
  page(dir: 1 | -1, horizontal: boolean, shift: boolean) {
    this.tabStartCol = null
    this.endMode = false
    const axis = horizontal ? this.sheet.cols : this.sheet.rows
    const span = horizontal ? this.paneWidth : this.paneHeight
    const from = shift ? this.selection.focus : this.selection.active
    const index = horizontal ? from.col : from.row
    const targetPx = axis.offset(index) + dir * span
    const targetIndex = Math.max(0, Math.min(axis.count - 1, axis.indexAt(Math.max(0, targetPx))))
    const target = horizontal ? { row: from.row, col: targetIndex } : { row: targetIndex, col: from.col }
    if (horizontal) this.scrollX = Math.max(this.minScrollX, this.scrollX + dir * span)
    else this.scrollY = Math.max(this.minScrollY, this.scrollY + dir * span)
    if (shift) this.setSelection(extendSelection(this.selection, target), target)
    else this.setSelection(selectCell(target), target)
  }

  toggleEndMode() {
    this.endMode = !this.endMode
    this.changed()
  }

  /** Ctrl+A: the current region first (when in data), then the whole sheet. */
  selectAllOrRegion() {
    const region = currentRegion(this.sheet, this.selection.active)
    const isData = region.r0 !== region.r1 || region.c0 !== region.c1 || this.sheet.cells.hasValue(region.r0, region.c0)
    const current = this.selection.ranges[this.selection.ranges.length - 1]
    if (isData && !(this.lastCtrlARegion && sameRange(current, this.lastCtrlARegion))) {
      const selection = selectRange(region, this.selection.active)
      this.setSelection(selection, null)
      this.lastCtrlARegion = region
      return
    }
    this.setSelection(selectAll(this.selection.active), null)
  }

  /** Ctrl+Shift+* : the current region. */
  selectCurrentRegion() {
    this.setSelection(selectRange(currentRegion(this.sheet, this.selection.active), this.selection.active), null)
  }

  selectColumns() {
    this.setSelection(expandToColumns(this.selection), null)
  }

  selectRowsOfSelection() {
    this.setSelection(expandToRows(this.selection), null)
  }

  // ── Mouse selection (called by the view) ──────────────────────────────────────────────────────────────────

  pointerCell(address: CellAddress, options: { shift: boolean; ctrl: boolean }) {
    this.tabStartCol = null
    this.endMode = false
    if (options.shift) {
      this.setSelection(extendSelection(this.selection, address), address)
    } else if (options.ctrl) {
      this.setSelection(addRange(this.selection, address), address)
    } else {
      this.setSelection(selectCell(address), address)
    }
  }

  /** Dragging with the button down: the last range follows the pointer. */
  dragTo(address: CellAddress) {
    const focus = keepWholeLines(this.selection, clampCell(address))
    if (focus.row === this.selection.focus.row && focus.col === this.selection.focus.col) return
    this.setSelection(extendSelection(this.selection, focus), address)
  }

  pointerRowHeader(row: number, options: { shift: boolean; ctrl: boolean }) {
    const leftCol = this.colMetric.indexAt(this.scrollX)
    if (options.shift) {
      const anchor = this.selection.anchor.row
      this.setSelection(
        extendSelection({ ...this.selection, anchor: { row: anchor, col: 0 } }, { row, col: MAX_COLS - 1 }),
        null,
      )
    } else if (options.ctrl) {
      this.setSelection(
        addRange(this.selection, { row, col: leftCol }, { r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }),
        null,
      )
    } else {
      this.setSelection(selectRows(row, row, leftCol), null)
    }
  }

  pointerColHeader(col: number, options: { shift: boolean; ctrl: boolean }) {
    const topRow = this.rowMetric.indexAt(this.scrollY)
    if (options.shift) {
      const anchor = this.selection.anchor.col
      this.setSelection(
        extendSelection({ ...this.selection, anchor: { row: 0, col: anchor } }, { row: MAX_ROWS - 1, col }),
        null,
      )
    } else if (options.ctrl) {
      this.setSelection(
        addRange(this.selection, { row: topRow, col }, { r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col }),
        null,
      )
    } else {
      this.setSelection(selectCols(col, col, topRow), null)
    }
  }

  selectEverything() {
    this.setSelection(selectAll(this.topLeft), null)
  }

  // ── Content commands ───────────────────────────────────────────────────────────────────────────────────────

  /** Delete: clears contents, keeps formats. */
  clearContents() {
    const ranges = this.selection.ranges
    if (!this.group.active && this.pivots.deleteCovered(ranges)) return
    const ops = this.group.each((index, sheet) => clearContentsOp(index, sheet, ranges))
    if (ops.length)
      this.transact(ops, { label: "Clear", before: this.selection, after: this.selection, perSheet: true })
  }

  undo() {
    if (this.edit) return
    const entry = this.wb.undo()
    if (entry) this.restoreView(entry.before)
  }

  redo() {
    if (this.edit) return
    const entry = this.wb.redo()
    if (entry) this.restoreView(entry.after)
  }

  /** After undo/redo: back to the sheet and selection of that moment (or the nearest sheet still there). */
  private restoreView(meta: ViewMeta | undefined) {
    const sheets = this.wb.sheets
    let index = meta?.sheet ?? sheets.indexOf(this.shown)
    if (index < 0) index = this.sheetIndex
    index = Math.min(index, sheets.length - 1)
    if (sheets[index].hidden) index = this.sheets.nearestVisible(index)
    this.activate(index)
    if (!meta || !("ranges" in meta)) return
    const { sheet: _sheet, ...selection } = meta
    this.restoreSelection(selection)
  }

  private restoreSelection(selection: Selection) {
    this.selection = selection
    this.clip = null
    this.scrollIntoView(selection.active)
    this.changed()
  }

  /** Ctrl+D / Ctrl+R: fill down/right from the first row/column of the selection. */
  fillFromEdge(dir: "down" | "right") {
    const ops: Op[] = []
    for (const range of this.selection.ranges) {
      const source = dir === "down" ? { ...range, r1: range.r0 } : { ...range, c1: range.c0 }
      // With one row (or column) selected, Excel fills from the one above (or to the left).
      if ((dir === "down" && range.r0 === range.r1) || (dir === "right" && range.c0 === range.c1)) {
        const from =
          dir === "down"
            ? { ...range, r0: range.r0 - 1, r1: range.r0 - 1 }
            : { ...range, c0: range.c0 - 1, c1: range.c0 - 1 }
        if (from.r0 < 0 || from.c0 < 0) continue
        ops.push(pasteBlockOp(this.sheetIndex, range, readBlock(this.sheet, from)))
        continue
      }
      ops.push(pasteBlockOp(this.sheetIndex, range, readBlock(this.sheet, source)))
    }
    if (ops.length) {
      this.transact(ops, {
        label: dir === "down" ? "Fill Down" : "Fill Right",
        before: this.selection,
        after: this.selection,
      })
    }
  }

  /** The fill handle was dropped: `target` is the grown selection. */
  fillTo(target: Range, invert: boolean) {
    const source = this.selection.ranges[this.selection.ranges.length - 1]
    this.fillPreview = null
    if (sameRange(source, target)) {
      this.changed()
      return
    }
    const shrinking = target.r1 < source.r1 || target.c1 < source.c1
    if (shrinking && target.r0 === source.r0 && target.c0 === source.c0) {
      // Dragging the handle back inside the selection clears what it leaves (Excel's gray-out erase).
      const cleared: Range = target.r1 < source.r1 ? { ...source, r0: target.r1 + 1 } : { ...source, c0: target.c1 + 1 }
      const op = clearContentsOp(this.sheetIndex, this.sheet, [cleared])
      const after = selectRange(target, this.selection.active)
      if (op) this.transact([op], { label: "Clear", before: this.selection, after })
      this.setSelection(after, null)
      return
    }
    const op = fillOp(this.sheetIndex, this.sheet, source, target, invert)
    const after = selectRange(target, this.selection.active)
    if (op) this.transact([op], { label: "AutoFill", before: this.selection, after })
    this.setSelection(after, null)
  }

  // ── Clipboard ──────────────────────────────────────────────────────────────────────────────────────────────

  /** Copy/Cut: what goes on the system clipboard, and the marching ants. Multiple ranges cannot be copied. */
  /** The Office Clipboard (Home › Clipboard): what was copied here, newest first, up to 24 items. */
  clipboardHistory: { text: string; html: string; chart: boolean }[] = []
  /** Home › Clipboard: the Office Clipboard pane beside the sheet. */
  clipboardPaneOpen = false

  toggleClipboardPane(open = !this.clipboardPaneOpen) {
    this.clipboardPaneOpen = open
    this.changed()
  }

  /** Clear All, or one item, from the Office Clipboard. */
  forgetCopied(index?: number) {
    if (index === undefined) this.clipboardHistory = []
    else this.clipboardHistory = this.clipboardHistory.filter((_, i) => i !== index)
    this.changed()
  }

  copy(cut: boolean): { text: string; html: string } | null {
    const chart = this.charts.selected !== null
    const content = this.copyNow(cut)
    if (content) {
      const same = this.clipboardHistory[0]
      if (!same || same.text !== content.text || same.html !== content.html)
        this.clipboardHistory = [{ ...content, chart }, ...this.clipboardHistory].slice(0, 24)
    }
    return content
  }

  private copyNow(cut: boolean): { text: string; html: string } | null {
    if (this.charts.selected !== null) return this.charts.copy(cut)
    if (this.objects.selected.length) return this.objects.copy(cut, (list) => this.objectPicture?.(list) ?? null)
    if (this.selection.ranges.length !== 1) return null
    const range = clampToUsed(this.sheet, this.selection.ranges[0])
    const content = this.rangeContent(this.sheetIndex, range)
    // Notes and comments go with a copy pasted in Cells (not onto the system clipboard's text).
    const block = {
      ...content.block,
      notes: readNotes(this.sheet, range),
      // Links too: a copy takes each on its cells; a cut only those wholly inside (Excel 365).
      links: readHyperlinks(this.sheet.hyperlinks, range, cut),
    }
    this.clip = { range, cut, text: content.text, block, sheet: this.sheet }
    this.changed()
    return { text: content.text, html: content.html }
  }

  /**
   * A range as the clipboard carries it (and as Inkspire's linked tables show it): TSV of what each cell shows,
   * formatted HTML that says where it came from, and the values.
   */
  rangeContent(sheetIndex: number, range: Range) {
    const sheet = this.wb.sheets[sheetIndex]
    const block = readBlock(sheet, range)
    const rows = block.cells.map((row) =>
      Array.from({ length: block.cols }, (_, c) => {
        const cell = row[c]
        return formatValue(cell?.value, this.wb.styles.get(cell?.style ?? 0).numberFormat) ?? ""
      }),
    )
    const values = block.cells.map((row) =>
      Array.from({ length: block.cols }, (_, c) => {
        const v = row[c]?.value
        return v === undefined ? null : typeof v === "object" ? v.error : v
      }),
    )
    const text = encodeTSV(rows)
    const file = this.shell.documentPath?.() ?? null
    const link = file ? openLink(file, `${quoteSheet(sheet.name)}!${rangeName(range)}`) : null
    const source = { file, sheet: sheet.name, range, link }
    return { text, html: rangeHtml(sheet, this.wb.styles, range, rows, source), values, block }
  }

  clearClip() {
    if (this.clip) {
      this.clip = null
      this.changed()
    }
  }

  /**
   * Paste: our own copy (full content and formats) when the clipboard still holds the text we put there, else the
   * clipboard's text parsed as typed input. A cut moves its cells and ends the marching ants.
   */
  paste(text: string | null, html?: string | null): boolean {
    if (this.charts.paste(html)) return true
    if (this.objects.paste(html)) return true
    if (this.selection.ranges.length !== 1) return false
    const target = this.selection.ranges[0]
    const own = this.clip && text !== null && normalizeNewlines(text) === normalizeNewlines(this.clip.text)
    let block: CellBlock
    if (own && this.clip) {
      block = this.clip.block
    } else if (text !== null && text !== "") {
      block = blockFromText(this.wb.styles, decodeTSV(text))
    } else if (html) {
      block = blockFromText(this.wb.styles, htmlTableRows(html))
    } else {
      return false
    }
    const before = this.selection
    if (own && this.clip?.cut) {
      const area = { r0: target.r0, c0: target.c0, r1: target.r0 + block.rows - 1, c1: target.c0 + block.cols - 1 }
      const from = this.wb.sheets.indexOf(this.clip.sheet)
      if (from < 0) return false
      const ops: Op[] =
        from === this.sheetIndex
          ? moveBlockOps(this.sheetIndex, this.clip.range, area)
          : [
              pasteBlockOp(this.sheetIndex, area, block),
              ...optional(pasteNotesOp(this.sheetIndex, this.sheet, area, block)),
              ...optional(pasteLinksOp(this.sheetIndex, this.sheet, area, block)),
              clearCellsOp(from, this.clip.sheet, this.clip.range),
              ...optional(clearNotesOp(from, this.clip.sheet, this.clip.range)),
              ...optional(cutLinksOp(from, this.clip.sheet, this.clip.range)),
            ]
      const after = selectRange(area, { row: area.r0, col: area.c0 })
      this.clip = null
      this.transact(ops, { label: "Move", before, after })
      this.setSelection(after, null)
      return true
    }
    const area = pasteArea(target, block)
    const after = selectRange(area, this.selection.active)
    this.transact(
      [
        pasteBlockOp(this.sheetIndex, area, block),
        ...optional(pasteNotesOp(this.sheetIndex, this.sheet, area, block)),
        ...optional(pasteLinksOp(this.sheetIndex, this.sheet, area, block)),
      ],
      { label: "Paste", before, after },
    )
    this.setSelection(after, null)
    return true
  }

  // ── Rows and columns ───────────────────────────────────────────────────────────────────────────────────────

  /** The rows (or columns) a command acts on: every row touched by the selection, merged. */
  selectedLines(axis: "rows" | "cols"): [number, number][] {
    const spans = this.selection.ranges.map((range) =>
      axis === "rows" ? ([range.r0, range.r1] as [number, number]) : ([range.c0, range.c1] as [number, number]),
    )
    spans.sort((a, b) => a[0] - b[0])
    const merged: [number, number][] = []
    for (const span of spans) {
      const last = merged.at(-1)
      if (last && span[0] <= last[1] + 1) last[1] = Math.max(last[1], span[1])
      else merged.push([...span])
    }
    return merged
  }

  insertLines(axis: "rows" | "cols") {
    // Insert from the bottom up so earlier spans keep their indexes.
    const spans = this.selectedLines(axis).reverse()
    const ops: Op[] = spans.map(([from, to]) => ({
      kind: axis === "rows" ? "insertRows" : "insertCols",
      sheet: this.sheetIndex,
      at: from,
      count: to - from + 1,
    }))
    this.transact(ops, {
      label: axis === "rows" ? "Insert Rows" : "Insert Columns",
      before: this.selection,
      after: this.selection,
    })
  }

  deleteLines(axis: "rows" | "cols") {
    const spans = this.selectedLines(axis).reverse()
    const ops: Op[] = spans.map(([from, to]) => ({
      kind: axis === "rows" ? "deleteRows" : "deleteCols",
      sheet: this.sheetIndex,
      at: from,
      count: to - from + 1,
    }))
    this.transact(ops, {
      label: axis === "rows" ? "Delete Rows" : "Delete Columns",
      before: this.selection,
      after: this.selection,
    })
  }

  /** Insert/Delete cells with a shift (the Insert/Delete dialogs' first two choices). */
  shiftCells(kind: "insert" | "delete", shift: "down" | "right" | "up" | "left") {
    if (this.selection.ranges.length !== 1) return
    const range = this.selection.ranges[0]
    const op: Op =
      kind === "insert"
        ? { kind: "insertCells", sheet: this.sheetIndex, range, shift: shift as "down" | "right" }
        : { kind: "deleteCells", sheet: this.sheetIndex, range, shift: shift as "up" | "left" }
    this.transact([op], {
      label: kind === "insert" ? "Insert Cells" : "Delete Cells",
      before: this.selection,
      after: this.selection,
    })
  }

  /** Ctrl+- / Ctrl++ : whole rows or columns act at once; a cell range asks how to shift (the shell's dialog). */
  insertOrDelete(kind: "insert" | "delete") {
    const range = this.selection.ranges[this.selection.ranges.length - 1]
    if (isWholeRows(range) && !isWholeColumns(range)) {
      kind === "insert" ? this.insertLines("rows") : this.deleteLines("rows")
    } else if (isWholeColumns(range) && !isWholeRows(range)) {
      kind === "insert" ? this.insertLines("cols") : this.deleteLines("cols")
    } else if (isWholeColumns(range) && isWholeRows(range)) {
      if (kind === "delete") this.clearContents()
    } else {
      this.shell.insertDelete({ kind })
    }
  }

  /** `auto`: AutoFit Row Height, a size fitted to the content that later edits may change again. */
  setLineSize(axis: "rows" | "cols", indexes: number[], size: number | null, auto = false) {
    const target = axis === "rows" ? this.sheet.rows : this.sheet.cols
    const entries: AxisEntry[] = indexes.map((index) => ({
      index,
      size,
      hidden: size === 0 ? false : target.isHidden(index),
      ...(auto && size !== null ? { auto: true as const } : {}),
    }))
    if (size === 0) {
      // Excel: a width/height of 0 hides the line.
      this.setHidden(axis, indexes, true)
      return
    }
    this.transact([{ kind: "axis", sheet: this.sheetIndex, axis, entries }], {
      label: axis === "rows" ? "Row Height" : "Column Width",
      before: this.selection,
      after: this.selection,
    })
  }

  setHidden(axis: "rows" | "cols", indexes: number[], hidden: boolean) {
    const target = axis === "rows" ? this.sheet.rows : this.sheet.cols
    const entries = indexes.map((index) => ({
      index,
      size: target.hasCustomSize(index) ? target.storedSize(index) : null,
      hidden,
    }))
    this.transact([{ kind: "axis", sheet: this.sheetIndex, axis, entries }], {
      label: hidden ? "Hide" : "Unhide",
      before: this.selection,
      after: this.selection,
    })
  }

  /** Every index in the selection along `axis` (capped: whole-sheet selections act on used lines only). */
  selectedIndexes(axis: "rows" | "cols"): number[] {
    const out: number[] = []
    for (const [from, to] of this.selectedLines(axis)) {
      const span = to - from + 1
      if (span > 20_000) {
        // Whole sheet: Excel acts on all; we act on the lines that differ from the default plus the used range.
        const used = this.sheet.cells.usedBounds()
        const end = used ? Math.min(to, axis === "rows" ? used.r1 : used.c1) : from
        for (let i = from; i <= end; i++) out.push(i)
        continue
      }
      for (let i = from; i <= to; i++) out.push(i)
    }
    return out
  }

  /** Ctrl+9 / Ctrl+0 hide, Ctrl+Shift+9 / Ctrl+Shift+0 unhide (the selection's rows or columns). */
  hideSelected(axis: "rows" | "cols", hidden: boolean) {
    this.setHidden(axis, this.selectedIndexes(axis), hidden)
  }

  // ── Scrolling ──────────────────────────────────────────────────────────────────────────────────────────────

  /** Width of the frozen columns and height of the frozen rows (sheet pixels). */
  get frozenWidth(): number {
    if (this.split) return Math.min(this.split.width, Math.max(0, this.viewWidth - SPLIT_MIN))
    const f = this.freeze
    return f ? this.sheet.cols.offset(f.col) - this.sheet.cols.offset(f.left) : 0
  }

  get frozenHeight(): number {
    if (this.split) return Math.min(this.split.height, Math.max(0, this.viewHeight - SPLIT_MIN))
    const f = this.freeze
    return f ? this.sheet.rows.offset(f.row) - this.sheet.rows.offset(f.top) : 0
  }

  /** The part of the view that scrolls (the cell area less the frozen panes). */
  get paneWidth(): number {
    return Math.max(1, this.viewWidth - this.frozenWidth)
  }

  get paneHeight(): number {
    return Math.max(1, this.viewHeight - this.frozenHeight)
  }

  /** With frozen panes the scrolling part starts at the first unfrozen row and column. */
  /** View › Split: on (from the active cell, as Excel) or off. Frozen panes turn into a split at the same place. */
  toggleSplit() {
    if (this.edit) this.commitEdit()
    if (this.split) {
      this.removeSplit()
      return
    }
    const rows = this.sheet.rows
    const cols = this.sheet.cols
    const f = this.freeze
    if (f) {
      // The frozen part becomes the top and left panes, scrolled where they were.
      this.freeze = null
      this.split = {
        width: cols.offset(f.col) - cols.offset(f.left),
        height: rows.offset(f.row) - rows.offset(f.top),
        scrollX: cols.offset(f.left),
        scrollY: rows.offset(f.top),
      }
      this.changed()
      return
    }
    const { row: top, col: left } = this.topLeft
    let { row, col } = this.selection.active
    // The split goes above and left of the active cell; from the top-left cell in view, at the middle of the window.
    if (row <= top && col <= left) {
      row = rows.indexAt(this.scrollY + this.viewHeight / 2)
      col = cols.indexAt(this.scrollX + this.viewWidth / 2)
    }
    const width = col > left ? cols.offset(col) - this.scrollX : 0
    const height = row > top ? rows.offset(row) - this.scrollY : 0
    if (width <= 0 && height <= 0) return
    this.split = {
      width: Math.max(0, width),
      height: Math.max(0, height),
      scrollX: this.scrollX,
      scrollY: this.scrollY,
    }
    // The bottom-right pane goes on where the top-left one ends, so nothing on screen moves.
    this.scrollX = this.scrollX + Math.max(0, width)
    this.scrollY = this.scrollY + Math.max(0, height)
    this.changed()
  }

  /** Removes the split: the top-left pane's scroll stays (Excel keeps what was at the top left). */
  removeSplit() {
    const sp = this.split
    if (!sp) return
    this.split = null
    this.scrollX = sp.width ? sp.scrollX : this.scrollX
    this.scrollY = sp.height ? sp.scrollY : this.scrollY
    this.changed()
  }

  /**
   * Moves the split bars (dragging them): sizes in sheet pixels. A bar dragged to the edge goes away; with neither,
   * the split does.
   */
  moveSplit(width: number, height: number) {
    const sp = this.split
    if (!sp) return
    const w = width < SPLIT_MIN || width > this.viewWidth - SPLIT_MIN ? 0 : Math.round(width)
    const h = height < SPLIT_MIN || height > this.viewHeight - SPLIT_MIN ? 0 : Math.round(height)
    if (!w && !h) {
      this.removeSplit()
      return
    }
    // A bar that is new (dragged out again) starts its pane where the bottom-right pane is.
    this.split = {
      width: w,
      height: h,
      scrollX: sp.width ? sp.scrollX : this.scrollX,
      scrollY: sp.height ? sp.scrollY : this.scrollY,
    }
    this.changed()
  }

  /** Scrolls the top or left pane of a split (its wheel or its own scroll bar). */
  scrollSplitTo(x: number, y: number) {
    const sp = this.split
    if (!sp) return
    const nx = Math.max(0, Math.min(x, this.sheet.cols.total - Math.max(1, sp.width)))
    const ny = Math.max(0, Math.min(y, this.sheet.rows.total - Math.max(1, sp.height)))
    if (nx === sp.scrollX && ny === sp.scrollY) return
    this.split = { ...sp, scrollX: nx, scrollY: ny }
    this.scrollVersion++
    for (const listener of this.scrollListeners) listener()
  }

  get minScrollX(): number {
    return this.freeze ? this.sheet.cols.offset(this.freeze.col) : 0
  }

  get minScrollY(): number {
    return this.freeze ? this.sheet.rows.offset(this.freeze.row) : 0
  }

  scrollTo(x: number, y: number) {
    const nx = Math.max(this.minScrollX, Math.min(x, this.colMetric.total - this.paneWidth))
    const ny = Math.max(this.minScrollY, Math.min(y, this.rowMetric.total - this.paneHeight))
    if (nx !== this.scrollX || ny !== this.scrollY) {
      this.scrollX = nx
      this.scrollY = ny
      this.scrollVersion++
      for (const listener of this.scrollListeners) listener()
    }
  }

  /** Scrolls the least amount that shows the whole cell (or as much as fits). Frozen cells are always in view. */
  scrollIntoView(address: CellAddress) {
    const rows = this.sheet.rows
    const cols = this.sheet.cols
    const ys = this.rowMetric
    const xs = this.colMetric
    const f = this.freeze
    let { scrollX, scrollY } = this
    if (!f || address.row >= f.row) {
      const top = ys.pos(address.row)
      const bottom = top + rows.size(address.row)
      if (top < scrollY) scrollY = top
      else if (bottom > scrollY + this.paneHeight) scrollY = Math.min(top, bottom - this.paneHeight)
    }
    if (!f || address.col >= f.col) {
      const left = xs.pos(address.col)
      const right = left + cols.size(address.col)
      if (left < scrollX) scrollX = left
      else if (right > scrollX + this.paneWidth) scrollX = Math.min(left, right - this.paneWidth)
    }
    this.scrollX = Math.max(this.minScrollX, scrollX)
    this.scrollY = Math.max(this.minScrollY, scrollY)
  }

  // ── View: freeze and zoom ─────────────────────────────────────────────────────────────────────────────────

  /** The first row and column in view (of the scrolling part). */
  get topLeft(): CellAddress {
    return { row: this.rowMetric.indexAt(this.scrollY), col: this.colMetric.indexAt(this.scrollX) }
  }

  /**
   * Freeze Panes: rows above and columns left of the active cell stay in view. With the top-left cell in view active,
   * Excel freezes at the middle of the window instead.
   */
  freezePanes() {
    if (this.edit) this.commitEdit()
    // With split panes, Freeze Panes freezes where the split is (Excel).
    const sp = this.split
    if (sp) {
      const top = this.sheet.rows.indexAt(sp.scrollY)
      const left = this.sheet.cols.indexAt(sp.scrollX)
      const row = sp.height ? this.sheet.rows.indexAt(sp.scrollY + sp.height) : top
      const col = sp.width ? this.sheet.cols.indexAt(sp.scrollX + sp.width) : left
      this.split = null
      this.setFreeze({ row, col, top, left })
      return
    }
    const { row: top, col: left } = this.topLeft
    let { row, col } = this.selection.active
    if (row <= top && col <= left) {
      row = this.sheet.rows.indexAt(this.scrollY + this.viewHeight / 2)
      col = this.sheet.cols.indexAt(this.scrollX + this.viewWidth / 2)
    }
    this.setFreeze({ row: Math.max(row, top), col: Math.max(col, left), top, left })
  }

  freezeTopRow() {
    if (this.edit) this.commitEdit()
    const { row: top, col: left } = this.topLeft
    this.setFreeze({ row: top + 1, col: left, top, left })
  }

  freezeFirstColumn() {
    if (this.edit) this.commitEdit()
    const { row: top, col: left } = this.topLeft
    this.setFreeze({ row: top, col: left + 1, top, left })
  }

  unfreeze() {
    const f = this.freeze
    if (!f) return
    this.freeze = null
    // The frozen rows and columns become ordinary ones at the top and left again.
    this.scrollX = this.sheet.cols.offset(f.left)
    this.scrollY = this.sheet.rows.offset(f.top)
    this.changed()
  }

  private setFreeze(f: Freeze) {
    if (f.row === f.top && f.col === f.left) return
    this.split = null
    this.freeze = f
    // The scrolling part goes on where the frozen part ends, so nothing on screen moves.
    this.scrollX = this.sheet.cols.offset(f.col)
    this.scrollY = this.sheet.rows.offset(f.row)
    this.scrollIntoView(this.selection.active)
    this.changed()
  }

  /**
   * View › Normal, Page Break Preview, Page Layout (and the status bar's view buttons). Each view keeps its own zoom
   * (Page Break Preview starts at 60%, as Excel). Page Break Preview shows the page breaks from then on.
   */
  setView(view: SheetViewMode) {
    if (this.edit) this.commitEdit()
    if (view === this.view) return
    this.headerFooter.commit()
    // Page Layout has no frozen panes or split (Excel takes them away).
    if (view === "pageLayout") {
      if (this.freeze) this.unfreeze()
      this.split = null
    }
    const topLeft = this.topLeft
    this.viewZooms = { ...this.viewZooms, [this.view]: this.zoom }
    this.view = view
    this.zoom = this.viewZooms[view] ?? VIEW_ZOOM[view]
    // The same cells at the top left: scroll positions mean other things in Page Layout.
    // A line that starts a page shows that page whole: its margin (with the header) and the gap above it.
    const pageStart = (m: AxisMetric, index: number) => {
      const band = "bands" in m ? (m as PageAxis).bands.find((b) => b.first === index) : undefined
      return band ? Math.max(0, band.start - PAGE_GAP) : m.pos(index)
    }
    this.scrollX = pageStart(this.colMetric, topLeft.col)
    this.scrollY = pageStart(this.rowMetric, topLeft.row)
    if (view !== "normal") this.print.showBreaks = true
    this.scrollIntoView(this.selection.active)
    this.changed()
  }

  /**
   * View › Zoom (10% to 400%). `live` (each step of a pinch) tells only the grid and the zoom readout, as scrolling
   * does, so the ribbon and panes don't re-render with every step; the pinch calls `changed()` once it ends.
   */
  setZoom(zoom: number, live = false) {
    const z = Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom)) * 100) / 100
    if (z === this.zoom) return
    this.zoom = z
    if (!live) return this.changed()
    this.scrollVersion++
    for (const listener of this.scrollListeners) listener()
  }

  /** Zoom to Selection: the selection fills the window (as far as 10%–400% allows), its top-left cell at the top left. */
  zoomToSelection(cssWidth: number, cssHeight: number) {
    const r = this.selection.ranges[this.selection.ranges.length - 1]
    const w = this.sheet.cols.offset(r.c1 + 1) - this.sheet.cols.offset(r.c0)
    const h = this.sheet.rows.offset(r.r1 + 1) - this.sheet.rows.offset(r.r0)
    const z = Math.floor(Math.min(cssWidth / Math.max(1, w), cssHeight / Math.max(1, h)) * 100) / 100
    this.zoom = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z))
    this.scrollX = Math.max(this.minScrollX, this.colMetric.pos(r.c0))
    this.scrollY = Math.max(this.minScrollY, this.rowMetric.pos(r.r0))
    this.changed()
  }

  /** Ctrl+Backspace: bring the active cell into view. */
  revealActive() {
    this.scrollIntoView(this.selection.active)
    this.changed()
  }

  /** Go To (F5 / Ctrl+G) and the Name Box: select a reference; false when it does not parse. */
  goTo(range: Range): void {
    const active = { row: range.r0, col: range.c0 }
    const selection = sameRange(range, rangeOf(active)) ? selectCell(active) : selectRange(range, active)
    this.tabStartCol = null
    this.setSelection(selection, active)
  }

  /**
   * A cell's in-cell dropdown: the function list on a table's total row, or a Data Validation list. `pick` puts an
   * item in.
   */
  cellList(row: number, col: number): { items: string[]; current: string; pick: (item: string) => void } | null {
    const totals = this.tables.totalsList(row, col)
    if (totals) return { ...totals, pick: (item) => this.tables.pickTotal(row, col, item) }
    const items = this.validation.dropdownAt(row, col)
    if (!items) return null
    const value = this.sheet.cells.get(row, col)?.value
    return { items, current: String(value ?? ""), pick: (item) => this.validation.pick(row, col, item) }
  }

  /**
   * Pick From Drop-down List (the cell menu, or Alt+Down on a cell without a list): the text entries in the unbroken
   * run of cells above and below in the column, A to Z, each once (as AutoComplete knows them). Null when there are
   * none. Formulas and numbers aren't offered, as in Excel.
   */
  pickList(row: number, col: number): { items: string[]; current: string; pick: (item: string) => void } | null {
    const cells = this.sheet.cells
    const texts = new Map<string, string>()
    const filled = (r: number) => {
      const value = cells.get(r, col)?.value
      return value !== undefined && value !== ""
    }
    // The run from its top down, so the first spelling in the column is the one kept.
    let top = row
    while (top > 0 && filled(top - 1)) top--
    for (let r = top; r === row || filled(r); r++) {
      if (r === row) continue
      const data = cells.get(r, col)!
      if (typeof data.value === "string" && !data.input?.startsWith("=") && !data.spill) {
        const key = data.value.toLowerCase()
        if (!texts.has(key)) texts.set(key, data.value)
      }
    }
    if (!texts.size) return null
    const items = [...texts.values()].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }))
    const value = cells.get(row, col)?.value
    return { items, current: String(value ?? ""), pick: (item) => this.validation.pick(row, col, item) }
  }

  /** The cell menu's Pick From Drop-down List: the list under the active cell. False when there is nothing to pick. */
  openPickList(): boolean {
    const { row, col } = this.selection.active
    if (!this.shell.openCellList || !(this.cellList(row, col) ?? this.pickList(row, col))) return false
    this.shell.openCellList(!this.cellList(row, col))
    return true
  }

  // ── Keyboard ───────────────────────────────────────────────────────────────────────────────────────────────

  /**
   * Keys pressed while the grid has focus and no entry is open. Returns true when handled (the view then
   * prevents the default). Printable characters are not handled here: the hidden editor receives them and starts
   * an entry.
   */
  keyReady(input: KeyInput): boolean {
    const { key, primary, shift, alt } = input
    const arrows: Record<string, Direction> = {
      ArrowUp: "up",
      ArrowDown: "down",
      ArrowLeft: "left",
      ArrowRight: "right",
    }
    // A selected chart: Delete removes it, Esc goes back to the cells.
    // Selected shapes, pictures and charts take the keys first (Delete, Escape, arrows nudge, Tab to the next).
    if (!alt && this.objects.key(key, primary, shift)) return true
    if (key === "F1" && alt && !primary && !shift) return (this.charts.insert(), true)
    if (arrows[key] && !alt) {
      this.arrow(arrows[key], primary, shift)
      return true
    }
    // Alt+Shift+Right: Group; Alt+Shift+Left: Ungroup (the Group dialog asks rows or columns when it can't tell).
    if (alt && shift && !primary && (key === "ArrowRight" || key === "ArrowLeft")) {
      const group = key === "ArrowRight"
      const axis = this.outline.axisOfSelection()
      if (axis) {
        if (group) this.outline.group(axis)
        else this.outline.ungroup(axis)
      } else this.shell.openDialog?.(group ? "outlineGroup" : "outlineUngroup")
      return true
    }
    // Ctrl+8: show or hide the outline symbols.
    if (primary && !alt && !shift && key === "8") return (this.outline.toggleSymbols(), true)
    switch (key) {
      case "Enter":
        if (alt || primary) return false
        this.enterMove(shift ? "up" : "down")
        return true
      case "Tab":
        if (primary || alt) return false
        this.enterMove(shift ? "left" : "right")
        return true
      case "Home":
        this.home(primary, shift)
        return true
      case "End":
        if (primary) this.ctrlEnd(shift)
        else this.toggleEndMode()
        return true
      case "PageDown":
      case "PageUp":
        if (primary) {
          if (shift && !alt) this.group.extend(key === "PageDown" ? 1 : -1)
          else if (!alt) this.sheets.step(key === "PageDown" ? 1 : -1)
          return true
        }
        this.page(key === "PageDown" ? 1 : -1, alt, shift)
        return true
      case "F2":
        // Shift+F2: new or edit note; Ctrl+Shift+F2: new comment (Excel 365).
        if (shift && primary) this.annotations.newComment()
        else if (shift) this.annotations.newNote()
        else this.beginEdit("edit")
        return true
      case "F3":
        // Ctrl+F3: Name Manager; Ctrl+Shift+F3: Create from Selection; Shift+F3: Insert Function (F3 alone, Paste
        // Name, is not built yet).
        if (alt) return false
        if (primary) {
          this.shell.openDialog?.(shift ? "createNames" : "nameManager")
          return true
        }
        if (!shift) return false
        this.shell.insertFunction?.()
        return true
      case "F7":
        // F7: Spelling; Shift+F7: Thesaurus (Excel).
        if (primary || alt) return false
        if (shift) this.setThesaurusPane(true)
        else this.shell.openDialog?.("spelling")
        return true
      case "F11":
        // Shift+F11 inserts a sheet; F11 alone makes a chart of the data on a chart sheet (Alt+F1: on this sheet).
        if (primary || alt) return false
        if (shift) this.sheets.insert()
        else this.charts.insertSheet()
        return true
      case "F5":
        // Shift+F5 is Find, as in Excel.
        if (shift) this.shell.findReplace?.("find")
        else this.shell.goTo()
        return true
      case "F4":
        // Shift+F4: Find Next with the last search (Ctrl+Shift+F4: the previous one).
        if (!shift || alt) return false
        if (this.find.last.text) this.find.next(this.find.last, primary)
        else this.shell.findReplace?.("find")
        return true
      case "F9":
        // F9 and Shift+F9 calculate what changed (one engine for all sheets); Ctrl+Alt+F9 recalculates everything.
        this.wb.calculate(primary && alt)
        return true
      case "Delete":
        this.clearContents()
        return true
      case "Backspace":
        if (primary) {
          this.revealActive()
          return true
        }
        this.beginEdit("enter", "")
        return true
      case "Escape":
        this.endMode = false
        this.format.cancelPainter()
        this.clearClip()
        this.changed()
        return true
    }
    if (alt && !primary && key === "ArrowDown" && this.data.openMenuAtActive()) return true
    if (alt && !primary && key === "ArrowDown" && this.pivots.openMenuAtActive()) return true
    if (alt && !primary && key === "ArrowDown" && this.openPickList()) return true
    if (primary && !alt) {
      const lower = key.toLowerCase()
      if (lower === "z" && !shift) return (this.undo(), true)
      if (lower === "y" || (lower === "z" && shift)) return (this.redo(), true)
      if (lower === "a" && !shift) return (this.selectAllOrRegion(), true)
      if (key === " " && shift) return (this.selectEverything(), true)
      if (key === " ") return (this.selectColumns(), true)
      if (lower === "g") return (this.shell.goTo(), true)
      if (key === "`" && !shift) return (this.toggleShowFormulas(), true)
      if (lower === "d") return (this.fillFromEdge("down"), true)
      if (lower === "e" && !shift) return (this.data.flashFill(), true)
      if (lower === "r") return (this.fillFromEdge("right"), true)
      if (lower === "u" && shift) return (this.toggleFormulaBarExpanded(), true)
      if (lower === "l" && shift) return (this.data.toggleFilter(), true)
      if ((lower === "t" || lower === "l") && !shift) return (this.tables.openCreate(), true)
      if (lower === "f" && !shift) return (this.shell.findReplace?.("find"), true)
      if (lower === "h" && !shift) return (this.shell.findReplace?.("replace"), true)
      if (lower === "k" && !shift) return (this.shell.openDialog?.("insertLink"), true)
      if (lower === "q" && !shift) return (this.quick.show(), true)
      if (this.formatKey(key, shift)) return true
      if (key === "-") return (this.insertOrDelete("delete"), true)
      if (key === "+" || (key === "=" && shift)) return (this.insertOrDelete("insert"), true)
      if (key === "*" || (key === "8" && shift)) return (this.selectCurrentRegion(), true)
      if (key === "9" || key === "(") return (this.hideSelected("rows", !shift && key === "9"), true)
      if (key === "0" || key === ")") return (this.hideSelected("cols", !shift && key === "0"), true)
    }
    if (!primary && !alt && shift && key === " ") {
      this.selectRowsOfSelection()
      return true
    }
    return false
  }

  /**
   * Keys while an entry is open. Returns "commit-move" style outcomes for the view: true when handled. Arrows
   * commit in enter mode; in edit mode they move the caret (not handled here).
   */
  keyEditing(input: KeyInput): boolean {
    const { key, primary, shift, alt } = input
    if (!this.edit) return false
    if (key === "Escape") {
      this.cancelEdit()
      return true
    }
    if (key === "Enter") {
      if (alt) return false // Alt+Enter: a line break inside the cell; the editor inserts it.
      if (primary && shift) {
        // Ctrl+Shift+Enter: dynamic arrays make it the same as Enter for one cell; the cell stays active.
        this.commitChecked()
        return true
      }
      const move = () => this.enterMove(shift ? "up" : "down")
      // Enter takes live Flash Fill's offer, as in Excel.
      const offer = primary ? null : this.data.liveFlash()
      if (offer) {
        this.data.acceptLiveFlash(offer, move)
        return true
      }
      this.commitChecked(primary, primary ? undefined : move)
      return true
    }
    const list = this.autocomplete
    if (list && key === "Tab" && !shift) return this.acceptAutocomplete()
    if (list && !alt && (key === "ArrowDown" || key === "ArrowUp")) {
      this.movePick(key === "ArrowDown" ? 1 : -1)
      return true
    }
    if (key === "Tab") {
      this.commitChecked(false, () => this.enterMove(shift ? "left" : "right"))
      return true
    }
    if (key === "F2") {
      this.toggleEditMode()
      return true
    }
    if (key === "F4") {
      this.toggleAbsoluteReference()
      return true
    }
    const arrows: Record<string, Direction> = {
      ArrowUp: "up",
      ArrowDown: "down",
      ArrowLeft: "left",
      ArrowRight: "right",
    }
    if (arrows[key] && this.edit.mode === "enter" && !alt && this.canPoint) {
      this.pointArrow(arrows[key], primary, shift)
      return true
    }
    if (arrows[key] && this.edit.mode === "enter" && this.edit.surface === "cell" && !alt) {
      this.commitChecked(false, () => this.arrow(arrows[key], primary, shift))
      return true
    }
    return false
  }

  /** Formatting shortcuts (Ctrl+B, Ctrl+1, Ctrl+Shift+$ …), keyed by the character typed (US layout). */
  private formatKey(key: string, shift: boolean): boolean {
    const f = this.format
    const lower = key.toLowerCase()
    if (!shift) {
      if (lower === "b" || key === "2") return (f.toggle("bold"), true)
      if (lower === "i" || key === "3") return (f.toggle("italic"), true)
      if (lower === "u" || key === "4") return (f.underline(), true)
      if (key === "5") return (f.toggle("strike"), true)
      if (key === "1") return (this.shell.formatCells(), true)
      return false
    }
    const codes: Record<string, string> = {
      "~": NUMBER_FORMATS.general,
      "!": NUMBER_FORMATS.number,
      "@": NUMBER_FORMATS.time,
      "#": NUMBER_FORMATS.dayMonthYear,
      $: NUMBER_FORMATS.currency,
      "%": NUMBER_FORMATS.percent,
      "^": NUMBER_FORMATS.scientific,
    }
    if (codes[key]) return (f.numberFormat(codes[key]), true)
    if (key === "&" || key === "7") return (f.border("outside"), true)
    if (key === "_") return (f.border("none"), true)
    if (lower === "f" || lower === "p") return (this.shell.formatCells("font"), true)
    return false
  }

  /** Ctrl+Shift+U: expand or collapse the formula bar. */
  toggleFormulaBarExpanded() {
    this.formulaBarExpanded = !this.formulaBarExpanded
    this.changed()
  }

  setFormulaBarLines(lines: number) {
    const next = Math.max(1, Math.min(20, Math.round(lines)))
    if (next === this.formulaBarLines) return
    this.formulaBarLines = next
    this.changed()
  }

  toggleShowFormulas() {
    this.showFormulas = !this.showFormulas
    this.changed()
  }

  // ── Status bar ─────────────────────────────────────────────────────────────────────────────────────────────

  get referenceText() {
    return rangeName(selectionBounds(this.selection))
  }
}

function truncate(text: string) {
  return text.length > 24 ? `${text.slice(0, 24)}…` : text
}

/** Shift+arrows on a whole-row (or column) selection keep it whole. */
function keepWholeLines(selection: Selection, focus: CellAddress): CellAddress {
  const range = selection.ranges[selection.ranges.length - 1]
  if (isWholeRows(range) && !isWholeColumns(range)) return { row: focus.row, col: MAX_COLS - 1 }
  if (isWholeColumns(range) && !isWholeRows(range)) return { row: MAX_ROWS - 1, col: focus.col }
  return focus
}

/** Copying whole columns or the whole sheet copies the used part (as the clipboard formats do in Excel). */
function clampToUsed(sheet: Sheet, range: Range): Range {
  const used = sheet.cells.usedBounds()
  if (!isWholeColumns(range) && !isWholeRows(range)) return range
  if (!used) return { ...range, r1: range.r0, c1: range.c0 }
  return {
    r0: range.r0,
    c0: range.c0,
    r1: isWholeColumns(range) ? Math.max(range.r0, used.r1) : range.r1,
    c1: isWholeRows(range) ? Math.max(range.c0, used.c1) : range.c1,
  }
}

function normalizeNewlines(text: string) {
  return text.replace(/\r\n/g, "\n").replace(/\n$/, "")
}

/** Rows of a pasted HTML table (other apps' copies), read with the browser's parser. */
export function htmlTableRows(html: string): string[][] {
  if (typeof DOMParser === "undefined") return []
  const doc = new DOMParser().parseFromString(html, "text/html")
  const rows = [...doc.querySelectorAll("tr")]
  return rows.map((row) =>
    [...row.querySelectorAll("td, th")].map((cell) => (cell as HTMLElement).innerText ?? cell.textContent ?? ""),
  )
}

function primaryPointerIsTouch(): boolean {
  return typeof matchMedia === "function" && matchMedia("(pointer: coarse)").matches
}

/** A sheet's rows and columns on its pages (View › Page Layout). */
function pagesOf(sheet: Sheet, pagination: Pagination): { rows: PageAxis; cols: PageAxis } {
  const used = sheet.cells.usedBounds()
  const lastRow = Math.max(used?.r1 ?? 0, ...pagination.areas.map((a) => a.r1))
  const lastCol = Math.max(used?.c1 ?? 0, ...pagination.areas.map((a) => a.c1))
  return {
    rows: pageAxis(sheet.rows, pagination, "rows", lastRow),
    cols: pageAxis(sheet.cols, pagination, "cols", lastCol),
  }
}
