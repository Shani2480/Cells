import {
  copyPageSetup,
  defaultPaperSize,
  isWholeColumns,
  isWholeRows,
  MARGIN_PRESETS,
  pageBreaks,
  paginate,
  pivotLayout,
  type PageMargins,
  type PageSetup,
  type Pagination,
  type PaginateOptions,
  type Range,
  type Sheet,
} from "@cells/core"
import type { SheetController } from "./controller"

/** The paper a sheet with no paper size of its own prints on, from the system's region. */
export function localPaper(): number {
  const locale = typeof navigator !== "undefined" ? navigator.language : "en-US"
  return defaultPaperSize(locale)
}

/** Which margin preset a sheet's margins are ("custom" when none). */
export function marginPreset(m: PageMargins): keyof typeof MARGIN_PRESETS | "custom" {
  for (const [name, preset] of Object.entries(MARGIN_PRESETS))
    if ((Object.keys(preset) as (keyof PageMargins)[]).every((k) => Math.abs(preset[k] - m[k]) < 1e-6))
      return name as keyof typeof MARGIN_PRESETS
  return "custom"
}

/**
 * Printing: the Page Layout tab's Page Setup, Scale to Fit and Sheet Options, the Page Setup dialog, print area and
 * titles, page breaks, and the pages File › Print shows and prints.
 */
export class PrintCommands {
  private shown = false

  /** Excel shows the page breaks as dashed lines on the sheet once it has been paged (preview, Page Setup). */
  get showBreaks(): boolean {
    return this.shown
  }

  set showBreaks(on: boolean) {
    if (on === this.shown) return
    this.shown = on
    this.c.changed()
  }

  constructor(private readonly c: SheetController) {}

  /** The paper a sheet prints on when it names none. */
  defaultPaper = localPaper()

  setup(sheet: Sheet = this.c.sheet): PageSetup {
    return sheet.pageSetup
  }

  /** One change to a sheet's page setup, as one undo step (Excel's label: "Page Setup"). */
  change(fn: (setup: PageSetup) => void, label = "Page Setup", sheetIndex = this.c.sheetIndex) {
    if (this.c.edit) this.c.commitEdit()
    const sheet = this.c.wb.sheets[sheetIndex]
    if (!sheet) return
    const next = copyPageSetup(sheet.pageSetup)
    fn(next)
    if (JSON.stringify(next) === JSON.stringify(sheet.pageSetup)) return
    const sel = this.c.selection
    this.c.transact([{ kind: "pageSetup", sheet: sheetIndex, setup: next }], { label, before: sel, after: sel })
    this.showBreaks = true
  }

  private cached: { sheet: Sheet; version: number; pagination: Pagination } | null = null

  /** Where the active sheet's pages break (the dashed lines on the sheet): first rows and columns of pages. */
  breaks(): { rows: number[]; cols: number[] } {
    return pageBreaks(this.current())
  }

  /** The active sheet's pages, worked out again only when the workbook changes (drawing asks every frame). */
  current(): Pagination {
    const sheet = this.c.sheet
    const version = this.c.wb.version
    if (this.cached?.sheet !== sheet || this.cached.version !== version)
      this.cached = { sheet, version, pagination: this.pages(sheet) }
    return this.cached.pagination
  }

  /**
   * Page Break Preview: a blue line dragged to the boundary before row or column `to`. A page break moves there
   * (it becomes one set by hand; dragged off its print area it goes away). An edge of the print area moves the
   * print area's edge (the printed cells become the print area when there was none).
   */
  moveLine(line: { axis: "rows" | "cols"; index: number; kind: "edge" | "break"; area: number }, to: number) {
    if (to === line.index) return
    const areas = this.current().areas.map((a) => ({ ...a }))
    const a = areas[line.area]
    if (!a) return
    const rows = line.axis === "rows"
    if (line.kind === "break") {
      const lo = rows ? a.r0 : a.c0
      const hi = rows ? a.r1 + 1 : a.c1 + 1
      const place = (s: PageSetup) => {
        const list = (rows ? s.rowBreaks : s.colBreaks).filter((b) => b !== line.index)
        if (to > lo && to < hi && !list.includes(to)) list.push(to)
        list.sort((x, y) => x - y)
        if (rows) s.rowBreaks = list
        else s.colBreaks = list
      }
      // Dragged past what a page holds, the sheet prints smaller so the page fits (Excel lowers Adjust to).
      const scale = to > line.index && to < hi ? this.scaleToHold(line.axis, to, place) : null
      this.change((s) => {
        place(s)
        if (scale !== null) s.scale = scale
      }, "Move Page Break")
      return
    }
    // An edge: the side nearest `line.index` follows the drag; the area keeps at least one row or column.
    if (rows) {
      if (line.index === a.r0) a.r0 = Math.min(to, a.r1)
      else a.r1 = Math.max(to - 1, a.r0)
    } else if (line.index === a.c0) a.c0 = Math.min(to, a.c1)
    else a.c1 = Math.max(to - 1, a.c0)
    this.change((s) => (s.printArea = areas), "Set Print Area")
  }

  /**
   * The largest "Adjust to" percentage (down to 10%) at which no automatic break falls on the page that now ends
   * before line `to`, or null when it fits already (or the sheet is fitted to pages, which ignores breaks).
   */
  private scaleToHold(axis: "rows" | "cols", to: number, place: (s: PageSetup) => void): number | null {
    const setup = this.c.sheet.pageSetup
    if (setup.fitTo) return null
    const next = copyPageSetup(setup)
    place(next)
    const manual = axis === "rows" ? next.rowBreaks : next.colBreaks
    const from = Math.max(-1, ...manual.filter((b) => b < to))
    const splits = (percent: number) => {
      const p = this.pages(this.c.sheet, {}, { ...next, scale: percent })
      return p.pages.some((page) => {
        const first = axis === "rows" ? page.rows[0] : page.cols[0]
        return first > from && first < to && !manual.includes(first)
      })
    }
    if (!splits(setup.scale)) return null
    for (let percent = setup.scale - 1; percent >= 10; percent--) if (!splits(percent)) return percent
    return 10
  }

  /** The pages a sheet prints (File › Print, the dashed page breaks). */
  pages(sheet: Sheet = this.c.sheet, options: PaginateOptions = {}, setup = sheet.pageSetup): Pagination {
    // Printed headings: the row numbers' column fits the last row number printed (pt).
    const used = setup.printArea?.at(-1)?.r1 ?? sheet.cells.usedBounds()?.r1 ?? 0
    const headings = { width: Math.max(20, String(used + 1).length * 6 + 9), height: 15 }
    return paginate(sheet, this.withPivotTitles(sheet, setup), {
      defaultPaper: this.defaultPaper,
      headings,
      ...options,
    })
  }

  /**
   * PivotTable Options › Printing › Set print titles: with no print titles of its own, the sheet prints a PivotTable's
   * header rows and row label columns on every page (the first such PivotTable's, as Excel).
   */
  private withPivotTitles(sheet: Sheet, setup: PageSetup): PageSetup {
    if (setup.titleRows || setup.titleCols) return setup
    for (const table of sheet.pivots) {
      const cache = this.c.wb.pivotCaches.find((x) => x.id === table.cacheId)
      if (!table.fieldPrintTitles || !cache) continue
      const layout = pivotLayout(table, cache)
      const { r0, c0 } = layout.range
      return {
        ...setup,
        titleRows: layout.firstDataRow > 0 ? [r0, r0 + layout.firstDataRow - 1] : null,
        titleCols: layout.firstDataCol > 0 ? [c0, c0 + layout.firstDataCol - 1] : null,
      }
    }
    return setup
  }

  // ── Page Layout › Page Setup ──

  setMargins(margins: PageMargins) {
    this.change((s) => (s.margins = { ...margins }))
  }

  setOrientation(orientation: PageSetup["orientation"]) {
    this.change((s) => (s.orientation = orientation))
  }

  setPaper(paperSize: number) {
    this.change((s) => (s.paperSize = paperSize))
  }

  /** The selection becomes the print area (Excel: Set Print Area). */
  setPrintArea() {
    const ranges = this.c.selection.ranges.map((r) => ({ ...r }))
    this.change((s) => (s.printArea = ranges), "Set Print Area")
  }

  /** Add to Print Area: shown once there is a print area and the selection isn't inside it. */
  canAddToPrintArea(): boolean {
    const area = this.c.sheet.pageSetup.printArea
    if (!area) return false
    return this.c.selection.ranges.some((r) => !area.some((a) => inside(r, a)))
  }

  addToPrintArea() {
    const area = this.c.sheet.pageSetup.printArea
    if (!area) return this.setPrintArea()
    const extra = this.c.selection.ranges.filter((r) => !area.some((a) => inside(r, a))).map((r) => ({ ...r }))
    this.change((s) => (s.printArea = [...area, ...extra]), "Add to Print Area")
  }

  clearPrintArea() {
    this.change((s) => (s.printArea = null), "Clear Print Area")
  }

  /**
   * Breaks › Insert Page Break: above and left of the active cell (only one of them in the first row or column, or
   * when whole rows or columns are selected).
   */
  insertBreak() {
    const sel = this.c.selection
    const range = sel.ranges[0]
    const { row, col } = sel.active
    const rowBreak = !isWholeColumns(range) && row > 0
    const colBreak = !isWholeRows(range) && col > 0
    if (!rowBreak && !colBreak) return
    this.change((s) => {
      if (rowBreak && !s.rowBreaks.includes(row)) s.rowBreaks = [...s.rowBreaks, row].sort((a, b) => a - b)
      if (colBreak && !s.colBreaks.includes(col)) s.colBreaks = [...s.colBreaks, col].sort((a, b) => a - b)
    }, "Insert Page Break")
  }

  /** Whether Remove Page Break has a break to remove at the active cell. */
  hasBreakAtCell(): boolean {
    const { row, col } = this.c.selection.active
    const s = this.c.sheet.pageSetup
    return s.rowBreaks.includes(row) || s.colBreaks.includes(col)
  }

  removeBreak() {
    const { row, col } = this.c.selection.active
    this.change((s) => {
      s.rowBreaks = s.rowBreaks.filter((r) => r !== row)
      s.colBreaks = s.colBreaks.filter((c) => c !== col)
    }, "Remove Page Break")
  }

  resetBreaks() {
    this.change((s) => {
      s.rowBreaks = []
      s.colBreaks = []
    }, "Reset All Page Breaks")
  }

  // ── Scale to Fit ──

  /** Width or Height: 0 is Automatic; any number fits the sheet to that many pages. */
  setFit(axis: "width" | "height", pages: number) {
    this.change((s) => {
      const fit = s.fitTo ?? { width: 0, height: 0 }
      const next = { ...fit, [axis]: pages }
      s.fitTo = next.width || next.height ? next : null
    })
  }

  /** Scale: only while Width and Height are Automatic (Excel greys it out otherwise). */
  setScale(percent: number) {
    const p = Math.round(percent)
    if (!Number.isFinite(p) || p < 10 || p > 400) return false
    this.change((s) => {
      s.scale = p
      s.fitTo = null
    })
    return true
  }

  // ── Sheet Options ──

  toggleGridlines() {
    this.change((s) => (s.gridlines = !s.gridlines))
  }

  toggleHeadings() {
    this.change((s) => (s.headings = !s.headings))
  }
}

const inside = (r: Range, a: Range) => r.r0 >= a.r0 && r.r1 <= a.r1 && r.c0 >= a.c0 && r.c1 <= a.c1
