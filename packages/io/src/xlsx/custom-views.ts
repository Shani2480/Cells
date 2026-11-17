import {
  columnName,
  copyPageSetup,
  defaultPageSetup,
  MAX_COLS,
  MAX_ROWS,
  quoteSheet,
  type CustomSheetView,
  type CustomView,
  type Sheet,
  type Workbook,
} from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import { at, attrText, kids, type XNode } from "./dom"
import {
  emptyExtras,
  finishPageSetup,
  printElementsXml,
  readPrintArea,
  readPrintTitles,
  readPrintElement,
} from "./page-setup"
import { readViewNode } from "./sheet-read"

/**
 * View › Custom Views in an Excel file, as Excel writes them: `customWorkbookViews` in the workbook (name, id, what
 * the view includes, the active sheet), a `customSheetView` with the same id in each sheet (zoom, scroll, panes,
 * selection, and the print settings), and hidden names Z_<id>_.wvu.Rows / .Cols / .PrintArea / .PrintTitles for
 * the hidden rows and columns and the print area and titles.
 */

/** A sheet's custom view as read, before the workbook's names add its hidden lines. */
export type CustomSheetViewRead = {
  guid: string
  view: CustomSheetView
  hiddenRows: boolean
  hiddenCols: boolean
}

const PRINT_ELEMENTS = new Set(["printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks"])

const yes = (v: string | undefined) => v === "1" || v === "true"
const no = (v: string | undefined) => v === "0" || v === "false"

/** A worksheet's customSheetViews element. */
export function readCustomSheetViews(node: XNode, warnings: string[]): CustomSheetViewRead[] {
  return kids(node, "customSheetView").flatMap((v) => {
    const guid = at(v, "guid")
    if (!guid) return []
    const read = readViewNode(v, warnings)
    const scale = Number(at(v, "scale") ?? 100)
    const mode = at(v, "view")
    let pageSetup = null
    if (v.children.some((k) => PRINT_ELEMENTS.has(k.local))) {
      pageSetup = defaultPageSetup()
      const extras = emptyExtras()
      let saw = false
      for (const k of v.children) {
        if (!PRINT_ELEMENTS.has(k.local)) continue
        if (k.local === "pageSetup") saw = true
        readPrintElement(k.local, k, pageSetup, extras)
      }
      finishPageSetup(pageSetup, yes(at(v, "fitToPage")), saw)
    }
    const view: CustomSheetView = {
      zoom: scale > 0 ? scale / 100 : 1,
      topLeft: read.topLeft,
      active: read.active,
      selection: read.selection,
      freeze: read.freeze,
      showGridlines: !no(at(v, "showGridLines")),
      showHeadings: !no(at(v, "showRowCol")),
      mode: mode === "pageBreakPreview" || mode === "pageLayout" ? mode : "normal",
      hiddenRows: [],
      hiddenCols: [],
      pageSetup,
    }
    return [{ guid, view, hiddenRows: yes(at(v, "hiddenRows")), hiddenCols: yes(at(v, "hiddenColumns")) }]
  })
}

/** A view's part of a defined name: Z_<guid without braces, dashes as _>_.wvu.<what>. */
export function viewNamePrefix(guid: string): string {
  return `Z_${guid.replace(/[{}]/g, "").replace(/-/g, "_")}_.wvu.`
}

/** "$5:$9,$12:$12" (or columns) as spans; null when it isn't plain whole rows or columns. */
function spansOf(text: string, axis: "rows" | "cols"): [number, number][] | null {
  const out: [number, number][] = []
  for (const part of text.split(",")) {
    const ref = part.slice(part.lastIndexOf("!") + 1).replace(/\$/g, "")
    const m = axis === "rows" ? /^(\d+):(\d+)$/.exec(ref) : /^([A-Z]+):([A-Z]+)$/i.exec(ref)
    if (!m) return null
    const index = (s: string) =>
      axis === "rows" ? Number(s) - 1 : [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
    out.push([index(m[1]), index(m[2])])
  }
  return out
}

/**
 * The workbook's custom views from customWorkbookViews, each sheet's customSheetViews and the views' hidden names.
 * `names` are the hidden names read (local to a sheet); those a view uses are taken out of it (returned true).
 */
export function readCustomViews(
  node: XNode | undefined,
  sheets: readonly Sheet[],
  sheetIds: readonly number[],
  perSheet: ReadonlyMap<Sheet, CustomSheetViewRead[]>,
  name: (sheet: Sheet, name: string) => string | null,
): CustomView[] {
  if (!node) return []
  return kids(node, "customWorkbookView").flatMap((w) => {
    const guid = at(w, "guid")
    const viewName = at(w, "name")
    if (!guid || !viewName) return []
    const prefix = viewNamePrefix(guid)
    const active = sheetIds.indexOf(Number(at(w, "activeSheetId") ?? 0))
    const view: CustomView = {
      name: viewName,
      guid,
      printSettings: !no(at(w, "includePrintSettings")),
      hiddenRowsCols: !no(at(w, "includeHiddenRowCol")),
      activeSheet: sheets[active] ?? null,
      sheets: new Map(),
    }
    for (const sheet of sheets) {
      const read = perSheet.get(sheet)?.find((r) => r.guid.toUpperCase() === guid.toUpperCase())
      if (!read) continue
      const v = read.view
      const rows = name(sheet, prefix + "Rows")
      const cols = name(sheet, prefix + "Cols")
      v.hiddenRows = (rows && spansOf(rows, "rows")) || []
      v.hiddenCols = (cols && spansOf(cols, "cols")) || []
      if (v.pageSetup) {
        const area = name(sheet, prefix + "PrintArea")
        const titles = name(sheet, prefix + "PrintTitles")
        if (area) v.pageSetup.printArea = readPrintArea(area)
        if (titles) {
          const t = readPrintTitles(titles)
          if (t) {
            v.pageSetup.titleRows = t.rows
            v.pageSetup.titleCols = t.cols
          }
        }
      }
      view.sheets.set(sheet, v)
    }
    return [view]
  })
}

/** The names a view's sheets use, for the workbook's definedNames. */
export function customViewNamesXml(wb: Workbook): string[] {
  const out: string[] = []
  for (const view of wb.customViews) {
    const prefix = viewNamePrefix(view.guid)
    wb.sheets.forEach((sheet, index) => {
      const v = view.sheets.get(sheet)
      if (!v) return
      const q = quoteSheet(sheet.name) + "!"
      const name = (what: string, ref: string) =>
        out.push(
          `<definedName name="${prefix}${what}" localSheetId="${index}" hidden="1">${escapeXmlText(ref)}</definedName>`,
        )
      if (view.hiddenRowsCols && v.hiddenRows.length)
        name("Rows", v.hiddenRows.map(([a, b]) => `${q}$${a + 1}:$${b + 1}`).join(","))
      if (view.hiddenRowsCols && v.hiddenCols.length)
        name("Cols", v.hiddenCols.map(([a, b]) => `${q}$${columnName(a)}:$${columnName(b)}`).join(","))
      const setup = view.printSettings ? v.pageSetup : null
      if (setup?.printArea?.length) {
        const cell = (row: number, col: number) => `$${columnName(col)}$${row + 1}`
        name(
          "PrintArea",
          setup.printArea
            .map((r) =>
              r.c0 === 0 && r.c1 === MAX_COLS - 1
                ? `${q}$${r.r0 + 1}:$${r.r1 + 1}`
                : r.r0 === 0 && r.r1 === MAX_ROWS - 1
                  ? `${q}$${columnName(r.c0)}:$${columnName(r.c1)}`
                  : `${q}${cell(r.r0, r.c0)}:${cell(r.r1, r.c1)}`,
            )
            .join(","),
        )
      }
      if (setup && (setup.titleRows || setup.titleCols)) {
        const parts: string[] = []
        if (setup.titleCols) parts.push(`${q}$${columnName(setup.titleCols[0])}:$${columnName(setup.titleCols[1])}`)
        if (setup.titleRows) parts.push(`${q}$${setup.titleRows[0] + 1}:$${setup.titleRows[1] + 1}`)
        name("PrintTitles", parts.join(","))
      }
    })
  }
  return out
}

/** The workbook's customWorkbookViews ("" when there are none). Excel needs a window size; Cells gives its own. */
export function customWorkbookViewsXml(wb: Workbook, sheetIds: readonly number[]): string {
  if (!wb.customViews.length) return ""
  const items = wb.customViews.map((v) => {
    const active = v.activeSheet ? wb.sheets.indexOf(v.activeSheet) : 0
    return `<customWorkbookView${attrText({
      name: v.name,
      guid: v.guid,
      includePrintSettings: v.printSettings ? undefined : "0",
      includeHiddenRowCol: v.hiddenRowsCols ? undefined : "0",
      maximized: "1",
      windowWidth: 1040,
      windowHeight: 784,
      activeSheetId: sheetIds[Math.max(0, active)] ?? 1,
    })}/>`
  })
  return `<customWorkbookViews>${items.join("")}</customWorkbookViews>`
}

/** A sheet's customSheetViews ("" when no view covers it). */
export function customSheetViewsXml(wb: Workbook, sheet: Sheet): string {
  const items: string[] = []
  for (const view of wb.customViews) {
    const v = view.sheets.get(sheet)
    if (!v) continue
    const cell = (p: { row: number; col: number }) => columnName(p.col) + (p.row + 1)
    const setup = view.printSettings && v.pageSetup ? v.pageSetup : null
    const attrs = attrText({
      guid: view.guid,
      scale: Math.abs(v.zoom - 1) > 1e-9 ? Math.round(v.zoom * 100) : undefined,
      showGridLines: v.showGridlines ? undefined : "0",
      showRowCol: v.showHeadings ? undefined : "0",
      fitToPage: setup?.fitTo ? "1" : undefined,
      hiddenRows: view.hiddenRowsCols && v.hiddenRows.length ? "1" : undefined,
      hiddenColumns: view.hiddenRowsCols && v.hiddenCols.length ? "1" : undefined,
      view: v.mode !== "normal" ? v.mode : undefined,
      topLeftCell: v.freeze
        ? v.freeze.top || v.freeze.left
          ? cell({ row: v.freeze.top, col: v.freeze.left })
          : undefined
        : v.topLeft.row || v.topLeft.col
          ? cell(v.topLeft)
          : undefined,
    })
    let inner = ""
    const f = v.freeze
    const sqref = v.selection
      .map((r) =>
        r.r0 === r.r1 && r.c0 === r.c1
          ? cell({ row: r.r0, col: r.c0 })
          : `${cell({ row: r.r0, col: r.c0 })}:${cell({ row: r.r1, col: r.c1 })}`,
      )
      .join(" ")
    let pane: string | undefined
    if (f && (f.row > f.top || f.col > f.left)) {
      const xs = f.col - f.left
      const ys = f.row - f.top
      pane = xs && ys ? "bottomRight" : ys ? "bottomLeft" : "topRight"
      const topLeft = { row: Math.max(v.topLeft.row, f.row), col: Math.max(v.topLeft.col, f.col) }
      inner += `<pane${attrText({ xSplit: xs || undefined, ySplit: ys || undefined, topLeftCell: cell(topLeft), activePane: pane, state: "frozen" })}/>`
    }
    inner += `<selection${attrText({ pane, activeCell: cell(v.active), sqref })}/>`
    if (setup) {
      // The print area and titles go in the view's names; the rest as a worksheet writes it, in the view's order.
      const parts = printElementsXml({ ...copyPageSetup(setup) }, undefined)
      inner +=
        parts.rowBreaks +
        parts.colBreaks +
        parts.pageMargins +
        parts.printOptions +
        parts.pageSetup +
        parts.headerFooter
    }
    items.push(`<customSheetView${attrs}>${inner}</customSheetView>`)
  }
  return items.length ? `<customSheetViews>${items.join("")}</customSheetViews>` : ""
}
