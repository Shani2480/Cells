import type { Axis, Range, Sheet } from "@cells/core"
import { HEADER_HEIGHT, type SheetController } from "@/sheet/controller"
import { GRIPPER_HIT, GRIPPER_RADIUS } from "./touch"

/** One drawn row or column: its index and its span in canvas pixels. */
export type Line = { index: number; start: number; size: number }

/**
 * What is on screen: header sizes and the visible rows and columns with their canvas positions. All sizes are in
 * sheet pixels at 100%; the view scales the canvas by the zoom.
 *
 * With frozen panes the cell area has up to four parts (`panes`), each a layout of its own with a clip rectangle:
 * the frozen rows and columns, and the part that scrolls. The whole layout's `rows` and `cols` list the frozen lines
 * first, then the scrolling ones, so headers and hit tests read them in screen order.
 */
export type GridLayout = {
  width: number
  height: number
  rowHeaderWidth: number
  headerHeight: number
  rows: Line[]
  cols: Line[]
  /** The scrolling part: sheet offsets at its top-left corner, and where that corner is on the canvas. */
  scrollX: number
  scrollY: number
  originX: number
  originY: number
  /** First unfrozen row and column, and the sheet offsets the frozen part shows from (null: nothing frozen). */
  frozen: { row: number; col: number; scrollX: number; scrollY: number } | null
  /** Where this layout may draw on the canvas (a pane's own rectangle; the whole cell area otherwise). */
  clip: { x0: number; y0: number; x1: number; y1: number }
  /** The parts of the cell area, drawn one by one (one when nothing is frozen). Panes have none of their own. */
  panes: GridLayout[]
}

function visibleLines(axis: Axis, scroll: number, origin: number, extent: number, end = axis.count): Line[] {
  const lines: Line[] = []
  let index = axis.indexAt(scroll)
  let start = origin + axis.offset(index) - scroll
  while (index < end && start < origin + extent) {
    const size = axis.size(index)
    if (size > 0) lines.push({ index, start, size })
    start += size
    index++
  }
  return lines
}

/** Excel widens the row header as row numbers get longer. */
export function rowHeaderWidthFor(lastRow: number, measure: (text: string) => number) {
  return Math.max(26, Math.ceil(measure(String(lastRow + 1))) + 12)
}

/** `width` and `height` are the canvas size in sheet pixels (CSS pixels divided by the zoom). */
export function computeLayout(
  controller: SheetController,
  width: number,
  height: number,
  measure: (text: string) => number,
): GridLayout {
  const sheet = controller.sheet
  const f = controller.freeze
  // View › Headings off: no header row or column (the cells start at the edge).
  const headerHeight = controller.showHeadings ? HEADER_HEIGHT : 0
  const frozenH = controller.frozenHeight
  const frozenW = controller.frozenWidth
  const originY = headerHeight + frozenH
  const fScrollY = f ? sheet.rows.offset(f.top) : 0
  const fScrollX = f ? sheet.cols.offset(f.left) : 0
  const frozenRows = f && frozenH > 0 ? visibleLines(sheet.rows, fScrollY, headerHeight, frozenH, f.row) : []
  const scrollRows = visibleLines(sheet.rows, controller.scrollY, originY, height - originY)
  // The row header's width depends on the last visible row number, which depends on the header height only.
  const lastRow = scrollRows.at(-1)?.index ?? frozenRows.at(-1)?.index ?? 0
  const rowHeaderWidth = controller.showHeadings ? rowHeaderWidthFor(lastRow, measure) : 0
  const originX = rowHeaderWidth + frozenW
  const frozenCols = f && frozenW > 0 ? visibleLines(sheet.cols, fScrollX, rowHeaderWidth, frozenW, f.col) : []
  const scrollCols = visibleLines(sheet.cols, controller.scrollX, originX, width - originX)
  const base = { width, height, rowHeaderWidth, headerHeight, frozen: null, panes: [] }
  const pane = (rows: Line[], cols: Line[], top: boolean, left: boolean): GridLayout => ({
    ...base,
    rows,
    cols,
    scrollX: left ? fScrollX : controller.scrollX,
    scrollY: top ? fScrollY : controller.scrollY,
    originX: left ? rowHeaderWidth : originX,
    originY: top ? headerHeight : originY,
    clip: {
      x0: left ? rowHeaderWidth : originX,
      y0: top ? headerHeight : originY,
      x1: left ? originX : width,
      y1: top ? originY : height,
    },
  })
  const panes = [pane(scrollRows, scrollCols, false, false)]
  if (frozenRows.length) panes.push(pane(frozenRows, scrollCols, true, false))
  if (frozenCols.length) panes.push(pane(scrollRows, frozenCols, false, true))
  if (frozenRows.length && frozenCols.length) panes.push(pane(frozenRows, frozenCols, true, true))
  return {
    ...base,
    rows: [...frozenRows, ...scrollRows],
    cols: [...frozenCols, ...scrollCols],
    scrollX: controller.scrollX,
    scrollY: controller.scrollY,
    originX,
    originY,
    frozen: f && (frozenH > 0 || frozenW > 0) ? { row: f.row, col: f.col, scrollX: fScrollX, scrollY: fScrollY } : null,
    clip: { x0: rowHeaderWidth, y0: headerHeight, x1: width, y1: height },
    panes,
  }
}

/** A column's left edge on the canvas (in the frozen part when it is frozen). */
export function colX(sheet: Sheet, layout: GridLayout, col: number): number {
  const f = layout.frozen
  if (f && col < f.col) return layout.rowHeaderWidth + sheet.cols.offset(col) - f.scrollX
  return layout.originX + sheet.cols.offset(col) - layout.scrollX
}

/** A row's top edge on the canvas. */
export function rowY(sheet: Sheet, layout: GridLayout, row: number): number {
  const f = layout.frozen
  if (f && row < f.row) return layout.headerHeight + sheet.rows.offset(row) - f.scrollY
  return layout.originY + sheet.rows.offset(row) - layout.scrollY
}

/** A range's rectangle in canvas pixels (may extend past the canvas). */
export function rangeRect(controller: SheetController, layout: GridLayout, range: Range) {
  const sheet = controller.sheet
  const x0 = colX(sheet, layout, range.c0)
  const x1 = colX(sheet, layout, range.c1) + sheet.cols.size(range.c1)
  const y0 = rowY(sheet, layout, range.r0)
  const y1 = rowY(sheet, layout, range.r1) + sheet.rows.size(range.r1)
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

/** The column under a canvas x (frozen part first). */
function colAtX(sheet: Sheet, layout: GridLayout, x: number): number {
  const f = layout.frozen
  if (f && x < layout.originX)
    return Math.min(f.col - 1, indexAtPixel(sheet.cols, x - layout.rowHeaderWidth + f.scrollX))
  return indexAtPixel(sheet.cols, x - layout.originX + layout.scrollX)
}

function rowAtY(sheet: Sheet, layout: GridLayout, y: number): number {
  const f = layout.frozen
  if (f && y < layout.originY) return Math.min(f.row - 1, indexAtPixel(sheet.rows, y - layout.headerHeight + f.scrollY))
  return indexAtPixel(sheet.rows, y - layout.originY + layout.scrollY)
}

export type Hit =
  | { zone: "corner" }
  | { zone: "colHeader"; col: number }
  | { zone: "rowHeader"; row: number }
  | { zone: "colResize"; col: number }
  | { zone: "rowResize"; row: number }
  | { zone: "fillHandle" }
  /** A touch gripper on the selection's top-left (`start`) or bottom-right (`end`) corner. */
  | { zone: "gripper"; end: "start" | "end" }
  | { zone: "cell"; row: number; col: number }

const RESIZE_SLOP = 3

/** What is under a point of the canvas (in CSS pixels). Points outside the cell area clamp to the nearest cell. */
export function hitTest(controller: SheetController, layout: GridLayout, x: number, y: number): Hit {
  const sheet = controller.sheet
  const inColHeader = y < layout.headerHeight
  const inRowHeader = x < layout.rowHeaderWidth
  if (inColHeader && inRowHeader) {
    return { zone: "corner" }
  }
  if (inColHeader) {
    for (const col of layout.cols) {
      const right = col.start + col.size
      if (Math.abs(x - right) <= RESIZE_SLOP) return { zone: "colResize", col: col.index }
    }
    return { zone: "colHeader", col: colAtX(sheet, layout, x) }
  }
  if (inRowHeader) {
    for (const row of layout.rows) {
      const bottom = row.start + row.size
      if (Math.abs(y - bottom) <= RESIZE_SLOP - 1) return { zone: "rowResize", row: row.index }
    }
    return { zone: "rowHeader", row: rowAtY(sheet, layout, y) }
  }
  const grippers = selectionGrippers(controller, layout)
  if (grippers) {
    for (const end of ["end", "start"] as const) {
      const g = grippers[end]
      if (
        grippers[end === "start" ? "startShown" : "endShown"] &&
        Math.hypot(x - g.x, y - g.y) <= GRIPPER_HIT / controller.zoom
      )
        return { zone: "gripper", end }
    }
  }
  const handle = fillHandleRect(controller, layout)
  if (
    handle &&
    x >= handle.x - 2 &&
    x <= handle.x + handle.size + 2 &&
    y >= handle.y - 2 &&
    y <= handle.y + handle.size + 2
  ) {
    return { zone: "fillHandle" }
  }
  return {
    zone: "cell",
    row: rowAtY(sheet, layout, y),
    col: colAtX(sheet, layout, x),
  }
}

/** The cell under a point, also for points dragged outside the cell area (clamped to the sheet). */
export function cellAt(controller: SheetController, layout: GridLayout, x: number, y: number) {
  const sheet = controller.sheet
  // While dragging, a point above or left of the scrolling part clamps to it when nothing frozen is there.
  return { row: rowAtY(sheet, layout, y), col: colAtX(sheet, layout, x) }
}

function indexAtPixel(axis: Axis, px: number) {
  return axis.indexAt(Math.max(0, px))
}

/** Touch grippers: circles on the top-left and bottom-right corners of a single selected range (Excel on tablets). */
export function selectionGrippers(controller: SheetController, layout: GridLayout) {
  const { selection } = controller
  if (!controller.touchHandles || selection.ranges.length !== 1 || controller.edit) return null
  const rect = rangeRect(controller, layout, selection.ranges[0])
  const inView = (p: { x: number; y: number }) =>
    p.x >= layout.rowHeaderWidth - GRIPPER_RADIUS &&
    p.y >= layout.headerHeight - GRIPPER_RADIUS &&
    p.x <= layout.width + GRIPPER_RADIUS &&
    p.y <= layout.height + GRIPPER_RADIUS
  const start = { x: rect.x, y: rect.y }
  const end = { x: rect.x + rect.w, y: rect.y + rect.h }
  if (!inView(start) && !inView(end)) return null
  return { start, end, startShown: inView(start), endShown: inView(end) }
}

export const FILL_HANDLE_SIZE = 7

/** The fill handle sits on the bottom-right corner of a single selected range. */
export function fillHandleRect(controller: SheetController, layout: GridLayout) {
  const { selection } = controller
  if (selection.ranges.length !== 1 || controller.edit || controller.touchHandles) return null
  const rect = rangeRect(controller, layout, selection.ranges[0])
  const x = rect.x + rect.w - Math.ceil(FILL_HANDLE_SIZE / 2)
  const y = rect.y + rect.h - Math.ceil(FILL_HANDLE_SIZE / 2)
  if (x < layout.rowHeaderWidth - 2 || y < layout.headerHeight - 2 || x > layout.width || y > layout.height) return null
  return { x, y, size: FILL_HANDLE_SIZE }
}
