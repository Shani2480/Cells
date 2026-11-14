import type { Axis, Range } from "@cells/core"
import { HEADER_HEIGHT, type SheetController } from "@/sheet/controller"

/** One drawn row or column: its index and its span in canvas pixels. */
export type Line = { index: number; start: number; size: number }

/** What is on screen: header sizes and the visible rows and columns with their canvas positions. */
export type GridLayout = {
  width: number
  height: number
  rowHeaderWidth: number
  headerHeight: number
  rows: Line[]
  cols: Line[]
  scrollX: number
  scrollY: number
}

function visibleLines(axis: Axis, scroll: number, origin: number, extent: number): Line[] {
  const lines: Line[] = []
  let index = axis.indexAt(scroll)
  let start = origin + axis.offset(index) - scroll
  while (index < axis.count && start < origin + extent) {
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

export function computeLayout(
  controller: SheetController,
  width: number,
  height: number,
  measure: (text: string) => number,
): GridLayout {
  const sheet = controller.sheet
  const headerHeight = HEADER_HEIGHT
  // The row header's width depends on the last visible row number, which depends on the header height only.
  const roughRows = visibleLines(sheet.rows, controller.scrollY, headerHeight, height - headerHeight)
  const lastRow = roughRows.at(-1)?.index ?? 0
  const rowHeaderWidth = rowHeaderWidthFor(lastRow, measure)
  return {
    width,
    height,
    rowHeaderWidth,
    headerHeight,
    rows: roughRows,
    cols: visibleLines(sheet.cols, controller.scrollX, rowHeaderWidth, width - rowHeaderWidth),
    scrollX: controller.scrollX,
    scrollY: controller.scrollY,
  }
}

/** A range's rectangle in canvas pixels (may extend past the canvas). */
export function rangeRect(controller: SheetController, layout: GridLayout, range: Range) {
  const sheet = controller.sheet
  const x0 = layout.rowHeaderWidth + sheet.cols.offset(range.c0) - layout.scrollX
  const x1 = layout.rowHeaderWidth + sheet.cols.offset(range.c1 + 1) - layout.scrollX
  const y0 = layout.headerHeight + sheet.rows.offset(range.r0) - layout.scrollY
  const y1 = layout.headerHeight + sheet.rows.offset(range.r1 + 1) - layout.scrollY
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

export type Hit =
  | { zone: "corner" }
  | { zone: "colHeader"; col: number }
  | { zone: "rowHeader"; row: number }
  | { zone: "colResize"; col: number }
  | { zone: "rowResize"; row: number }
  | { zone: "fillHandle" }
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
    return { zone: "colHeader", col: indexAtPixel(sheet.cols, x - layout.rowHeaderWidth + layout.scrollX) }
  }
  if (inRowHeader) {
    for (const row of layout.rows) {
      const bottom = row.start + row.size
      if (Math.abs(y - bottom) <= RESIZE_SLOP - 1) return { zone: "rowResize", row: row.index }
    }
    return { zone: "rowHeader", row: indexAtPixel(sheet.rows, y - layout.headerHeight + layout.scrollY) }
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
    row: indexAtPixel(sheet.rows, y - layout.headerHeight + layout.scrollY),
    col: indexAtPixel(sheet.cols, x - layout.rowHeaderWidth + layout.scrollX),
  }
}

/** The cell under a point, also for points dragged outside the cell area (clamped to the sheet). */
export function cellAt(controller: SheetController, layout: GridLayout, x: number, y: number) {
  const sheet = controller.sheet
  return {
    row: indexAtPixel(sheet.rows, y - layout.headerHeight + layout.scrollY),
    col: indexAtPixel(sheet.cols, x - layout.rowHeaderWidth + layout.scrollX),
  }
}

function indexAtPixel(axis: Axis, px: number) {
  return axis.indexAt(Math.max(0, px))
}

export const FILL_HANDLE_SIZE = 7

/** The fill handle sits on the bottom-right corner of a single selected range. */
export function fillHandleRect(controller: SheetController, layout: GridLayout) {
  const { selection } = controller
  if (selection.ranges.length !== 1 || controller.edit) return null
  const rect = rangeRect(controller, layout, selection.ranges[0])
  const x = rect.x + rect.w - Math.ceil(FILL_HANDLE_SIZE / 2)
  const y = rect.y + rect.h - Math.ceil(FILL_HANDLE_SIZE / 2)
  if (x < layout.rowHeaderWidth - 2 || y < layout.headerHeight - 2 || x > layout.width || y > layout.height) return null
  return { x, y, size: FILL_HANDLE_SIZE }
}
