import {
  columnName,
  formatValue,
  generalAlign,
  isWholeColumns,
  isWholeRows,
  rangeIntersection,
  type CellValue,
  type Range,
} from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { fillHandleRect, rangeRect, type GridLayout, type Line } from "./layout"
import { CELL_FONT, type GridTheme } from "./theme"

/** The canvas calls the drawing needs (a real 2D context satisfies it; tests pass a recorder). */
export type Painter = Pick<
  CanvasRenderingContext2D,
  | "save"
  | "restore"
  | "beginPath"
  | "rect"
  | "clip"
  | "fill"
  | "fillRect"
  | "stroke"
  | "strokeRect"
  | "moveTo"
  | "lineTo"
  | "closePath"
  | "fillText"
  | "setLineDash"
  | "setTransform"
  | "clearRect"
> & {
  fillStyle: CanvasRenderingContext2D["fillStyle"]
  strokeStyle: CanvasRenderingContext2D["strokeStyle"]
  lineWidth: number
  font: string
  textAlign: CanvasTextAlign
  textBaseline: CanvasTextBaseline
  lineDashOffset: number
  globalAlpha: number
}

export type DrawOptions = {
  painter: Painter
  layout: GridLayout
  controller: SheetController
  theme: GridTheme
  dpr: number
  measure: (text: string) => number
  /** Marching ants offset (animates while a copy is on the clipboard). */
  antsOffset: number
}

const PAD = 3
/** Distance from a cell's bottom to the text baseline (bottom-aligned, as Excel's default). */
const BASELINE_FROM_BOTTOM = 5

type CellText = { text: string; align: "left" | "center" | "right"; numeric: boolean }

function cellText(
  controller: SheetController,
  row: number,
  col: number,
  width: number,
  measure: (t: string) => number,
): CellText | null {
  const sheet = controller.sheet
  if (controller.showFormulas) {
    // Ctrl+`: formulas as typed, left-aligned.
    const input = sheet.cells.input(row, col)
    if (input?.startsWith("=")) return { text: input, align: "left", numeric: false }
    // A spill's other cells hold no formula of their own; Excel leaves them blank here.
    if (sheet.cells.isSpill(row, col)) return null
  }
  const value: CellValue | undefined = sheet.cells.value(row, col)
  if (value === undefined) return null
  const style = controller.wb.styles.get(sheet.cells.style(row, col))
  const align = generalAlign(value)
  if (typeof value === "number") {
    const digit = measure("0") || 7
    const maxChars = Math.max(1, Math.floor((width - 2 * PAD + 1) / digit))
    let text = formatValue(value, style.numberFormat, maxChars)
    if (text !== null && measure(text) > width - PAD) {
      text = style.numberFormat ? null : formatValue(value, undefined, Math.max(1, maxChars - 1))
    }
    if (text === null || measure(text) > width - 2) {
      // A number that does not fit shows as #### across the cell (Excel never overflows numbers).
      const count = Math.max(1, Math.floor((width - 2) / (measure("#") || 7)))
      text = "#".repeat(count)
    }
    return { text, align, numeric: true }
  }
  return { text: formatValue(value) ?? "", align, numeric: false }
}

export function drawGrid(options: DrawOptions) {
  const { painter: p, layout, theme, dpr } = options
  p.setTransform(dpr, 0, 0, dpr, 0, 0)
  p.fillStyle = theme.background
  p.fillRect(0, 0, layout.width, layout.height)
  p.font = CELL_FONT
  p.textBaseline = "alphabetic"

  drawCells(options)
  drawSelection(options)
  drawHeaders(options)

  // The thin line under the column headers and right of the row headers.
  p.strokeStyle = theme.headerLine
  p.lineWidth = 1
  p.beginPath()
  p.moveTo(0, layout.headerHeight - 0.5)
  p.lineTo(layout.width, layout.headerHeight - 0.5)
  p.moveTo(layout.rowHeaderWidth - 0.5, 0)
  p.lineTo(layout.rowHeaderWidth - 0.5, layout.height)
  p.stroke()
}

function drawCells({ painter: p, layout, controller, theme, measure }: DrawOptions) {
  const { rows, cols, rowHeaderWidth, headerHeight, width, height } = layout
  if (!rows.length || !cols.length) return
  const sheet = controller.sheet
  p.save()
  p.beginPath()
  p.rect(rowHeaderWidth, headerHeight, width - rowHeaderWidth, height - headerHeight)
  p.clip()

  const editing = controller.edit
  const lastCol = cols[cols.length - 1]
  const gridRight = lastCol.start + lastCol.size

  // Horizontal gridlines: full width.
  p.strokeStyle = theme.gridline
  p.lineWidth = 1
  p.beginPath()
  for (const row of rows) {
    const y = Math.round(row.start + row.size) - 0.5
    p.moveTo(rowHeaderWidth, y)
    p.lineTo(gridRight, y)
  }
  p.stroke()

  // Text, row by row; text that runs into empty neighbours hides the vertical gridlines it crosses.
  const vertical: { x: number; y0: number; y1: number }[] = []
  for (const row of rows) {
    const covered = new Set<number>()
    const baseline = row.start + row.size - BASELINE_FROM_BOTTOM
    const draws: { text: string; x: number; align: CanvasTextAlign; clip: [number, number] }[] = []

    const place = (line: Line, info: CellText, extendLeft: number, extendRight: number) => {
      const textWidth = measure(info.text)
      const inner = line.size - 2 * PAD
      let x: number
      let align: CanvasTextAlign
      if (info.align === "right") {
        x = line.start + line.size - PAD
        align = "right"
      } else if (info.align === "center") {
        x = line.start + line.size / 2
        align = "center"
      } else {
        x = line.start + PAD
        align = "left"
      }
      let clip0 = line.start
      let clip1 = line.start + line.size
      if (!info.numeric && textWidth > inner) {
        clip0 = line.start - extendLeft
        clip1 = line.start + line.size + extendRight
      }
      draws.push({ text: info.text, x, align, clip: [clip0, clip1] })
    }

    /** How far text in `col` may run right (into empty cells) and left, in pixels, marking crossed boundaries. */
    const overflow = (col: number, info: CellText, size: number) => {
      if (info.numeric) return { left: 0, right: 0 }
      const textWidth = measure(info.text) + 2 * PAD
      if (textWidth <= size) return { left: 0, right: 0 }
      const need = info.align === "center" ? (textWidth - size) / 2 : textWidth - size
      let right = 0
      let left = 0
      if (info.align !== "right") {
        for (let c = col + 1; right < need && c < sheet.cols.count; c++) {
          if (sheet.cells.hasValue(row.index, c) || (editing && editing.row === row.index && editing.col === c)) break
          covered.add(c)
          right += sheet.cols.size(c)
        }
      }
      if (info.align !== "left") {
        for (let c = col - 1; left < need && c >= 0; c--) {
          if (sheet.cells.hasValue(row.index, c)) break
          covered.add(c + 1)
          left += sheet.cols.size(c)
        }
      }
      return { left, right }
    }

    // Text that starts left of the view (or right of it, for right-aligned text) may run into it.
    const used = sheet.cells.usedColumns()
    const first = cols[0].index
    const last = cols[cols.length - 1].index
    for (let i = used.length - 1; i >= 0; i--) {
      const c = used[i]
      if (c >= first) continue
      if (!sheet.cells.hasValue(row.index, c)) continue
      const size = sheet.cols.size(c)
      const info = cellText(controller, row.index, c, size, measure)
      if (info && !info.numeric && info.align !== "right") {
        const start = rowHeaderWidth + sheet.cols.offset(c) - layout.scrollX
        const extent = overflow(c, info, size)
        if (start + size + extent.right > rowHeaderWidth)
          place({ index: c, start, size }, info, extent.left, extent.right)
      }
      break
    }
    for (const c of used) {
      if (c <= last) continue
      if (!sheet.cells.hasValue(row.index, c)) continue
      const size = sheet.cols.size(c)
      const info = cellText(controller, row.index, c, size, measure)
      if (info && !info.numeric && info.align === "right") {
        const start = rowHeaderWidth + sheet.cols.offset(c) - layout.scrollX
        const extent = overflow(c, info, size)
        if (start - extent.left < width) place({ index: c, start, size }, info, extent.left, extent.right)
      }
      break
    }

    for (const col of cols) {
      if (editing && editing.row === row.index && editing.col === col.index) continue
      const info = cellText(controller, row.index, col.index, col.size, measure)
      if (!info) continue
      const extent = overflow(col.index, info, col.size)
      place(col, info, extent.left, extent.right)
    }

    for (const col of cols) {
      if (!covered.has(col.index))
        vertical.push({ x: Math.round(col.start) - 0.5, y0: row.start, y1: row.start + row.size })
    }
    // The last column's right edge.
    if (!covered.has(lastCol.index + 1)) {
      vertical.push({ x: Math.round(gridRight) - 0.5, y0: row.start, y1: row.start + row.size })
    }

    p.fillStyle = theme.text
    for (const draw of draws) {
      const lines = draw.text.split("\n")
      p.textAlign = draw.align
      p.save()
      p.beginPath()
      p.rect(draw.clip[0], row.start, draw.clip[1] - draw.clip[0], row.size)
      p.clip()
      lines.forEach((text, i) => {
        p.fillText(text, draw.x, baseline - (lines.length - 1 - i) * 18)
      })
      p.restore()
    }
  }

  p.strokeStyle = theme.gridline
  p.beginPath()
  for (const line of vertical) {
    p.moveTo(line.x, line.y0)
    p.lineTo(line.x, line.y1)
  }
  p.stroke()
  p.restore()
}

function clampRect(rect: { x: number; y: number; w: number; h: number }, layout: GridLayout) {
  // Keep huge rectangles (whole columns) within a few pixels of the canvas, so canvas maths stays exact.
  const x0 = Math.max(rect.x, layout.rowHeaderWidth - 4)
  const y0 = Math.max(rect.y, layout.headerHeight - 4)
  const x1 = Math.min(rect.x + rect.w, layout.width + 4)
  const y1 = Math.min(rect.y + rect.h, layout.height + 4)
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

function drawSelection({ painter: p, layout, controller, theme, antsOffset }: DrawOptions) {
  const { selection } = controller
  p.save()
  p.beginPath()
  p.rect(
    layout.rowHeaderWidth,
    layout.headerHeight,
    layout.width - layout.rowHeaderWidth,
    layout.height - layout.headerHeight,
  )
  p.clip()

  const activeRect = rangeRect(controller, layout, {
    r0: selection.active.row,
    c0: selection.active.col,
    r1: selection.active.row,
    c1: selection.active.col,
  })

  // Tint over every selected range except the active cell, which stays clear.
  p.fillStyle = theme.selectionFill
  selection.ranges.forEach((range) => {
    const rect = clampRect(rangeRect(controller, layout, range), layout)
    if (rect.w <= 0 || rect.h <= 0) return
    const single = range.r0 === range.r1 && range.c0 === range.c1
    if (single && selection.ranges.length === 1) return
    p.beginPath()
    p.rect(rect.x, rect.y, rect.w, rect.h)
    const holdsActive =
      selection.active.row >= range.r0 &&
      selection.active.row <= range.r1 &&
      selection.active.col >= range.c0 &&
      selection.active.col <= range.c1
    if (holdsActive) {
      // Even-odd: the active cell is cut out of the tint.
      p.rect(activeRect.x + activeRect.w, activeRect.y, -activeRect.w, activeRect.h)
    }
    p.fill()
  })

  p.strokeStyle = theme.selectionBorder
  if (selection.ranges.length === 1) {
    const rect = clampRect(rangeRect(controller, layout, selection.ranges[0]), layout)
    p.lineWidth = 2
    p.strokeRect(Math.round(rect.x) - 0, Math.round(rect.y) - 0, Math.round(rect.w) - 1, Math.round(rect.h) - 1)
  } else {
    // Several ranges: a thin outline marks the active cell.
    p.lineWidth = 1
    p.strokeRect(
      Math.round(activeRect.x) + 0.5,
      Math.round(activeRect.y) + 0.5,
      Math.round(activeRect.w) - 2,
      Math.round(activeRect.h) - 2,
    )
  }

  const handle = fillHandleRect(controller, layout)
  if (handle) {
    p.fillStyle = theme.background
    p.fillRect(handle.x - 1, handle.y - 1, handle.size + 2, handle.size + 2)
    p.fillStyle = theme.selectionBorder
    p.fillRect(handle.x, handle.y, handle.size, handle.size)
  }

  if (controller.fillPreview) {
    const rect = clampRect(rangeRect(controller, layout, controller.fillPreview), layout)
    p.strokeStyle = theme.fillPreview
    p.lineWidth = 1
    p.setLineDash([3, 2])
    p.strokeRect(Math.round(rect.x) + 0.5, Math.round(rect.y) + 0.5, Math.round(rect.w) - 2, Math.round(rect.h) - 2)
    p.setLineDash([])
  }

  // A spill range shows a thin blue border while one of its cells is active.
  const spill = controller.edit ? null : controller.spillAt(selection.active.row, selection.active.col)
  if (spill) {
    const rect = clampRect(rangeRect(controller, layout, spill.range), layout)
    p.strokeStyle = theme.spillBorder
    p.lineWidth = 1
    p.strokeRect(Math.round(rect.x) + 0.5, Math.round(rect.y) + 0.5, Math.round(rect.w) - 1, Math.round(rect.h) - 1)
  }

  // References of the formula being typed: tinted and outlined in their colours; the one being pointed at with ants.
  for (const ref of controller.editReferences()) {
    const rect = clampRect(rangeRect(controller, layout, ref.range), layout)
    if (rect.w <= 0 || rect.h <= 0) continue
    const color = theme.refColors[ref.color % theme.refColors.length]
    p.fillStyle = color
    p.globalAlpha = 0.1
    p.fillRect(rect.x, rect.y, rect.w, rect.h)
    p.globalAlpha = 1
    p.strokeStyle = color
    p.lineWidth = 1
    if (ref.pointing) {
      p.setLineDash([4, 3])
      p.lineDashOffset = -antsOffset
    }
    p.strokeRect(Math.round(rect.x) + 0.5, Math.round(rect.y) + 0.5, Math.round(rect.w) - 2, Math.round(rect.h) - 2)
    if (ref.pointing) {
      p.setLineDash([])
      p.lineDashOffset = 0
    } else {
      // Corner handles, as Excel draws them.
      for (const [x, y] of [
        [rect.x, rect.y],
        [rect.x + rect.w - 1, rect.y],
        [rect.x, rect.y + rect.h - 1],
        [rect.x + rect.w - 1, rect.y + rect.h - 1],
      ])
        p.fillRect(Math.round(x) - 2, Math.round(y) - 2, 4, 4)
    }
  }

  if (controller.clip) {
    const rect = clampRect(rangeRect(controller, layout, controller.clip.range), layout)
    p.lineWidth = 2
    p.strokeStyle = theme.background
    p.strokeRect(Math.round(rect.x), Math.round(rect.y), Math.round(rect.w) - 1, Math.round(rect.h) - 1)
    p.strokeStyle = theme.selectionBorder
    p.setLineDash([4, 3])
    p.lineDashOffset = -antsOffset
    p.strokeRect(Math.round(rect.x), Math.round(rect.y), Math.round(rect.w) - 1, Math.round(rect.h) - 1)
    p.setLineDash([])
    p.lineDashOffset = 0
  }
  p.restore()
}

/** Which headers the selection touches: "part" (some cells) or "full" (the whole row/column is selected). */
function headerStates(controller: SheetController, layout: GridLayout) {
  const cols = new Map<number, "part" | "full">()
  const rows = new Map<number, "part" | "full">()
  const visibleCols: Range = {
    r0: 0,
    r1: 0,
    c0: layout.cols[0]?.index ?? 0,
    c1: layout.cols.at(-1)?.index ?? 0,
  }
  const visibleRows: Range = {
    r0: layout.rows[0]?.index ?? 0,
    r1: layout.rows.at(-1)?.index ?? 0,
    c0: 0,
    c1: 0,
  }
  for (const range of controller.selection.ranges) {
    const c = rangeIntersection({ ...visibleCols, r0: range.r0, r1: range.r1 }, range)
    if (c)
      for (let i = c.c0; i <= c.c1; i++)
        if (cols.get(i) !== "full") cols.set(i, isWholeColumns(range) ? "full" : "part")
    const r = rangeIntersection({ ...visibleRows, c0: range.c0, c1: range.c1 }, range)
    if (r)
      for (let i = r.r0; i <= r.r1; i++) if (rows.get(i) !== "full") rows.set(i, isWholeRows(range) ? "full" : "part")
  }
  return { cols, rows }
}

function drawHeaders({ painter: p, layout, controller, theme }: DrawOptions) {
  const { rowHeaderWidth, headerHeight, width, height } = layout
  const states = headerStates(controller, layout)
  p.font = CELL_FONT

  // Column headers.
  p.fillStyle = theme.headerBackground
  p.fillRect(rowHeaderWidth, 0, width - rowHeaderWidth, headerHeight)
  p.textAlign = "center"
  for (const col of layout.cols) {
    const state = states.cols.get(col.index)
    if (state) {
      p.fillStyle = state === "full" ? theme.headerFullBackground : theme.headerSelectedBackground
      p.fillRect(col.start, 0, col.size, headerHeight)
      p.fillStyle = theme.headerSelectedBar
      p.fillRect(col.start, headerHeight - 2, col.size, 2)
    }
    p.fillStyle = state === "full" ? theme.headerFullText : state ? theme.headerSelectedText : theme.headerText
    p.save()
    p.beginPath()
    p.rect(col.start, 0, col.size, headerHeight)
    p.clip()
    p.fillText(columnName(col.index), col.start + col.size / 2, headerHeight - BASELINE_FROM_BOTTOM)
    p.restore()
  }

  // Row headers.
  p.fillStyle = theme.headerBackground
  p.fillRect(0, headerHeight, rowHeaderWidth, height - headerHeight)
  for (const row of layout.rows) {
    const state = states.rows.get(row.index)
    if (state) {
      p.fillStyle = state === "full" ? theme.headerFullBackground : theme.headerSelectedBackground
      p.fillRect(0, row.start, rowHeaderWidth, row.size)
      p.fillStyle = theme.headerSelectedBar
      p.fillRect(rowHeaderWidth - 2, row.start, 2, row.size)
    }
    p.fillStyle = state === "full" ? theme.headerFullText : state ? theme.headerSelectedText : theme.headerText
    p.save()
    p.beginPath()
    p.rect(0, row.start, rowHeaderWidth, row.size)
    p.clip()
    p.fillText(String(row.index + 1), rowHeaderWidth / 2, row.start + row.size - BASELINE_FROM_BOTTOM)
    p.restore()
  }

  // Header separators.
  p.strokeStyle = theme.headerLine
  p.lineWidth = 1
  p.beginPath()
  for (const col of layout.cols) {
    const x = Math.round(col.start + col.size) - 0.5
    p.moveTo(x, 0)
    p.lineTo(x, headerHeight)
  }
  for (const row of layout.rows) {
    const y = Math.round(row.start + row.size) - 0.5
    p.moveTo(0, y)
    p.lineTo(rowHeaderWidth, y)
  }
  p.stroke()

  // The corner (Select All) with its triangle.
  p.fillStyle = theme.headerBackground
  p.fillRect(0, 0, rowHeaderWidth, headerHeight)
  p.fillStyle = theme.cornerTriangle
  p.beginPath()
  p.moveTo(rowHeaderWidth - 4, headerHeight - 4)
  p.lineTo(rowHeaderWidth - 4, headerHeight - 14)
  p.lineTo(rowHeaderWidth - 14, headerHeight - 4)
  p.closePath()
  p.fill()
}
