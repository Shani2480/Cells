import { columnName, isWholeColumns, isWholeRows, rangeIntersection, type Range } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import type { Measure } from "./cell-look"
import { drawCells } from "./draw-cells"
import { colX, fillHandleRect, rangeRect, rowY, selectionGrippers, type GridLayout } from "./layout"
import { GRIPPER_RADIUS } from "./touch"
import { CELL_FONT, type GridTheme } from "./theme"

/** The canvas calls the drawing needs (a real 2D context satisfies it; tests pass a recorder). */
export type Painter = Pick<
  CanvasRenderingContext2D,
  | "save"
  | "restore"
  | "beginPath"
  | "arc"
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
  | "translate"
  | "rotate"
  | "createLinearGradient"
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
  measure: Measure
  /** Marching ants offset (animates while a copy is on the clipboard). */
  antsOffset: number
}

/** Distance from a cell's bottom to the text baseline (bottom-aligned, as Excel's default). */
const BASELINE_FROM_BOTTOM = 5

export function drawGrid(options: DrawOptions) {
  const { painter: p, layout, theme, dpr } = options
  p.setTransform(dpr, 0, 0, dpr, 0, 0)
  p.fillStyle = theme.background
  p.fillRect(0, 0, layout.width, layout.height)
  p.font = CELL_FONT
  p.textBaseline = "alphabetic"

  // Each part of the cell area (frozen rows and columns, the scrolling part) draws in its own rectangle.
  for (const pane of layout.panes) {
    drawCells({ ...options, layout: pane })
    drawPageBreaks({ ...options, layout: pane })
    drawMarks({ ...options, layout: pane })
    p.font = CELL_FONT
    drawSelection({ ...options, layout: pane })
  }
  drawHeaders(options)

  // Frozen panes: a thin dark line where the frozen rows and columns end, across headers and cells (Excel).
  if (layout.frozen) {
    p.strokeStyle = theme.freezeLine
    p.lineWidth = 1
    p.beginPath()
    if (layout.originY > layout.headerHeight) {
      p.moveTo(0, layout.originY - 0.5)
      p.lineTo(layout.width, layout.originY - 0.5)
    }
    if (layout.originX > layout.rowHeaderWidth) {
      p.moveTo(layout.originX - 0.5, 0)
      p.lineTo(layout.originX - 0.5, layout.height)
    }
    p.stroke()
  }

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

/** Page breaks as Excel shows them in Normal view once the sheet has been paged: dashed lines between pages. */
function drawPageBreaks({ painter: p, layout, controller, theme }: DrawOptions) {
  if (!controller.print.showBreaks || !layout.rows.length || !layout.cols.length) return
  const { rows, cols } = controller.print.breaks()
  if (!rows.length && !cols.length) return
  const sheet = controller.sheet
  const { clip } = layout
  const first = (lines: { index: number }[]) => lines[0].index
  const last = (lines: { index: number }[]) => lines[lines.length - 1].index
  p.save()
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  p.clip()
  p.strokeStyle = theme.pageBreak
  p.lineWidth = 1
  p.setLineDash([3, 2])
  p.beginPath()
  for (const r of rows) {
    if (r <= first(layout.rows) || r > last(layout.rows) + 1) continue
    const y = Math.round(rowY(sheet, layout, r)) - 0.5
    p.moveTo(clip.x0, y)
    p.lineTo(clip.x1, y)
  }
  for (const c of cols) {
    if (c <= first(layout.cols) || c > last(layout.cols) + 1) continue
    const x = Math.round(colX(sheet, layout, c)) - 0.5
    p.moveTo(x, clip.y0)
    p.lineTo(x, clip.y1)
  }
  p.stroke()
  p.restore()
}

/**
 * Notes and comments: a small triangle in the cell's top-right corner, red for a note, purple for a comment (grey
 * when the thread is resolved), as Excel marks them.
 */
function drawMarks({ painter: p, layout, controller, theme }: DrawOptions) {
  const sheet = controller.sheet
  if (!sheet.annotations.size || !layout.rows.length || !layout.cols.length) return
  const rows = new Map(layout.rows.map((l) => [l.index, l]))
  const cols = new Map(layout.cols.map((l) => [l.index, l]))
  p.save()
  p.beginPath()
  p.rect(layout.clip.x0, layout.clip.y0, layout.clip.x1 - layout.clip.x0, layout.clip.y1 - layout.clip.y0)
  p.clip()
  for (const a of sheet.annotations.all()) {
    const merge = sheet.merges.at(a.row, a.col)
    const top = rows.get(merge ? merge.r0 : a.row)
    const right = cols.get(merge ? merge.c1 : a.col)
    if (!top || !right) continue
    const x = right.start + right.size
    const y = top.start
    const v = a.value
    const size = v.kind === "note" ? 6 : 7
    p.fillStyle = v.kind === "note" ? theme.noteMark : v.resolved ? theme.freezeLine : theme.commentMark
    p.beginPath()
    p.moveTo(x - size - 1, y + 1)
    p.lineTo(x - 1, y + 1)
    p.lineTo(x - 1, y + size + 1)
    p.closePath()
    p.fill()
  }
  p.restore()
}

function clampRect(rect: { x: number; y: number; w: number; h: number }, layout: GridLayout) {
  // Keep huge rectangles (whole columns) within a few pixels of the pane, so canvas maths stays exact.
  const x0 = Math.max(rect.x, layout.clip.x0 - 4)
  const y0 = Math.max(rect.y, layout.clip.y0 - 4)
  const x1 = Math.min(rect.x + rect.w, layout.clip.x1 + 4)
  const y1 = Math.min(rect.y + rect.h, layout.clip.y1 + 4)
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

function drawSelection({ painter: p, layout, controller, theme, antsOffset }: DrawOptions) {
  const { selection } = controller
  p.save()
  p.beginPath()
  p.rect(layout.clip.x0, layout.clip.y0, layout.clip.x1 - layout.clip.x0, layout.clip.y1 - layout.clip.y0)
  p.clip()

  // The active cell, or the whole merged cell it belongs to.
  const activeArea = controller.activeArea()
  const activeRect = rangeRect(controller, layout, activeArea)

  // Tint over every selected range except the active cell, which stays clear.
  p.fillStyle = theme.selectionFill
  selection.ranges.forEach((range) => {
    const rect = clampRect(rangeRect(controller, layout, range), layout)
    if (rect.w <= 0 || rect.h <= 0) return
    const single =
      range.r0 === activeArea.r0 &&
      range.r1 === activeArea.r1 &&
      range.c0 === activeArea.c0 &&
      range.c1 === activeArea.c1
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

  // Touch grippers: white circles ringed in the selection colour.
  const grippers = selectionGrippers(controller, layout)
  if (grippers) {
    p.lineWidth = 1.5
    for (const [g, shown] of [
      [grippers.start, grippers.startShown],
      [grippers.end, grippers.endShown],
    ] as const) {
      if (!shown) continue
      p.beginPath()
      p.arc(g.x, g.y, GRIPPER_RADIUS, 0, Math.PI * 2)
      p.fillStyle = theme.background
      p.fill()
      p.strokeStyle = theme.selectionBorder
      p.stroke()
    }
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
  if (!controller.showHeadings) return
  const states = headerStates(controller, layout)
  p.font = CELL_FONT

  // Column headers.
  p.fillStyle = theme.headerBackground
  p.fillRect(rowHeaderWidth, 0, width - rowHeaderWidth, headerHeight)
  p.textAlign = "center"
  const frozenCol = layout.frozen?.col ?? 0
  const frozenRow = layout.frozen?.row ?? 0
  /** The part of the header row (or column) a line belongs to: frozen lines don't scroll under the others. */
  const colBand = (index: number) => (index < frozenCol ? [rowHeaderWidth, layout.originX] : [layout.originX, width])
  const rowBand = (index: number) => (index < frozenRow ? [headerHeight, layout.originY] : [layout.originY, height])
  for (const col of layout.cols) {
    const [b0, b1] = colBand(col.index)
    p.save()
    p.beginPath()
    p.rect(b0, 0, b1 - b0, headerHeight)
    p.clip()
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
    p.restore()
  }

  // Row headers (a filtered range's row numbers in blue, as Excel).
  const filtered = controller.data
    .hosts()
    .filter((h) => Object.keys(h.filter.columns).length)
    .map((h) => h.filter.range)
  p.fillStyle = theme.headerBackground
  p.fillRect(0, headerHeight, rowHeaderWidth, height - headerHeight)
  for (const row of layout.rows) {
    const [b0, b1] = rowBand(row.index)
    p.save()
    p.beginPath()
    p.rect(0, b0, rowHeaderWidth, b1 - b0)
    p.clip()
    const state = states.rows.get(row.index)
    if (state) {
      p.fillStyle = state === "full" ? theme.headerFullBackground : theme.headerSelectedBackground
      p.fillRect(0, row.start, rowHeaderWidth, row.size)
      p.fillStyle = theme.headerSelectedBar
      p.fillRect(rowHeaderWidth - 2, row.start, 2, row.size)
    }
    const inFilter = filtered.some((f) => row.index > f.r0 && row.index <= f.r1)
    p.fillStyle =
      state === "full"
        ? theme.headerFullText
        : inFilter
          ? theme.headerFilteredText
          : state
            ? theme.headerSelectedText
            : theme.headerText
    p.save()
    p.beginPath()
    p.rect(0, row.start, rowHeaderWidth, row.size)
    p.clip()
    p.fillText(String(row.index + 1), rowHeaderWidth / 2, row.start + row.size - BASELINE_FROM_BOTTOM)
    p.restore()
    p.restore()
  }

  // Header separators.
  p.strokeStyle = theme.headerLine
  p.lineWidth = 1
  p.beginPath()
  for (const col of layout.cols) {
    const x = Math.round(col.start + col.size) - 0.5
    const [b0, b1] = colBand(col.index)
    if (x < b0 || x > b1) continue
    p.moveTo(x, 0)
    p.lineTo(x, headerHeight)
  }
  for (const row of layout.rows) {
    const y = Math.round(row.start + row.size) - 0.5
    const [b0, b1] = rowBand(row.index)
    if (y < b0 || y > b1) continue
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
