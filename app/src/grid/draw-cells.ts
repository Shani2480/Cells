import {
  hasFormatMarks,
  tableAt,
  tableCellLook,
  withTableLook,
  type Border,
  type CellValue,
  type CFResult,
  type Range,
  tableStyleDef,
} from "@cells/core"
import { drawIcon } from "./cf-icons"
import { drawSparkline, sparkAt, sparkItems, type SparkItem } from "./sparklines"
import type { SheetController } from "@/sheet/controller"
import {
  darkVariant,
  fontFor,
  formatColor,
  layoutMarks,
  lookOf,
  prefersDarkText,
  shownText,
  textWidth,
  type CellLook,
  type Measure,
  type Shown,
} from "./cell-look"
import { colX, rowY, type GridLayout, type Line } from "./layout"
import type { GridTheme } from "./theme"

/**
 * The cells: fills, gridlines, text and borders, in Excel's order. Each cell draws with its own format (font, colour,
 * alignment, wrap, rotation…); merged cells draw as one cell with their top-left cell's content.
 */

/** The painter calls this module uses (a subset of the canvas context). */
export type CellPainter = Pick<
  CanvasRenderingContext2D,
  | "save"
  | "restore"
  | "beginPath"
  | "rect"
  | "clip"
  | "fillRect"
  | "stroke"
  | "moveTo"
  | "lineTo"
  | "fillText"
  | "setLineDash"
  | "translate"
  | "rotate"
  | "fill"
  | "closePath"
  | "strokeRect"
  | "createLinearGradient"
> & {
  fillStyle: CanvasRenderingContext2D["fillStyle"]
  strokeStyle: CanvasRenderingContext2D["strokeStyle"]
  lineWidth: number
  font: string
  textAlign: CanvasTextAlign
}

export type CellDrawOptions = {
  painter: CellPainter
  layout: GridLayout
  controller: SheetController
  theme: GridTheme
  measure: Measure
}

/** Horizontal padding inside a cell. */
export const PAD = 3

type Rect = { x: number; y: number; w: number; h: number }

type TextDraw = {
  shown: Shown
  look: CellLook
  /** The cell (or merged area) the text belongs to. */
  rect: Rect
  /** Where the text may draw (wider than the cell when it runs into empty neighbours). */
  clip: [number, number]
  color: string
}

/** Typical descent below the baseline, as a share of the font's pixel size. */
const DESCENT = 0.26
/** Line height of wrapped text, as a share of the font's pixel size (15 pt rows for 11 pt text). */
const LINE = 1.36

export function drawCells({ painter: p, layout, controller, theme, measure }: CellDrawOptions) {
  const { rows, cols, clip } = layout
  if (!rows.length || !cols.length) return
  const sheet = controller.sheet
  const styles = controller.wb.styles
  const editing = controller.edit
  const darkTheme = !prefersDarkText(theme.background)
  /** Explicit colours as the theme shows them (dark mode turns their lightness around). */
  const paint = (color: string) => (darkTheme ? darkVariant(color) : color)
  // Conditional formatting: what the rules do to each cell, and the look it gives.
  const cfOn = sheet.conditional.length > 0 || controller.conditional.preview !== null
  const cfOf = (r: number, c: number): CFResult | null => (cfOn ? controller.cfAt(r, c) : null)
  // Tables: their style lies under the cells' own formats; conditional formatting goes over both.
  const tablesOn = sheet.tables.length > 0
  const lookAt = (r: number, c: number, cf: CFResult | null = cfOf(r, c)): CellLook => {
    const id = sheet.styleAt(r, c)
    const table = tablesOn ? tableAt(sheet, r, c) : null
    const tableLook = table ? tableCellLook(table, r, c, tableStyleDef(table.style, controller.wb.theme)) : null
    const cfLook = cf && (cf.style || cf.fill)
    if (!tableLook && !cfLook) return lookOf(styles, id)
    const merged = { ...withTableLook(styles.get(id), tableLook), ...cf?.style }
    if (cf?.fill) merged.fill = cf.fill
    return lookOf(controller.cfStyles, controller.cfStyles.intern(merged))
  }
  const bars: { rect: Rect; bar: NonNullable<CFResult["bar"]> }[] = []
  const icons: { rect: Rect; icon: NonNullable<CFResult["icon"]> }[] = []
  const sparkMap = sheet.sparklines.length ? sparkItems(controller.wb.sheets, sheet, controller.wb.version) : null
  const sparks: { rect: Rect; item: SparkItem }[] = []
  p.save()
  p.beginPath()
  p.rect(clip.x0, clip.y0, clip.x1 - clip.x0, clip.y1 - clip.y0)
  p.clip()

  const lastCol = cols[cols.length - 1]
  const gridRight = lastCol.start + lastCol.size
  const view: Range = { r0: rows[0].index, r1: rows[rows.length - 1].index, c0: cols[0].index, c1: lastCol.index }
  const merges = sheet.merges.size ? sheet.merges.overlapping(view) : []
  const mergeAt = (r: number, c: number) => (merges.length ? sheet.merges.at(r, c) : null)
  const rectOf = (range: Range): Rect => ({
    x: colX(sheet, layout, range.c0),
    y: rowY(sheet, layout, range.r0),
    w: sheet.cols.offset(range.c1 + 1) - sheet.cols.offset(range.c0),
    h: sheet.rows.offset(range.r1 + 1) - sheet.rows.offset(range.r0),
  })
  const isEditing = (r: number, c: number) => editing !== null && editing.row === r && editing.col === c

  const fills: { rect: Rect; color: string }[] = []
  const texts: TextDraw[] = []
  const bordered: {
    rect: Rect
    look: CellLook
    edges: { top: boolean; bottom: boolean; left: boolean; right: boolean }
  }[] = []
  /** Per row: column boundaries (the left edge of column c) hidden by text running across them. */
  const covered = new Map<number, Set<number>>()

  const textColor = (look: CellLook, shown: Shown) => {
    const fromFormat = formatColor(shown.color)
    if (fromFormat) return paint(fromFormat)
    if (look.color) return paint(look.color)
    // Automatic text stays readable on whatever fill the cell shows.
    if (look.fill) return prefersDarkText(paint(look.fill)) ? "#000000" : "#ffffff"
    return theme.text
  }

  const shownAt = (r: number, c: number, look: CellLook, w: number): Shown | null => {
    if (controller.showFormulas) {
      const input = sheet.cells.input(r, c)
      if (input?.startsWith("=")) return { text: input, numeric: false, align: "left" }
      if (sheet.cells.isSpill(r, c)) return null
    }
    const value: CellValue | undefined = sheet.cells.value(r, c)
    if (value === undefined) return null
    return shownText(value, look, w, PAD, measure)
  }

  /** How far text in (r, c) may run left and right into empty cells, marking the boundaries it crosses. */
  const overflow = (r: number, c: number, shown: Shown, look: CellLook, w: number) => {
    if (shown.numeric || look.wrap || look.shrink || look.rotation !== null || shown.text.includes("\n")) {
      return { left: 0, right: 0 }
    }
    const need0 = textWidth(shown.text, look.font, measure) + 2 * PAD + look.indent
    if (need0 <= w) return { left: 0, right: 0 }
    const need = shown.align === "center" ? (need0 - w) / 2 : need0 - w
    const marks = covered.get(r) ?? new Set<number>()
    covered.set(r, marks)
    let right = 0
    let left = 0
    const blocked = (cc: number) => sheet.cells.hasValue(r, cc) || isEditing(r, cc) || mergeAt(r, cc) !== null
    if (shown.align !== "right") {
      for (let cc = c + 1; right < need && cc < sheet.cols.count; cc++) {
        if (blocked(cc)) break
        marks.add(cc)
        right += sheet.cols.size(cc)
      }
    }
    if (shown.align !== "left") {
      for (let cc = c - 1; left < need && cc >= 0; cc--) {
        if (blocked(cc)) break
        marks.add(cc + 1)
        left += sheet.cols.size(cc)
      }
    }
    return { left, right }
  }

  const addText = (r: number, c: number, look: CellLook, cellRect: Rect, spill = true) => {
    let rect = cellRect
    if (look.hAlign === "centerContinuous" && spill) {
      // Center Across Selection: the text centres over its cell and the empty cells to the right that have the same
      // alignment.
      let w = rect.w
      for (let cc = c + 1; cc < sheet.cols.count; cc++) {
        if (sheet.cells.hasValue(r, cc) || mergeAt(r, cc)) break
        if (lookAt(r, cc).hAlign !== "centerContinuous") break
        w += sheet.cols.size(cc)
      }
      rect = { ...rect, w }
    }
    const shown = shownAt(r, c, look, rect.w)
    if (!shown) return
    const extent = spill ? overflow(r, c, shown, look, rect.w) : { left: 0, right: 0 }
    texts.push({
      shown,
      look,
      rect,
      clip: [rect.x - extent.left, rect.x + rect.w + extent.right],
      color: textColor(look, shown),
    })
  }

  const hasBorder = (look: CellLook) =>
    look.top || look.bottom || look.left || look.right || (look.diagonal && (look.diagonalUp || look.diagonalDown))

  // Cells in view (merged areas are drawn whole below).
  for (const row of rows) {
    for (const col of cols) {
      if (mergeAt(row.index, col.index)) continue
      const cf = cfOf(row.index, col.index)
      const look = lookAt(row.index, col.index, cf)
      const rect = { x: col.start, y: row.start, w: col.size, h: row.size }
      if (look.fill) fills.push({ rect, color: paint(look.fill) })
      if (hasBorder(look)) bordered.push({ rect, look, edges: { top: true, bottom: true, left: true, right: true } })
      if (cf?.bar) bars.push({ rect, bar: cf.bar })
      if (cf?.icon) icons.push({ rect, icon: cf.icon })
      const spark = sparkMap && sparkAt(sparkMap, row.index, col.index)
      if (spark) sparks.push({ rect, item: spark })
      if (!isEditing(row.index, col.index) && !cf?.hideValue) addText(row.index, col.index, look, rect)
    }
    // Text that starts left of the view (or right of it, when right-aligned) may run into it.
    const used = sheet.cells.usedColumns()
    for (let i = used.length - 1; i >= 0; i--) {
      const c = used[i]
      if (c >= view.c0) continue
      if (!sheet.cells.hasValue(row.index, c)) continue
      if (!mergeAt(row.index, c)) {
        const look = lookAt(row.index, c)
        const rect = {
          x: colX(sheet, layout, c),
          y: row.start,
          w: sheet.cols.size(c),
          h: row.size,
        }
        const shown = shownAt(row.index, c, look, rect.w)
        if (shown && shown.align !== "right") addText(row.index, c, look, rect)
      }
      break
    }
    for (const c of used) {
      if (c <= view.c1) continue
      if (!sheet.cells.hasValue(row.index, c)) continue
      if (!mergeAt(row.index, c)) {
        const look = lookAt(row.index, c)
        const rect = {
          x: colX(sheet, layout, c),
          y: row.start,
          w: sheet.cols.size(c),
          h: row.size,
        }
        const shown = shownAt(row.index, c, look, rect.w)
        if (shown && shown.align === "right") addText(row.index, c, look, rect)
      }
      break
    }
  }

  // Merged areas: one fill, one text (the top-left cell's), the outer borders of their edge cells.
  for (const m of merges) {
    const rect = rectOf(m)
    const look = lookAt(m.r0, m.c0)
    if (look.fill) fills.push({ rect, color: paint(look.fill) })
    const corner = (r: number, c: number) => lookOf(styles, sheet.styleAt(r, c))
    const edgeLook: CellLook = {
      ...look,
      right: corner(m.r0, m.c1).right ?? look.right,
      bottom: corner(m.r1, m.c0).bottom ?? look.bottom,
    }
    if (hasBorder(edgeLook))
      bordered.push({ rect, look: edgeLook, edges: { top: true, bottom: true, left: true, right: true } })
    if (!isEditing(m.r0, m.c0)) addText(m.r0, m.c0, look, rect, false)
    const spark = sparkMap && sparkAt(sparkMap, m.r0, m.c0)
    if (spark) sparks.push({ rect, item: spark })
  }

  // Gridlines (unless switched off in View), not inside merged areas or under text running across.
  if (controller.showGridlines) {
    p.strokeStyle = theme.gridline
    p.lineWidth = 1
    p.beginPath()
    for (const row of rows) {
      const y = Math.round(row.start + row.size) - 0.5
      if (!merges.length) {
        p.moveTo(clip.x0, y)
        p.lineTo(gridRight, y)
        continue
      }
      for (const col of cols) {
        const m = mergeAt(row.index, col.index)
        if (m && row.index < m.r1) continue
        p.moveTo(col.start - 0.5, y)
        p.lineTo(col.start + col.size - 0.5, y)
      }
    }
    for (const row of rows) {
      const hidden = covered.get(row.index)
      for (const col of cols) {
        if (hidden?.has(col.index)) continue
        const m = mergeAt(row.index, col.index)
        if (m && col.index > m.c0) continue
        const x = Math.round(col.start) - 0.5
        p.moveTo(x, row.start)
        p.lineTo(x, row.start + row.size)
      }
      if (!hidden?.has(lastCol.index + 1)) {
        const x = Math.round(gridRight) - 0.5
        p.moveTo(x, row.start)
        p.lineTo(x, row.start + row.size)
      }
    }
    p.stroke()
  }

  // Fills cover the gridlines around them, as in Excel.
  for (const f of fills) {
    p.fillStyle = f.color
    const x0 = Math.round(f.rect.x) - 1
    const y0 = Math.round(f.rect.y) - 1
    p.fillRect(x0, y0, Math.round(f.rect.x + f.rect.w) - x0, Math.round(f.rect.y + f.rect.h) - y0)
  }

  for (const { rect, bar } of bars) drawBar(p, rect, bar, paint)
  for (const { rect, icon } of icons) {
    const size = Math.max(6, Math.min(16, rect.h - 4))
    drawIcon(p, icon.set, icon.index, rect.x + 2, rect.y + rect.h - size - 2, size)
  }

  for (const { rect, item } of sparks) drawSparkline(p, rect, item, controller.wb.theme, paint)

  for (const t of texts) drawText(p, t, measure)

  drawBorders(p, bordered, theme.text, paint)
  p.restore()
}

/** A data bar: a gradient (or solid) bar from the cell's left, or either side of the axis for negative numbers. */
function drawBar(p: CellPainter, rect: Rect, bar: NonNullable<CFResult["bar"]>, paint: (c: string) => string) {
  const inner = rect.w - 4
  const x0 = Math.round(rect.x + 2 + inner * bar.start)
  const x1 = Math.round(rect.x + 2 + inner * bar.end)
  const y0 = Math.round(rect.y + 2)
  const h = Math.max(1, Math.round(rect.h - 4))
  const color = paint(bar.color)
  if (x1 - x0 >= 1) {
    if (bar.gradient) {
      // Excel's gradient: full colour at the bar's start fading almost to white at its end, with a border.
      const g = p.createLinearGradient(bar.negative ? x1 : x0, 0, bar.negative ? x0 : x1, 0)
      g.addColorStop(0, color)
      g.addColorStop(1, mix(color, "#ffffff", 0.9))
      p.fillStyle = g
      p.fillRect(x0, y0, x1 - x0, h)
      p.strokeStyle = color
      p.lineWidth = 1
      p.strokeRect(x0 + 0.5, y0 + 0.5, x1 - x0 - 1, h - 1)
    } else {
      p.fillStyle = color
      p.fillRect(x0, y0, x1 - x0, h)
    }
  }
  if (bar.axis !== null) {
    const x = Math.round(rect.x + 2 + inner * bar.axis) + 0.5
    p.strokeStyle = "#000000"
    p.lineWidth = 1
    p.setLineDash([1, 1])
    p.beginPath()
    p.moveTo(x, rect.y)
    p.lineTo(x, rect.y + rect.h)
    p.stroke()
    p.setLineDash([])
  }
}

function mix(a: string, b: string, t: number): string {
  const pa = parseInt(a.slice(1, 7), 16)
  const pb = parseInt(b.slice(1, 7), 16)
  const c = (s: number) => Math.round(((pa >> s) & 255) + (((pb >> s) & 255) - ((pa >> s) & 255)) * t)
  return "#" + ((c(16) << 16) | (c(8) << 8) | c(0)).toString(16).padStart(6, "0")
}

/** Breaks text into lines that fit `width` (at spaces when it can, else mid-word), keeping explicit line breaks. */
export function wrapLines(text: string, font: string, width: number, measure: Measure): string[] {
  const out: string[] = []
  for (const paragraph of text.split("\n")) {
    if (measure(paragraph, font) <= width || !paragraph) {
      out.push(paragraph)
      continue
    }
    let line = ""
    for (const word of paragraph.split(/(?<= )/)) {
      const candidate = line + word
      if (measure(candidate.trimEnd(), font) <= width) {
        line = candidate
        continue
      }
      if (line) out.push(line.trimEnd())
      // A word wider than the cell breaks between characters.
      let piece = ""
      for (const ch of word) {
        if (measure(piece + ch, font) > width && piece) {
          out.push(piece)
          piece = ""
        }
        piece += ch
      }
      line = piece
    }
    out.push(line.trimEnd())
  }
  return out
}

function drawText(p: CellPainter, t: TextDraw, measure: Measure) {
  const { shown, look, rect } = t
  const inner = Math.max(0, rect.w - 2 * PAD - look.indent)
  let font = look.font
  let px = look.fontPx
  let lines = [shown.text]
  if (look.wrap || shown.text.includes("\n")) {
    lines =
      look.wrap && !hasFormatMarks(shown.text) ? wrapLines(shown.text, font, inner, measure) : shown.text.split("\n")
  } else if (look.shrink && !shown.numeric) {
    const w = textWidth(shown.text, font, measure)
    if (w > inner && w > 0) {
      const shrunk = fontFor(look.style, inner / w)
      font = shrunk.font
      px = shrunk.px
    }
  }
  const lineHeight = Math.round(px * LINE)
  const descent = Math.max(3, Math.round(px * DESCENT) + 1)
  const blockHeight = lineHeight * (lines.length - 1)

  p.save()
  p.beginPath()
  p.rect(t.clip[0], rect.y, t.clip[1] - t.clip[0], rect.h)
  p.clip()
  p.font = font
  p.fillStyle = t.color

  if (look.rotation !== null) {
    drawRotated(p, t, font, px, measure)
    p.restore()
    return
  }

  let baseline: number
  if (look.vAlign === "top") baseline = rect.y + 2 + px * (1 - DESCENT)
  else if (look.vAlign === "center" || look.vAlign === "justify" || look.vAlign === "distributed")
    baseline = rect.y + rect.h / 2 - blockHeight / 2 + px * 0.33
  else baseline = rect.y + rect.h - descent - blockHeight

  lines.forEach((text, i) => {
    const y = baseline + i * lineHeight
    const marks = hasFormatMarks(text)
    const laid = marks ? layoutMarks(text, font, measure, inner) : null
    const w = laid ? laid.width : measure(text, font)
    let x: number
    if (shown.align === "right") x = rect.x + rect.w - PAD - look.indent - w
    else if (shown.align === "center") x = rect.x + rect.w / 2 - w / 2
    else x = rect.x + PAD + look.indent
    p.textAlign = "left"
    if (laid) for (const run of laid.runs) p.fillText(run.text, x + run.x, y)
    else if (look.hAlign === "fill" && text && !shown.numeric) {
      // Fill: the text repeated across the cell.
      const n = Math.max(1, Math.floor(inner / Math.max(1, w)))
      p.fillText(text.repeat(n), x, y)
    } else p.fillText(text, x, y)
    decorate(p, look, x, y, w, px, t.color)
  })
  p.restore()
}

/** Underline and strikethrough, as lines under and through the text. */
function decorate(p: CellPainter, look: CellLook, x: number, y: number, w: number, px: number, color: string) {
  if (!look.underline && !look.strike) return
  p.fillStyle = color
  const thickness = Math.max(1, Math.round(px / 14))
  if (look.underline === "single" || look.underline === "singleAccounting") {
    p.fillRect(x, Math.round(y + px * 0.12), w, thickness)
  } else if (look.underline === "double" || look.underline === "doubleAccounting") {
    p.fillRect(x, Math.round(y + px * 0.08), w, thickness)
    p.fillRect(x, Math.round(y + px * 0.08) + 2 * thickness, w, thickness)
  }
  if (look.strike) p.fillRect(x, Math.round(y - px * 0.3), w, thickness)
}

/** Rotated text (Format Cells › Alignment › Orientation), anchored as Excel does for bottom-aligned cells. */
function drawRotated(p: CellPainter, t: TextDraw, font: string, px: number, measure: Measure) {
  const { rect, look, shown } = t
  const text = shown.text.replace(/\n/g, " ")
  if (look.rotation === "stacked") {
    // Letters one under the other, centred.
    const step = Math.round(px * 1.2)
    p.textAlign = "center"
    let y = rect.y + 2 + px
    for (const ch of text) {
      p.fillText(ch, rect.x + rect.w / 2, y)
      y += step
    }
    return
  }
  const angle = ((look.rotation as number) * Math.PI) / 180
  const w = measure(text, font)
  p.translate(rect.x + rect.w / 2, rect.y + rect.h / 2)
  p.rotate(-angle)
  p.textAlign = "left"
  p.fillText(text, -w / 2, px * 0.33)
}

type Segment = { x0: number; y0: number; x1: number; y1: number; border: Border; weight: number }

const WIDTH: Record<Border["style"], number> = {
  thin: 1,
  hair: 1,
  dotted: 1,
  dashed: 1,
  dashDot: 1,
  dashDotDot: 1,
  medium: 2,
  mediumDashed: 2,
  mediumDashDot: 2,
  mediumDashDotDot: 2,
  slantDashDot: 2,
  thick: 3,
  double: 3,
}

const DASH: Partial<Record<Border["style"], number[]>> = {
  hair: [1, 1],
  dotted: [1, 2],
  dashed: [3, 1],
  dashDot: [6, 2, 2, 2],
  dashDotDot: [6, 2, 2, 2, 2, 2],
  mediumDashed: [6, 2],
  mediumDashDot: [6, 2, 2, 2],
  mediumDashDotDot: [6, 2, 2, 2, 2, 2],
  slantDashDot: [8, 2, 3, 2],
}

/** Excel draws the heavier of two borders that meet on an edge; drawing lightest first gives the same result. */
const weightOf = (b: Border) => WIDTH[b.style] * 10 + (b.style === "double" ? 5 : 0) + (DASH[b.style] ? 0 : 1)

function drawBorders(
  p: CellPainter,
  cells: { rect: Rect; look: CellLook; edges: { top: boolean; bottom: boolean; left: boolean; right: boolean } }[],
  automatic: string,
  paint: (color: string) => string,
) {
  if (!cells.length) return
  const segments: Segment[] = []
  const diagonals: Segment[] = []
  for (const { rect, look, edges } of cells) {
    const x0 = Math.round(rect.x)
    const x1 = Math.round(rect.x + rect.w)
    const y0 = Math.round(rect.y)
    const y1 = Math.round(rect.y + rect.h)
    const add = (b: Border | undefined, s: Omit<Segment, "border" | "weight">) => {
      if (b) segments.push({ ...s, border: b, weight: weightOf(b) })
    }
    if (edges.top) add(look.top, { x0, y0, x1, y1: y0 })
    if (edges.bottom) add(look.bottom, { x0, y0: y1, x1, y1 })
    if (edges.left) add(look.left, { x0, y0, x1: x0, y1 })
    if (edges.right) add(look.right, { x0: x1, y0, x1, y1 })
    if (look.diagonal) {
      if (look.diagonalDown) diagonals.push({ x0, y0, x1, y1, border: look.diagonal, weight: 0 })
      if (look.diagonalUp) diagonals.push({ x0, y0: y1, x1, y1: y0, border: look.diagonal, weight: 0 })
    }
  }
  segments.sort((a, b) => a.weight - b.weight)
  for (const s of segments) strokeEdge(p, s, s.border.color ? paint(s.border.color) : automatic)
  for (const d of diagonals) {
    p.strokeStyle = d.border.color ? paint(d.border.color) : automatic
    p.lineWidth = WIDTH[d.border.style]
    p.setLineDash(DASH[d.border.style] ?? [])
    p.beginPath()
    p.moveTo(d.x0, d.y0)
    p.lineTo(d.x1, d.y1)
    p.stroke()
  }
  p.setLineDash([])
}

function strokeEdge(p: CellPainter, s: Segment, color: string) {
  const w = WIDTH[s.border.style]
  p.strokeStyle = color
  p.setLineDash(DASH[s.border.style] ?? [])
  const horizontal = s.y0 === s.y1
  // On the gridline's pixel: odd widths centre on it, even widths take it and the one after.
  const centre = (v: number) => v - 0.5 + (w % 2 === 0 ? 0.5 : 0)
  const reach = Math.floor(w / 2) + 1
  p.beginPath()
  if (s.border.style === "double") {
    p.lineWidth = 1
    for (const offset of [-1, 1]) {
      if (horizontal) {
        p.moveTo(s.x0 - reach, centre(s.y0) + offset)
        p.lineTo(s.x1 + reach - 1, centre(s.y0) + offset)
      } else {
        p.moveTo(centre(s.x0) + offset, s.y0 - reach)
        p.lineTo(centre(s.x0) + offset, s.y1 + reach - 1)
      }
    }
  } else {
    p.lineWidth = w
    if (horizontal) {
      p.moveTo(s.x0 - reach, centre(s.y0))
      p.lineTo(s.x1 + reach - 1, centre(s.y0))
    } else {
      p.moveTo(centre(s.x0), s.y0 - reach)
      p.lineTo(centre(s.x0), s.y1 + reach - 1)
    }
  }
  p.stroke()
}

/** The lines of a row, for callers that need cell geometry (exported for tests). */
export type { Line }
