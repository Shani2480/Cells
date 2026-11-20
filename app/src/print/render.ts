import { renderChart } from "@cells/chart"
import {
  columnName,
  isChart,
  slicerConnected,
  type CellAddress,
  type HeaderFooterText,
  type Pagination,
  type PrintPage,
  type Sheet,
} from "@cells/core"
import { canvasMeasure, objectAt } from "@/drawing/render"
import { chartBox, type ChartBox } from "@/sheet/charts"
import type { SheetController } from "@/sheet/controller"
import { drawCells } from "@/grid/draw-cells"
import { setBodyFont } from "@/grid/cell-look"
import { colX, rowY, type GridLayout, type Line } from "@/grid/layout"
import { CELL_FONT, CELL_FONT_FAMILY, LIGHT_THEME, type GridTheme } from "@/grid/theme"

/**
 * One printed page drawn on a canvas: the cells through the grid's own cell drawing (so they print as they look),
 * titles repeated, headings, charts, and the header and footer. The canvas is the paper; `pxPerPt` sets the
 * resolution (preview, print and PDF draw the same page at different sizes).
 */

const PX_PER_PT = 96 / 72

/** Paper is white, and gridlines print black, as Excel prints them. */
const PAPER_THEME: GridTheme = { ...LIGHT_THEME, gridline: "#000000" }

/** What the header and footer codes read: &F, &Z, &D, &T and the page count. */
export type PrintContext = {
  fileName: string
  /** The file's folder (&Z), or "" for a workbook not saved yet. */
  folder: string
  now: Date
  /** The number of pages printed (&N). */
  total: number
}

/** The cell drawing's view of the controller for a sheet on paper: no editing, no selection, its own gridlines. */
function paperController(c: SheetController, sheet: Sheet, gridlines: boolean): SheetController {
  const view = Object.create(c) as SheetController
  Object.defineProperties(view, {
    sheet: { value: sheet },
    edit: { value: null },
    showGridlines: { value: gridlines },
    showFormulas: { value: false },
    conditional: { value: { ...c.conditional, preview: null } },
  })
  return view
}

/** The page's cell area as the grid lays out frozen panes: titles are the frozen part, the page's cells the rest. */
function pageLayout(sheet: Sheet, page: PrintPage, headingsW: number, headingsH: number): GridLayout {
  const lines = (axis: typeof sheet.rows, a: number, b: number, start: number) => {
    const out: Line[] = []
    let at = start
    for (let i = a; i <= b; i++) {
      const size = axis.size(i)
      if (size <= 0) continue
      out.push({ index: i, start: at, size })
      at += size
    }
    return { out, end: at }
  }
  const tRows = page.titleRows ? lines(sheet.rows, page.titleRows[0], page.titleRows[1], headingsH) : null
  const tCols = page.titleCols ? lines(sheet.cols, page.titleCols[0], page.titleCols[1], headingsW) : null
  const originY = tRows?.end ?? headingsH
  const originX = tCols?.end ?? headingsW
  const rows = lines(sheet.rows, page.rows[0], page.rows[1], originY)
  const cols = lines(sheet.cols, page.cols[0], page.cols[1], originX)
  const width = cols.end
  const height = rows.end
  const frozen =
    tRows || tCols
      ? {
          row: page.titleRows ? page.titleRows[1] + 1 : 0,
          col: page.titleCols ? page.titleCols[1] + 1 : 0,
          scrollY: page.titleRows ? sheet.rows.offset(page.titleRows[0]) : 0,
          scrollX: page.titleCols ? sheet.cols.offset(page.titleCols[0]) : 0,
        }
      : null
  const base = { width, height, rowHeaderWidth: headingsW, headerHeight: headingsH, frozen, panes: [] }
  const scrollX = sheet.cols.offset(page.cols[0])
  const scrollY = sheet.rows.offset(page.rows[0])
  const pane = (r: Line[], c: Line[], top: boolean, left: boolean): GridLayout => ({
    ...base,
    rows: r,
    cols: c,
    scrollX: left ? frozen!.scrollX : scrollX,
    scrollY: top ? frozen!.scrollY : scrollY,
    originX: left ? headingsW : originX,
    originY: top ? headingsH : originY,
    clip: {
      x0: left ? headingsW : originX,
      y0: top ? headingsH : originY,
      x1: left ? originX : width,
      y1: top ? originY : height,
    },
  })
  const panes = [pane(rows.out, cols.out, false, false)]
  if (tRows) panes.push(pane(tRows.out, cols.out, true, false))
  if (tCols) panes.push(pane(rows.out, tCols.out, false, true))
  if (tRows && tCols) panes.push(pane(tRows.out, tCols.out, true, true))
  return {
    ...base,
    rows: [...(tRows?.out ?? []), ...rows.out],
    cols: [...(tCols?.out ?? []), ...cols.out],
    scrollX,
    scrollY,
    originX,
    originY,
    clip: { x0: headingsW, y0: headingsH, x1: width, y1: height },
    panes,
  }
}

/**
 * The sheet's charts, shapes and pictures as pictures, loaded once per drawing pass (SVG to an image the canvas can
 * draw), in stacking order. Hidden objects don't print. Each image covers the object's turned box.
 */
async function objectImages(
  c: SheetController,
  sheet: Sheet,
): Promise<{ img: HTMLImageElement; box: ChartBox; from: CellAddress }[]> {
  const out = await Promise.all(
    sheet.objects.map(async (o) => {
      if (!isChart(o) && o.hidden) return null
      if (!isChart(o) && (o.kind === "slicer" || o.kind === "timeline") && !slicerConnected(c.wb.sheets, c.wb.pivotCaches, o)) return null
      const box = chartBox(sheet, o.anchor)
      // Room around the box for a rotated object's corners and a shadow.
      const pad = Math.ceil(Math.max(box.w, box.h) / 2) + 8
      let body: string
      try {
        body = isChart(o)
          ? renderChart(c.charts.liveSpec(o), box.w, box.h)
          : objectAt(o, { x: 0, y: 0, w: box.w, h: box.h }, c.wb.theme, canvasMeasure)
      } catch {
        return null
      }
      const w = box.w + 2 * pad
      const h = box.h + 2 * pad
      const markup = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="${-pad} ${-pad} ${w} ${h}">${body}</svg>`
      const img = new Image()
      img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(markup)
      try {
        await img.decode()
      } catch {
        // An object that can't be drawn prints as nothing.
        return null
      }
      return { img, box: { x: box.x - pad, y: box.y - pad, w, h }, from: o.anchor.from }
    }),
  )
  return out.filter((x) => x !== null)
}

/** Header and footer codes: the text each section prints (formatting codes left out). */
export function headerFooterLines(text: string, page: PrintPage, sheet: Sheet, ctx: PrintContext): string[] {
  let out = ""
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch !== "&" || i + 1 >= text.length) {
      out += ch
      continue
    }
    const code = text[++i]
    switch (code) {
      case "&":
        out += "&"
        break
      case "P":
      case "p": {
        // &P+n and &P-n add to the page number.
        const m = /^([+-])(\d+)/.exec(text.slice(i + 1))
        let n = page.number
        if (m) {
          n += (m[1] === "+" ? 1 : -1) * Number(m[2])
          i += m[0].length
        }
        out += String(n)
        break
      }
      case "N":
      case "n":
        out += String(ctx.total)
        break
      case "D":
      case "d":
        out += ctx.now.toLocaleDateString()
        break
      case "T":
      case "t":
        out += ctx.now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
        break
      case "F":
      case "f":
        out += ctx.fileName
        break
      case "Z":
      case "z":
        out += ctx.folder
        break
      case "A":
      case "a":
        out += sheet.name
        break
      case '"': {
        // A font: &"Name,Style".
        const end = text.indexOf('"', i + 1)
        i = end < 0 ? text.length : end
        break
      }
      case "K":
        i += 6
        break
      default:
        // A size (&12) or a style switch (&B, &I, &U, &S, &E, &X, &Y, &G picture).
        if (/\d/.test(code)) while (/\d/.test(text[i + 1] ?? "")) i++
    }
  }
  return out.split(/\r?\n/)
}

/** Draws one page on a canvas `pxPerPt` pixels per point big (the canvas must be paper-sized). */
export async function drawPage(
  canvas: HTMLCanvasElement,
  c: SheetController,
  sheet: Sheet,
  pagination: Pagination,
  page: PrintPage,
  ctx: PrintContext,
  pxPerPt: number,
) {
  const p = canvas.getContext("2d")!
  const setup = sheet.pageSetup
  const { paper, body, scale } = pagination
  p.setTransform(pxPerPt, 0, 0, pxPerPt, 0, 0)
  p.fillStyle = "#ffffff"
  p.fillRect(0, 0, paper.width, paper.height)

  // The cells, in sheet pixels scaled to points and the page's scale.
  const headingsW = setup.headings ? Math.max(20, String(page.rows[1] + 1).length * 6 + 9) * PX_PER_PT : 0
  const headingsH = setup.headings ? 20 : 0
  const layout = pageLayout(sheet, page, headingsW, headingsH)
  const k = scale / PX_PER_PT
  const contentW = layout.width * k
  const contentH = layout.height * k
  const x0 = body.x + (setup.centerHorizontally ? Math.max(0, (body.width - contentW) / 2) : 0)
  const y0 = body.y + (setup.centerVertically ? Math.max(0, (body.height - contentH) / 2) : 0)
  const unit = pxPerPt * k
  p.setTransform(unit, 0, 0, unit, x0 * pxPerPt, y0 * pxPerPt)

  const measure = (text: string, font: string = CELL_FONT) => {
    const before = p.font
    p.font = font
    const w = p.measureText(text).width
    p.font = before
    return w
  }
  const view = paperController(c, sheet, setup.gridlines)
  p.font = CELL_FONT
  p.textBaseline = "alphabetic"
  setBodyFont(view.wb.theme.fonts.minor)
  for (const pane of layout.panes)
    drawCells({ painter: p, layout: pane, controller: view, theme: PAPER_THEME, measure })
  if (setup.headings) drawHeadings(p, layout)

  // Charts, shapes and pictures on the page, clipped to its cells.
  if (sheet.objects.length) {
    const images = await objectImages(c, sheet)
    p.save()
    p.beginPath()
    p.rect(layout.originX, layout.originY, layout.width - layout.originX, layout.height - layout.originY)
    p.clip()
    for (const { img, box, from } of images) {
      // Objects in the titles' rows or columns aren't repeated with them, as in Excel.
      const x = colX(sheet, { ...layout, frozen: null }, from.col) + box.x - sheet.cols.offset(from.col)
      const y = rowY(sheet, { ...layout, frozen: null }, from.row) + box.y - sheet.rows.offset(from.row)
      p.drawImage(img, x, y, box.w, box.h)
    }
    p.restore()
  }

  // Header and footer, in the margins, at 11 pt in the body font, scaled with the sheet (Excel's "Scale with document").
  p.setTransform(pxPerPt, 0, 0, pxPerPt, 0, 0)
  drawHeaderFooter(p, setup.header, sheet, page, ctx, pagination, "header")
  drawHeaderFooter(p, setup.footer, sheet, page, ctx, pagination, "footer")
}

function drawHeadings(p: CanvasRenderingContext2D, layout: GridLayout) {
  p.save()
  p.font = CELL_FONT
  p.fillStyle = "#000000"
  p.strokeStyle = PAPER_THEME.gridline
  p.lineWidth = 1
  p.textAlign = "center"
  p.textBaseline = "middle"
  const h = layout.headerHeight
  const w = layout.rowHeaderWidth
  p.beginPath()
  for (const col of layout.cols) {
    p.fillText(columnName(col.index), col.start + col.size / 2, h / 2)
    p.rect(col.start + 0.5, 0.5, col.size, h)
  }
  for (const row of layout.rows) {
    p.fillText(String(row.index + 1), w / 2, row.start + row.size / 2)
    p.rect(0.5, row.start + 0.5, w, row.size)
  }
  p.rect(0.5, 0.5, w, h)
  p.stroke()
  p.restore()
}

function drawHeaderFooter(
  p: CanvasRenderingContext2D,
  hf: HeaderFooterText,
  sheet: Sheet,
  page: PrintPage,
  ctx: PrintContext,
  pagination: Pagination,
  where: "header" | "footer",
) {
  const m = sheet.pageSetup.margins
  const { paper } = pagination
  // The canvas is already in points here, so the font's size is given in canvas units (px), not pt.
  const size = 11 * pagination.scale
  const lineHeight = size * 1.2
  p.font = `${size}px ${CELL_FONT_FAMILY}`
  p.fillStyle = "#000000"
  p.textBaseline = "alphabetic"
  const left = m.left * 72
  const right = paper.width - m.right * 72
  for (const [section, align, x] of [
    ["left", "left", left],
    ["center", "center", (left + right) / 2],
    ["right", "right", right],
  ] as const) {
    const text = hf[section]
    if (!text) continue
    const lines = headerFooterLines(text, page, sheet, ctx)
    p.textAlign = align
    // A header hangs from its margin; a footer stands on its own.
    const top =
      where === "header" ? m.header * 72 + size : paper.height - m.footer * 72 - (lines.length - 1) * lineHeight
    lines.forEach((line, i) => p.fillText(line, x, top + i * lineHeight))
  }
}
