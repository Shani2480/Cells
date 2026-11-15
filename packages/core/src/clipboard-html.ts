/**
 * The HTML a copied range carries, as Excel's: a table with each cell's font, colours, fill, alignment and borders
 * inline, the column widths and row heights, and merged cells as colspan/rowspan. Word, LibreOffice, mail and
 * Inkspire paste it as a formatted table.
 *
 * The table also says where it came from (`data-cells`, JSON): the workbook's path, the sheet, the range and a
 * cells:// link, so Inkspire can paste it as a linked table (docs/inkspire-clipboard.md).
 */
import { rangeName, type Range } from "./address"
import type { Sheet } from "./sheet"
import type { Border, CellStyle, StyleTable } from "./styles"

/** Where a copied range lives; `file` is null for a workbook never saved. */
export type CopySource = { file: string | null; sheet: string; range: Range; link: string | null }

/** What `data-cells` holds on a copied table. */
export type CopiedRangeInfo = {
  app: "cells"
  version: 1
  file: string | null
  sheet: string
  /** "A1:D10". */
  range: string
  link: string | null
}

const escapeHtml = (text: string) =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

const PX_PER_PT = 96 / 72

function borderCss(b: Border | undefined): string | null {
  if (!b) return null
  const color = b.color ?? "#000000"
  switch (b.style) {
    case "medium":
    case "mediumDashed":
    case "mediumDashDot":
    case "mediumDashDotDot":
    case "slantDashDot":
      return `2px ${b.style === "medium" ? "solid" : "dashed"} ${color}`
    case "thick":
      return `3px solid ${color}`
    case "double":
      return `3px double ${color}`
    case "dotted":
    case "hair":
      return `1px dotted ${color}`
    case "dashed":
    case "dashDot":
    case "dashDotDot":
      return `1px dashed ${color}`
    default:
      return `1px solid ${color}`
  }
}

/** A cell's look as inline CSS. `numeric` right-aligns General, as the grid does. */
export function cellCss(style: CellStyle, numeric: boolean): string {
  const css: string[] = []
  css.push(`font-family:'${(style.fontName ?? "Aptos Narrow").replace(/'/g, "")}',sans-serif`)
  css.push(`font-size:${style.fontSize ?? 11}pt`)
  if (style.bold) css.push("font-weight:700")
  if (style.italic) css.push("font-style:italic")
  const lines = [style.underline ? "underline" : "", style.strike ? "line-through" : ""].filter(Boolean)
  if (lines.length) css.push(`text-decoration:${lines.join(" ")}`)
  if (style.underline === "double" || style.underline === "doubleAccounting") css.push("text-decoration-style:double")
  if (style.fontColor) css.push(`color:${style.fontColor}`)
  if (style.fill) css.push(`background:${style.fill}`)
  const h = style.hAlign ?? "general"
  const align =
    h === "general"
      ? numeric
        ? "right"
        : null
      : h === "centerContinuous" || h === "distributed"
        ? "center"
        : h === "fill"
          ? "left"
          : h
  if (align) css.push(`text-align:${align}`)
  const v = style.vAlign ?? "bottom"
  css.push(`vertical-align:${v === "center" ? "middle" : v === "top" ? "top" : "bottom"}`)
  if (style.indent) css.push(`padding-left:${style.indent * 9}pt`)
  css.push(style.wrap ? "white-space:normal" : "white-space:nowrap")
  for (const [side, b] of [
    ["top", style.borderTop],
    ["bottom", style.borderBottom],
    ["left", style.borderLeft],
    ["right", style.borderRight],
  ] as const) {
    const value = borderCss(b)
    if (value) css.push(`border-${side}:${value}`)
  }
  return css.join(";")
}

/**
 * The copied range as HTML. `texts[r][c]` is what each cell shows (its formatted value); hidden rows and columns
 * are left out, as Excel leaves them out of what it pastes elsewhere.
 */
export function rangeHtml(
  sheet: Sheet,
  styles: StyleTable,
  range: Range,
  texts: readonly (readonly string[])[],
  source?: CopySource,
): string {
  const cols: number[] = []
  for (let c = range.c0; c <= range.c1; c++) if (sheet.cols.size(c) > 0) cols.push(c)
  const rows: number[] = []
  for (let r = range.r0; r <= range.r1; r++) if (sheet.rows.size(r) > 0) rows.push(r)
  const info: CopiedRangeInfo | null = source
    ? {
        app: "cells",
        version: 1,
        file: source.file,
        sheet: source.sheet,
        range: rangeName(source.range),
        link: source.link,
      }
    : null
  const covered = new Set<string>()
  let html = `<meta charset="utf-8"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse;table-layout:fixed"`
  if (info) html += ` data-cells="${escapeHtml(JSON.stringify(info))}"`
  html += ">"
  html += cols.map((c) => `<col width="${Math.round(sheet.cols.size(c))}">`).join("")
  html += "<tbody>"
  for (const r of rows) {
    html += `<tr style="height:${(sheet.rows.size(r) / PX_PER_PT).toFixed(1)}pt">`
    for (const c of cols) {
      if (covered.has(`${r},${c}`)) continue
      const merge = sheet.merges.at(r, c)
      let span = ""
      if (merge && merge.r0 === r && merge.c0 === c) {
        const rowSpan = rows.filter((x) => x >= merge.r0 && x <= Math.min(merge.r1, range.r1)).length
        const colSpan = cols.filter((x) => x >= merge.c0 && x <= Math.min(merge.c1, range.c1)).length
        for (const rr of rows)
          for (const cc of cols)
            if (rr >= merge.r0 && rr <= merge.r1 && cc >= merge.c0 && cc <= merge.c1 && !(rr === r && cc === c))
              covered.add(`${rr},${cc}`)
        if (rowSpan > 1) span += ` rowspan="${rowSpan}"`
        if (colSpan > 1) span += ` colspan="${colSpan}"`
      }
      const data = sheet.cells.get(r, c)
      const style = styles.get(sheet.styleAt(r, c))
      const text = texts[r - range.r0]?.[c - range.c0] ?? ""
      const numeric = typeof data?.value === "number"
      const body = escapeHtml(text).replace(/\r?\n/g, "<br>")
      html += `<td${span} style="${escapeHtml(cellCss(style, numeric))}">${body}</td>`
    }
    html += "</tr>"
  }
  return html + "</tbody></table>"
}

/** The `data-cells` note of copied HTML, when it came from Cells. */
export function copiedRangeInfo(html: string): CopiedRangeInfo | null {
  const m = /data-cells="([^"]*)"/.exec(html)
  if (!m) return null
  try {
    const json = m[1]
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
    const info = JSON.parse(json) as CopiedRangeInfo
    return info && info.app === "cells" ? info : null
  } catch {
    return null
  }
}
