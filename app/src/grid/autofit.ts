import { DEFAULT_ROW_HEIGHT, formatValue, formatWithCode, MAX_COLS, MAX_ROWS } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { lookOf, shownText, type Measure } from "./cell-look"
import { wrapLines } from "./draw-cells"

const PAD = 3

/** Excel's column width in characters of the default font's widest digit (7 px): 64 px is 8.43. */
export function columnChars(px: number) {
  const chars = px < 12 ? px / 12 : (px - 5) / 7
  return Math.round(chars * 100) / 100
}

export function charsToPixels(chars: number) {
  return chars < 1 ? Math.round(chars * 12) : Math.round(chars * 7 + 5)
}

export function columnWidthLabel(px: number) {
  return `${columnChars(px).toFixed(2)} (${px} pixels)`
}

/** Row heights are in points: 20 px is 15 pt at 96 dpi. */
export function rowPoints(px: number) {
  return Math.round(px * 0.75 * 100) / 100
}

export function pointsToPixels(points: number) {
  return Math.round(points / 0.75)
}

export function rowHeightLabel(px: number) {
  return `${rowPoints(px).toFixed(2)} (${px} pixels)`
}

/** Double-click on a column border: wide enough for the column's longest text, or the default when empty. */
export function autofitColumnWidth(
  controller: SheetController,
  col: number,
  measure: (text: string) => number,
  rows: { r0: number; r1: number } = { r0: 0, r1: MAX_ROWS - 1 },
): number | null {
  let widest = 0
  controller.sheet.cells.forEach(
    { r0: rows.r0, r1: rows.r1, c0: col, c1: col },
    (_row, _col, data) => {
      const style = controller.wb.styles.get(data.style ?? 0)
      // Numbers as the cell shows them: with their format (dates as dates, as the grid draws them), or General up to
      // 11 characters (0.05, not 0.05000000000000000277).
      const code = style.numberFormat
      const text =
        typeof data.value === "number" && code && code !== "General"
          ? formatWithCode(data.value, code, { cell: true }).text
          : (formatValue(data.value, code, 11) ?? "")
      for (const line of text.split("\n")) widest = Math.max(widest, measure(line))
    },
    true,
  )
  return widest === 0 ? null : Math.ceil(widest + 2 * PAD + 2)
}

/** A font's line height as the cells draw it (11 pt: 20 px, the default row). */
export const lineHeightOf = (fontPx: number) => Math.round(fontPx * 1.36)

/**
 * The height a row's content needs, as Excel fits rows by itself: the tallest cell's lines (wrapped text, line
 * breaks) at its font's line height. Merged and rotated cells don't count (Excel skips them). Null: the default.
 */
export function fittedRowHeight(controller: SheetController, row: number, measure: Measure): number | null {
  const sheet = controller.sheet
  const styles = controller.wb.styles
  let tallest = 0
  const fit = (col: number, styleId: number) => {
    if (sheet.merges.at(row, col)) return
    const look = lookOf(styles, styleId)
    if (look.rotation !== null) return
    let lines = 1
    const value = sheet.cells.value(row, col)
    if (value !== undefined && value !== null && value !== "") {
      const width = sheet.cols.size(col)
      const text = shownText(value, look, width, PAD, measure).text
      lines = look.wrap
        ? wrapLines(text, look.font, Math.max(1, width - 2 * PAD - look.indent), measure).length
        : text.split("\n").length
    }
    tallest = Math.max(tallest, lines * lineHeightOf(look.fontPx))
  }
  sheet.cells.forEach({ r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }, (_r, col) => fit(col, sheet.styleAt(row, col)))
  // A row format with a bigger font heightens the row even where it is empty.
  const rowStyle = sheet.rowStyles.get(row)
  if (rowStyle) tallest = Math.max(tallest, lineHeightOf(lookOf(styles, rowStyle).fontPx))
  return tallest > DEFAULT_ROW_HEIGHT ? tallest : null
}
