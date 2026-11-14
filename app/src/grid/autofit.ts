import { formatValue, MAX_ROWS } from "@cells/core"
import type { SheetController } from "@/sheet/controller"

const PAD = 3
const LINE_HEIGHT = 18

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
): number | null {
  let widest = 0
  controller.sheet.cells.forEach(
    { r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col },
    (_row, _col, data) => {
      const style = controller.wb.styles.get(data.style ?? 0)
      const text = formatValue(data.value, style.numberFormat, 30) ?? ""
      for (const line of text.split("\n")) widest = Math.max(widest, measure(line))
    },
    true,
  )
  return widest === 0 ? null : Math.ceil(widest + 2 * PAD + 2)
}

/** Double-click on a row border: tall enough for the row's line breaks, or the default. */
export function autofitRowHeight(controller: SheetController, row: number): number | null {
  let lines = 1
  for (const col of controller.sheet.cells.usedColumns()) {
    const value = controller.sheet.cells.value(row, col)
    if (typeof value === "string") lines = Math.max(lines, value.split("\n").length)
  }
  return lines <= 1 ? null : lines * LINE_HEIGHT + 2
}
