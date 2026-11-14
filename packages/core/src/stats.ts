import type { Range } from "./address"
import type { Sheet } from "./sheet"

/** The status bar's figures for a selection: Count is COUNTA; Sum and Average cover numbers only. */
export type SelectionStats = {
  count: number
  numericCount: number
  sum: number
  average: number | null
  min: number | null
  max: number | null
}

/**
 * Figures over the selected cells, each counted once even where ranges overlap. Hidden rows and columns still count
 * (as in Excel, which drops them only for filtered rows).
 */
export function selectionStats(sheet: Sheet, ranges: readonly Range[]): SelectionStats {
  let count = 0
  let numericCount = 0
  let sum = 0
  let min = Infinity
  let max = -Infinity
  const seen = ranges.length > 1 ? new Set<number>() : null
  for (const range of ranges) {
    sheet.cells.forEach(
      range,
      (row, col, data) => {
        if (seen) {
          const key = row * 16_384 + col
          if (seen.has(key)) return
          seen.add(key)
        }
        count++
        if (typeof data.value === "number") {
          numericCount++
          sum += data.value
          min = Math.min(min, data.value)
          max = Math.max(max, data.value)
        }
      },
      true,
    )
  }
  return {
    count,
    numericCount,
    sum,
    average: numericCount ? sum / numericCount : null,
    min: numericCount ? min : null,
    max: numericCount ? max : null,
  }
}
