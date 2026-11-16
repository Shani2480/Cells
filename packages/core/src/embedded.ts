import { columnName } from "./address"
import { contentForInput } from "./edit"
import { displayText } from "./format"
import { parseFormula } from "./formula"
import type { CellWrite } from "./ops"
import { regional, withRegional, type Regional } from "./regional"
import type { CellValue, ErrorCode } from "./values"
import { Workbook } from "./workbook"

/**
 * The formula engine for small grids that live outside a workbook: Inkspire's Spreadsheet block and its formula fields
 * (docs/shared-packages.md). Cells are typed as a person types them into Cells (numbers, dates, percentages, text,
 * formulas with Excel's functions) and come back calculated, with the text a cell shows.
 */

export type GridCell = {
  /** What the cell holds after calculation; undefined when it is empty. */
  value: number | string | boolean | undefined
  /** Excel's error code when the cell holds an error ("#DIV/0!", "#NAME?"…). */
  error?: ErrorCode
  /** What the cell shows, formatted the way the person typed it ("10%", "1,000", "3/14/2026"). */
  text: string
}

export type GridOptions = {
  /** The regional formats used to read and show numbers and dates (default: Cells' current ones, US English). */
  regional?: Regional
  /** The sheet's name, for formulas that name it. */
  sheetName?: string
}

/** Calculates a grid of typed cells (`inputs[row][col]`, A1 at [0][0]); the result has the same shape. */
export function calculateGrid(inputs: readonly (readonly string[])[], options: GridOptions = {}): GridCell[][] {
  return withRegional(options.regional ?? regional(), () => calculate(inputs, options.sheetName))
}

function calculate(inputs: readonly (readonly string[])[], sheetName = "Sheet1"): GridCell[][] {
  const wb = new Workbook([sheetName], 0)
  const sheet = wb.sheets[0]!
  const writes: CellWrite[] = []
  inputs.forEach((row, ri) =>
    row.forEach((text, ci) => {
      if (!text) return
      const data = contentForInput(wb.styles, undefined, text)
      if (data) writes.push({ row: ri, col: ci, data })
    }),
  )
  wb.transact([{ kind: "setCells", sheet: 0, cells: writes }], { label: "Grid" })
  return inputs.map((row, ri) =>
    row.map((_, ci) => {
      const value = sheet.cells.value(ri, ci)
      const format = wb.styles.get(sheet.cells.get(ri, ci)?.style ?? 0).numberFormat
      return cellOf(value, displayText(value, format))
    }),
  )
}

export type FormulaResult =
  { ok: true; value: number | string | boolean; text: string } | { ok: false; error: "Invalid formula" | ErrorCode }

/**
 * One formula on its own ("=2+3*4", "=ROUND(PI(),2)", "=TEXT(DATE(2026,1,1),\"mmmm\")"): its value and the text it
 * shows, or why there is none. References read cells of `cells` when given (a block's grid), else blanks.
 */
export function calculateFormula(
  source: string,
  cells: readonly (readonly string[])[] = [],
  options: GridOptions = {},
): FormulaResult {
  const body = source.trim().replace(/^=\s*/, "")
  if (!body || !parseFormula(body).ok) return { ok: false, error: "Invalid formula" }
  // The formula goes in a column to the right of the cells it may read.
  const width = Math.max(1, ...cells.map((row) => row.length))
  const grid = cells.map((row) => [...row])
  grid[0] = [...(grid[0] ?? []), ...Array<string>(Math.max(0, width - (grid[0]?.length ?? 0))).fill("")]
  grid[0][width] = `=${body}`
  const cell = calculateGrid(grid, options)[0]![width]!
  if (cell.error) return { ok: false, error: cell.error }
  return { ok: true, value: cell.value ?? 0, text: cell.text }
}

/** The column letters of a grid `cols` wide, for its header ("A", "B"…). */
export function gridColumns(cols: number): string[] {
  return Array.from({ length: cols }, (_, i) => columnName(i))
}

function cellOf(value: CellValue | undefined, text: string): GridCell {
  if (typeof value === "object" && value !== null) return { value: undefined, error: value.error, text: value.error }
  return { value, text }
}
