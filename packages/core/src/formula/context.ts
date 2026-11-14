import type { Node } from "./ast"
import type { Area, Scalar, Value } from "./types"
import type { Table } from "../table/table"

/** What the evaluator needs from the workbook. The calc engine implements it; tests use a small in-memory one. */
export interface EvalContext {
  /** Where the formula is. */
  readonly sheet: number
  readonly row: number
  readonly col: number
  /** A cell's value (a formula cell is calculated first if it is out of date). */
  cell(sheet: number, row: number, col: number): Scalar
  /** Visits the non-blank cells of an area, row by row. Far cheaper than reading every cell of A:A. */
  values(area: Area, visit: (value: Scalar, row: number, col: number) => void): void
  /** The used part of an area (whole columns end at the last used row), or null when nothing is there. */
  usedArea(area: Area): Area | null
  sheetIndex(name: string): number
  sheetName(index: number): string
  sheetCount(): number
  /** A defined name's formula, workbook-wide or for the sheet, or null. */
  defined(name: string, sheet: number): Node | null
  /** The area a dynamic array formula at this cell spills into, or null. */
  spillArea(sheet: number, row: number, col: number): Area | null
  isFormula(sheet: number, row: number, col: number): boolean
  formulaText(sheet: number, row: number, col: number): string | null
  /** Serial of the current date and time (volatile). */
  now(): number
  random(): number
  /** Called by INDIRECT and OFFSET so the engine knows the formula's dynamic precedents. */
  dynamicReference?(area: Area): void
  /**
   * A number that changes whenever a cell of the column changes; lookup indexes are cached against it. Without it,
   * lookups scan.
   */
  columnVersion?(sheet: number, col: number): number
  /** Scratch space that lives as long as the workbook (lookup indexes). */
  cache?: Map<string, unknown>
  /** Hidden rows (manually or by a filter): SUBTOTAL 101–111 and AGGREGATE skip them. */
  isHiddenRow?(sheet: number, row: number): boolean
  /** The workbook's tables, for structured references (Table1[Qty], [@Qty]). */
  tables?(): readonly { sheet: number; table: Table }[]
  /** Names bound by LET (and LAMBDA parameters). */
  bound?: Map<string, Value>
}
