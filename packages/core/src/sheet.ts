import { MAX_COLS, MAX_ROWS } from "./address"
import { Axis } from "./axis"
import { CellStore } from "./cell-store"

/**
 * Default sizes at 100% zoom, as Excel draws a new workbook with an 11 pt body font: rows 15 pt (20 px), columns
 * 8.43 characters (64 px).
 */
export const DEFAULT_ROW_HEIGHT = 20
export const DEFAULT_COL_WIDTH = 64

export class Sheet {
  name: string
  readonly cells = new CellStore()
  readonly rows = new Axis(MAX_ROWS, DEFAULT_ROW_HEIGHT)
  readonly cols = new Axis(MAX_COLS, DEFAULT_COL_WIDTH)

  constructor(name: string) {
    this.name = name
  }
}
