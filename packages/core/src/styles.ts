/**
 * Cell formats, interned: a cell, row or column stores a style id, the workbook keeps one record per distinct format.
 * Id 0 is the default (General, no direct formatting): a cell with style 0 shows its row's format, else its
 * column's. Id 1 is also the default format, but set on purpose (Clear Formats on a cell of a formatted column), so
 * it does not fall back to the row or column.
 *
 * Fields follow Excel's cell format record (xf); an absent field is Excel's default.
 */

export type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "hair"
  | "dotted"
  | "dashed"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot"
  | "double"

/** A border line; no colour means Automatic (the window text colour, black). */
export type Border = { style: BorderStyle; color?: string }

export type HAlign = "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed"
export type VAlign = "top" | "center" | "bottom" | "justify" | "distributed"
export type Underline = "single" | "double" | "singleAccounting" | "doubleAccounting"

export type CellStyle = {
  numberFormat?: string
  // Font
  fontName?: string
  /** Points. */
  fontSize?: number
  bold?: true
  italic?: true
  underline?: Underline
  strike?: true
  vertAlign?: "superscript" | "subscript"
  /** "#rrggbb"; absent is Automatic. */
  fontColor?: string
  // Fill
  /** Solid fill, "#rrggbb"; absent is No Fill. */
  fill?: string
  // Alignment
  hAlign?: HAlign
  /** Absent is bottom, Excel's default. */
  vAlign?: VAlign
  indent?: number
  wrap?: true
  shrink?: true
  /** Excel's encoding: 1–90 counterclockwise, 91–180 clockwise (91 = 1° down), 255 stacked vertical text. */
  rotation?: number
  // Borders
  borderTop?: Border
  borderBottom?: Border
  borderLeft?: Border
  borderRight?: Border
  borderDiagonal?: Border
  diagonalUp?: true
  diagonalDown?: true
  // Protection (applies when the sheet is protected)
  /** Cells are locked by default; false unlocks. */
  locked?: false
  hiddenFormula?: true
  /** The cell style applied from the gallery ("Good", "Heading 1"…), Normal when absent. */
  cellStyle?: string
}

export type StyleKey = keyof CellStyle

/** Changes to a format: a value sets the field, null removes it (back to Excel's default). */
export type StylePatch = { [K in StyleKey]?: CellStyle[K] | null }

export const DEFAULT_STYLE_ID = 0
/** The default format, set explicitly: it does not fall back to the row's or column's format. */
export const EXPLICIT_DEFAULT_ID = 1

export class StyleTable {
  private readonly styles: CellStyle[] = [Object.freeze({}), Object.freeze({})]
  private readonly index = new Map<string, number>([[keyOf({}), 0]])
  /** (base id, patch) → result id: formatting 100,000 cells with a few distinct formats derives a few times. */
  private readonly patched = new Map<string, number>()

  get(id: number): CellStyle {
    return this.styles[id] ?? this.styles[0]
  }

  /** The id of a style with these fields, adding it if new. */
  intern(style: CellStyle): number {
    const key = keyOf(style)
    const found = this.index.get(key)
    if (found !== undefined) {
      return found
    }
    const id = this.styles.length
    this.styles.push(Object.freeze({ ...style }))
    this.index.set(key, id)
    return id
  }

  /** The id of `base` with `patch` applied (null fields removed). */
  derive(base: number, patch: StylePatch): number {
    // A field the patch leaves undefined is not touched; null removes it.
    const merged: Record<string, unknown> = { ...this.get(base) }
    for (const [key, value] of Object.entries(patch)) if (value !== undefined) merged[key] = value
    for (const key of Object.keys(merged)) {
      if (merged[key] === undefined || merged[key] === null) {
        delete merged[key]
      }
    }
    return this.intern(merged as CellStyle)
  }

  /** As derive, cached by the patch's key (from `patchKey`). */
  deriveCached(base: number, patch: StylePatch, key: string): number {
    const cacheKey = `${base}|${key}`
    let id = this.patched.get(cacheKey)
    if (id === undefined) {
      id = this.derive(base, patch)
      this.patched.set(cacheKey, id)
    }
    return id
  }

  /** Whether the id stands for Excel's default format (0, or 1 set on purpose). */
  isDefault(id: number) {
    return id === DEFAULT_STYLE_ID || id === EXPLICIT_DEFAULT_ID
  }

  get size() {
    return this.styles.length
  }
}

/** A stable text key for a patch, for caching. */
export function patchKey(patch: StylePatch): string {
  return JSON.stringify(
    Object.keys(patch)
      .filter((key) => patch[key as StyleKey] !== undefined)
      .sort()
      .map((key) => [key, patch[key as StyleKey]]),
  )
}

function keyOf(style: CellStyle) {
  return patchKey(style)
}

/** The patch that removes every field: Clear Formats. */
export const CLEAR_ALL: StylePatch = {
  numberFormat: null,
  fontName: null,
  fontSize: null,
  bold: null,
  italic: null,
  underline: null,
  strike: null,
  vertAlign: null,
  fontColor: null,
  fill: null,
  hAlign: null,
  vAlign: null,
  indent: null,
  wrap: null,
  shrink: null,
  rotation: null,
  borderTop: null,
  borderBottom: null,
  borderLeft: null,
  borderRight: null,
  borderDiagonal: null,
  diagonalUp: null,
  diagonalDown: null,
  locked: null,
  hiddenFormula: null,
  cellStyle: null,
}
