/**
 * Cell formats, interned: a cell stores a style id, the workbook keeps one record per distinct format. Id 0 is the
 * default (General, no direct formatting). Phase 1 only needs the number format set on entry; fonts, fills, borders
 * and alignment join this record in Phase 3.
 */
export type CellStyle = {
  numberFormat?: string
}

export const DEFAULT_STYLE_ID = 0

export class StyleTable {
  private readonly styles: CellStyle[] = [{}]
  private readonly index = new Map<string, number>([[keyOf({}), 0]])

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

  /** The id of `base` with `patch` applied. */
  derive(base: number, patch: Partial<CellStyle>): number {
    const merged: CellStyle = { ...this.get(base), ...patch }
    for (const key of Object.keys(merged) as (keyof CellStyle)[]) {
      if (merged[key] === undefined) {
        delete merged[key]
      }
    }
    return this.intern(merged)
  }

  get size() {
    return this.styles.length
  }
}

function keyOf(style: CellStyle) {
  return JSON.stringify(
    Object.keys(style)
      .sort()
      .map((key) => [key, style[key as keyof CellStyle]]),
  )
}
