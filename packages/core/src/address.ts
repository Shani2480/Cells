/** Sheet bounds, as Excel's .xlsx sheets. Rows and columns are zero-based everywhere in the core. */
export const MAX_ROWS = 1_048_576
export const MAX_COLS = 16_384

export type CellAddress = { row: number; col: number }

/** An inclusive, normalised rectangle (r0 ≤ r1, c0 ≤ c1). */
export type Range = { r0: number; c0: number; r1: number; c1: number }

export function cell(row: number, col: number): CellAddress {
  return { row, col }
}

export function rangeOf(a: CellAddress, b: CellAddress = a): Range {
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  }
}

export function sameCell(a: CellAddress, b: CellAddress) {
  return a.row === b.row && a.col === b.col
}

export function sameRange(a: Range, b: Range) {
  return a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1
}

export function rangeContains(range: Range, row: number, col: number) {
  return row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1
}

export function rangeIntersection(a: Range, b: Range): Range | null {
  const r0 = Math.max(a.r0, b.r0)
  const c0 = Math.max(a.c0, b.c0)
  const r1 = Math.min(a.r1, b.r1)
  const c1 = Math.min(a.c1, b.c1)
  return r0 <= r1 && c0 <= c1 ? { r0, c0, r1, c1 } : null
}

export function rangeRows(range: Range) {
  return range.r1 - range.r0 + 1
}

export function rangeCols(range: Range) {
  return range.c1 - range.c0 + 1
}

export function isWholeColumns(range: Range) {
  return range.r0 === 0 && range.r1 === MAX_ROWS - 1
}

export function isWholeRows(range: Range) {
  return range.c0 === 0 && range.c1 === MAX_COLS - 1
}

export function clampCell(address: CellAddress): CellAddress {
  return {
    row: Math.min(MAX_ROWS - 1, Math.max(0, address.row)),
    col: Math.min(MAX_COLS - 1, Math.max(0, address.col)),
  }
}

/** 0 → "A", 25 → "Z", 26 → "AA", 16383 → "XFD". */
export function columnName(col: number): string {
  let n = col + 1
  let name = ""
  while (n > 0) {
    const rem = (n - 1) % 26
    name = String.fromCharCode(65 + rem) + name
    n = Math.floor((n - 1) / 26)
  }
  return name
}

/** "A" → 0, "xfd" → 16383; -1 when not a column name inside the sheet. */
export function columnIndex(name: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(name)) {
    return -1
  }
  let n = 0
  for (const ch of name.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64)
  }
  return n - 1 < MAX_COLS ? n - 1 : -1
}

export function cellName(address: CellAddress): string {
  return `${columnName(address.col)}${address.row + 1}`
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/

/** Parses "B7" or "$B$7"; null when it is not a cell inside the sheet. */
export function parseCell(text: string): CellAddress | null {
  const match = CELL_RE.exec(text.trim())
  if (!match) {
    return null
  }
  const col = columnIndex(match[1])
  const row = Number(match[2]) - 1
  if (col < 0 || row < 0 || row >= MAX_ROWS) {
    return null
  }
  return { row, col }
}

/**
 * Parses a reference as typed in the Name Box or Go To: "B7", "A1:C3", "A:C" (whole columns), "3:5" (whole rows),
 * with or without `$`. Returns null for anything else.
 */
export function parseRange(text: string): Range | null {
  const trimmed = text.trim()
  const parts = trimmed.split(":")
  if (parts.length === 1) {
    const single = parseCell(parts[0])
    return single ? rangeOf(single) : null
  }
  if (parts.length !== 2) {
    return null
  }
  const [left, right] = parts.map((part) => part.replace(/\$/g, ""))
  const a = parseCell(left)
  const b = parseCell(right)
  if (a && b) {
    return rangeOf(a, b)
  }
  const ca = columnIndex(left)
  const cb = columnIndex(right)
  if (ca >= 0 && cb >= 0) {
    return { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(ca, cb), c1: Math.max(ca, cb) }
  }
  if (/^\d{1,7}$/.test(left) && /^\d{1,7}$/.test(right)) {
    const ra = Number(left) - 1
    const rb = Number(right) - 1
    if (ra >= 0 && rb >= 0 && ra < MAX_ROWS && rb < MAX_ROWS) {
      return { r0: Math.min(ra, rb), r1: Math.max(ra, rb), c0: 0, c1: MAX_COLS - 1 }
    }
  }
  return null
}

/** How Excel writes a range: "B7", "A1:C3", "A:C", "3:5". */
export function rangeName(range: Range): string {
  if (isWholeColumns(range) && !isWholeRows(range)) {
    return `${columnName(range.c0)}:${columnName(range.c1)}`
  }
  if (isWholeRows(range) && !isWholeColumns(range)) {
    return `${range.r0 + 1}:${range.r1 + 1}`
  }
  const start = cellName({ row: range.r0, col: range.c0 })
  if (range.r0 === range.r1 && range.c0 === range.c1) {
    return start
  }
  return `${start}:${cellName({ row: range.r1, col: range.c1 })}`
}
