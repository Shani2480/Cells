import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { Op } from "./ops"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"
import { sha512 } from "./sha512"
import { editRangeOpen } from "./edit-ranges"

/**
 * Protect Sheet: what users may still do on a protected sheet (Excel's "Allow all users of this worksheet to:"
 * list), and the password's hash. Excel's defaults allow only selecting cells.
 */
/**
 * A password stored the way Excel 365 does: a salted hash repeated `spinCount` times (SHA-512 by default), all
 * base64. Kept as read so a saved file still asks for the same password.
 */
export type PasswordHash = { algorithm: string; hash: string; salt: string; spinCount: number }

export type SheetProtection = {
  /** Excel's legacy 16-bit hash (hex) of the password, as .xlsx stores it; absent without a password. */
  password?: string
  /** The modern hash, when the file had one instead of `password`. */
  hash?: PasswordHash
  selectLocked: boolean
  selectUnlocked: boolean
  formatCells: boolean
  formatColumns: boolean
  formatRows: boolean
  insertColumns: boolean
  insertRows: boolean
  insertHyperlinks: boolean
  deleteColumns: boolean
  deleteRows: boolean
  sort: boolean
  autoFilter: boolean
  pivotTables: boolean
  editObjects: boolean
  editScenarios: boolean
}

/** Protect Workbook (Structure): sheets can't be added, deleted, renamed, moved, copied, hidden or unhidden. */
export type WorkbookProtection = {
  /** Excel's legacy hash of the password; absent without one. */
  password?: string
  hash?: PasswordHash
}

/** The sheet operations workbook protection refuses. */
export const STRUCTURE_OPS: ReadonlySet<Op["kind"]> = new Set(["addSheet", "removeSheet", "renameSheet", "moveSheet"])

/**
 * Whether a protected workbook refuses an operation: the structure ops, and hiding or unhiding a sheet (its tab colour
 * may still change).
 */
export function structureBlocks(op: Op): boolean {
  if (STRUCTURE_OPS.has(op.kind)) return true
  return op.kind === "sheetProps" && op.hidden !== undefined
}

export type ProtectionPermission = Exclude<keyof SheetProtection, "password" | "hash">

export const DEFAULT_PROTECTION: SheetProtection = {
  selectLocked: true,
  selectUnlocked: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  editObjects: false,
  editScenarios: false,
}

/** Excel's legacy sheet-password hash (the `password` attribute of `<sheetProtection>`), upper-case hex. */
export function hashPassword(password: string): string {
  let hash = 0
  for (let i = 0; i < password.length; i++) {
    const value = password.charCodeAt(i) << (i + 1)
    hash ^= (value & 0x7fff) | (value >> 15)
  }
  hash ^= password.length
  hash ^= 0xce4b
  return hash.toString(16).toUpperCase()
}

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

function fromBase64(text: string): Uint8Array {
  const clean = text.replace(/[^A-Za-z0-9+/]/g, "")
  const out = new Uint8Array(Math.floor((clean.length * 3) / 4))
  let bits = 0
  let acc = 0
  let o = 0
  for (const ch of clean) {
    acc = (acc << 6) | B64.indexOf(ch)
    bits += 6
    if (bits >= 8) {
      bits -= 8
      out[o++] = (acc >> bits) & 0xff
    }
  }
  return out.subarray(0, o)
}

/**
 * Excel 365's password hash (ECMA-376 Agile): H0 = SHA-512(salt + password as UTF-16LE), then `spinCount` times
 * H = SHA-512(H + iteration as 32-bit little-endian).
 */
export function modernHashMatches(hash: PasswordHash, password: string): boolean {
  if (hash.algorithm.replace(/-/g, "").toUpperCase() !== "SHA512") return false
  const salt = fromBase64(hash.salt)
  const input = new Uint8Array(salt.length + password.length * 2)
  input.set(salt)
  for (let i = 0; i < password.length; i++) {
    const c = password.charCodeAt(i)
    input[salt.length + i * 2] = c & 0xff
    input[salt.length + i * 2 + 1] = c >> 8
  }
  let h = sha512(input)
  const next = new Uint8Array(68)
  for (let i = 0; i < hash.spinCount; i++) {
    next.set(h)
    next[64] = i & 0xff
    next[65] = (i >>> 8) & 0xff
    next[66] = (i >>> 16) & 0xff
    next[67] = (i >>> 24) & 0xff
    h = sha512(next)
  }
  const want = fromBase64(hash.hash)
  return want.length === h.length && want.every((b, i) => b === h[i])
}

/** Whether a protection (sheet or workbook) has a password at all. */
export const hasPassword = (p: { password?: string; hash?: PasswordHash } | null | undefined) =>
  Boolean(p?.password || p?.hash)

/** Whether `password` unlocks the protection: the legacy hash, or the modern one from files Excel 365 saved. */
export function passwordMatches(p: { password?: string; hash?: PasswordHash }, password: string): boolean {
  if (p.hash) return modernHashMatches(p.hash, password)
  if (p.password) return hashPassword(password) === p.password.toUpperCase()
  return true
}

/** The message Excel shows when an edit hits a locked cell on a protected sheet. */
export const PROTECTED_MESSAGE =
  "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password."

/** Whether a cell's contents can change now: an unprotected sheet, an unlocked cell, or an open edit range. */
export function canEditCell(sheet: Sheet, styles: StyleTable, row: number, col: number): boolean {
  return (
    !sheet.protection ||
    !isLocked(sheet, styles, row, col) ||
    editRangeOpen(sheet.editRanges, sheet.unlockedEditRanges, row, col)
  )
}

/** Whether a cell is locked: the default, unless its format unlocks it (Format Cells › Protection). */
export function isLocked(sheet: Sheet, styles: StyleTable, row: number, col: number): boolean {
  return styles.get(sheet.styleAt(row, col)).locked !== false
}

function anyLocked(sheet: Sheet, styles: StyleTable, ranges: readonly Range[]): boolean {
  const locked = (id: number) => styles.get(id).locked !== false
  for (const r of ranges) {
    if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) <= 100_000) {
      for (let row = r.r0; row <= r.r1; row++)
        for (let col = r.c0; col <= r.c1; col++) if (isLocked(sheet, styles, row, col)) return true
      continue
    }
    // Big ranges: cells with their own format, then the row and column formats the other cells fall back to.
    let hit = false
    sheet.cells.forEachStyle(r, (_row, _col, id) => {
      if (id && locked(id)) hit = true
    })
    if (hit) return true
    let unstyledRow = false
    for (let row = r.r0; row <= r.r1 && !unstyledRow; row++) {
      const id = sheet.rowStyles.get(row)
      if (id) {
        if (locked(id)) return true
      } else unstyledRow = true
    }
    if (unstyledRow) for (let col = r.c0; col <= r.c1; col++) if (locked(sheet.colStyles.get(col))) return true
  }
  return false
}

/**
 * What Protect Sheet lets people select: every cell (Select locked cells on, Excel's default), only unlocked cells
 * (Select locked cells off), or none (both off).
 */
export function selectableCells(sheet: Sheet): "all" | "unlocked" | "none" {
  const p = sheet.protection
  if (!p || p.selectLocked) return "all"
  return p.selectUnlocked ? "unlocked" : "none"
}

/** Whether the sheet's protection lets these ranges be selected. */
export function canSelect(sheet: Sheet, styles: StyleTable, ranges: readonly Range[]): boolean {
  const mode = selectableCells(sheet)
  return mode === "all" || (mode === "unlocked" && !anyLocked(sheet, styles, ranges))
}

/**
 * The cell the cursor goes to from `from` when only unlocked cells can be selected, or null when there's none.
 * "line": the nearest unlocked cell in that direction along the row or column (the arrow keys). "order": the next
 * (or, going left or up, the previous) unlocked cell in reading order, row by row for left/right (Tab) and column by
 * column for up/down (Enter), wrapping round. Reading order looks within the cells that have a format of their own.
 */
export function nextSelectable(
  sheet: Sheet,
  styles: StyleTable,
  from: { row: number; col: number },
  dir: "up" | "down" | "left" | "right",
  how: "line" | "order",
): { row: number; col: number } | null {
  const open = (row: number, col: number) => !isLocked(sheet, styles, row, col)
  const forward = dir === "down" || dir === "right"
  if (how === "line") {
    const dr = dir === "down" ? 1 : dir === "up" ? -1 : 0
    const dc = dir === "right" ? 1 : dir === "left" ? -1 : 0
    for (let row = from.row + dr, col = from.col + dc; row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLS;) {
      if (open(row, col)) return { row, col }
      row += dr
      col += dc
    }
    return null
  }
  const bounds = sheet.cells.usedBounds()
  if (!bounds) return null
  const byRows = dir === "left" || dir === "right"
  const height = bounds.r1 - bounds.r0 + 1
  const width = bounds.c1 - bounds.c0 + 1
  const total = height * width
  if (total > 5_000_000) return null
  // The cell's place in the order, then each next place in turn until it comes round again.
  const index = (row: number, col: number) =>
    byRows ? (row - bounds.r0) * width + (col - bounds.c0) : (col - bounds.c0) * height + (row - bounds.r0)
  const cell = (i: number) =>
    byRows
      ? { row: bounds.r0 + Math.floor(i / width), col: bounds.c0 + (i % width) }
      : { row: bounds.r0 + (i % height), col: bounds.c0 + Math.floor(i / height) }
  const inside = from.row >= bounds.r0 && from.row <= bounds.r1 && from.col >= bounds.c0 && from.col <= bounds.c1
  const start = inside ? index(from.row, from.col) : forward ? -1 : total
  for (let k = 1; k <= total; k++) {
    const i = (((start + (forward ? k : -k)) % total) + total) % total
    const c = cell(i)
    if (open(c.row, c.col)) return c
  }
  return null
}

/**
 * Whether a protected sheet refuses an operation (Excel shows PROTECTED_MESSAGE). Only the user's own commands are
 * checked; undo and redo replay what was allowed. Sheet-level changes (add, rename, move) are workbook protection,
 * which this is not.
 */
export function protectionBlocks(sheet: Sheet, styles: StyleTable, op: Op): boolean {
  const p = sheet.protection
  if (!p) return false
  switch (op.kind) {
    case "setCells":
      return op.cells.some(
        (w) =>
          isLocked(sheet, styles, w.row, w.col) &&
          !editRangeOpen(sheet.editRanges, sheet.unlockedEditRanges, w.row, w.col),
      )
    case "moveCells": {
      const s = op.source
      return anyLocked(sheet, styles, [s, { r0: s.r0 + op.dr, r1: s.r1 + op.dr, c0: s.c0 + op.dc, c1: s.c1 + op.dc }])
    }
    case "format":
    case "styles":
      return !p.formatCells
    case "axis":
      return op.axis === "rows" ? !p.formatRows : !p.formatColumns
    case "insertRows":
      return !p.insertRows
    case "insertCols":
      return !p.insertColumns
    case "deleteRows":
      return !p.deleteRows
    case "deleteCols":
      return !p.deleteColumns
    case "annotations":
    case "charts":
      return !p.editObjects
    case "editRanges":
      // Allow Edit Ranges is greyed out on a protected sheet.
      return true
    case "hyperlinks":
      // Excel: "Insert hyperlinks, even in unlocked cells" is its own option.
      return !p.insertHyperlinks
    case "autoFilter":
      // Use AutoFilter lets filters change; turning the filter on or off never works on a protected sheet.
      return !p.autoFilter || (sheet.autoFilter === null) !== (op.filter === null)
    case "pivots":
      // Use PivotTable & PivotChart.
      return !p.pivotTables
    case "insertCells":
    case "deleteCells":
    case "merge":
    case "unmerge":
    case "layout":
    case "conditional":
    case "validations":
    case "tables":
    case "sparklines":
      return true
    default:
      return false
  }
}
