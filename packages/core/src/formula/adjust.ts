import { MAX_COLS, MAX_ROWS, type Range } from "../address"
import type { RefNode, RefPoint } from "./ast"
import { tokenize } from "./lexer"
import { parseRefText } from "./parser"
import { refText, sheetPrefix } from "./print"

/**
 * Reference rewriting: what Excel does to formula text when rows or cells move under it, when a formula is copied,
 * and when cells are cut and pasted. Only the text of reference tokens changes, so the user's spacing survives.
 * Text inside INDIRECT strings is never touched (Excel doesn't either).
 */

/** A rectangle as a reference covers it; `null` bounds are open (a whole column has no rows). */
type Span = { r0: number; r1: number; c0: number; c1: number }

/**
 * Maps one reference (as an absolute rectangle on its sheet) to its new rectangle, or null when it is destroyed.
 * `onSheet` says whether the reference points at the sheet being changed.
 */
export type RefMapper = (ref: RefNode, onSheet: boolean) => MappedRef | null | undefined

/** A mapper's answer; `toSheet` moves the reference to another sheet (it's then written with that sheet's name). */
export type MappedRef = RefNode & { toSheet?: string }

/**
 * Rewrites `formula` (with or without its "="). `hostSheet` is the name of the sheet the formula is on and
 * `targetSheet` the sheet being changed; references to other sheets are left alone unless the mapper says otherwise.
 * Returns null when nothing changed.
 */
export function rewriteFormula(
  formula: string,
  hostSheet: string,
  targetSheet: string | null,
  map: RefMapper,
): string | null {
  const eq = formula.startsWith("=") ? 1 : 0
  const tokens = tokenize(formula.slice(eq))
  let changed = false
  const target = targetSheet?.toLowerCase() ?? null
  const host = hostSheet.toLowerCase()
  for (const token of tokens) {
    if (token.kind !== "ref") continue
    const ref = parseRefText(token.text)
    if (!ref || ref.invalid) continue
    // 3-D references span sheets; Excel only adjusts them for sheet moves, which Phase 4 handles.
    if (ref.sheetTo !== null) continue
    const sheet = (ref.sheet ?? hostSheet).toLowerCase()
    const result = map(ref, target === null || sheet === target || (ref.sheet === null && host === target))
    if (result === undefined) continue
    const bang = token.text.lastIndexOf("!")
    const prefix =
      result?.toSheet !== undefined ? sheetPrefix(result.toSheet, null) : bang >= 0 ? token.text.slice(0, bang + 1) : ""
    const text =
      result === null ? `${prefix}#REF!` : prefix + refText({ ...withoutSheet(result), sheet: null, sheetTo: null })
    if (text !== token.text) {
      token.text = text
      changed = true
    }
  }
  if (!changed) return null
  return (eq ? "=" : "") + tokens.map((t) => t.text).join("")
}

function withoutSheet({ toSheet: _, ...ref }: MappedRef): RefNode {
  return ref
}

function spanOf(ref: RefNode): Span {
  const to = ref.to ?? ref.from
  const rows = [ref.from.row, to.row]
  const cols = [ref.from.col, to.col]
  return {
    r0: rows[0] === null ? 0 : Math.min(rows[0], rows[1]!),
    r1: rows[0] === null ? MAX_ROWS - 1 : Math.max(rows[0], rows[1]!),
    c0: cols[0] === null ? 0 : Math.min(cols[0], cols[1]!),
    c1: cols[0] === null ? MAX_COLS - 1 : Math.max(cols[0], cols[1]!),
  }
}

/** Puts a new rectangle back into a reference, keeping its $ marks and which end was written first. */
function withSpan(ref: RefNode, span: Span): RefNode {
  const to = ref.to ?? ref.from
  const point = (p: RefPoint, rowEnd: "lo" | "hi", colEnd: "lo" | "hi"): RefPoint => ({
    ...p,
    row: p.row === null ? null : rowEnd === "lo" ? span.r0 : span.r1,
    col: p.col === null ? null : colEnd === "lo" ? span.c0 : span.c1,
  })
  const fromRowLo = ref.from.row === null || to.row === null || ref.from.row <= to.row
  const fromColLo = ref.from.col === null || to.col === null || ref.from.col <= to.col
  const from = point(ref.from, fromRowLo ? "lo" : "hi", fromColLo ? "lo" : "hi")
  const end = ref.to ? point(ref.to, fromRowLo ? "hi" : "lo", fromColLo ? "hi" : "lo") : null
  return { ...ref, from, to: end }
}

const sameSpan = (a: Span, b: Span) => a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1

/**
 * Inserting (count > 0) or deleting (count < 0) lines along one axis at `at`, optionally only within the band
 * `lo..hi` of the other axis (Insert/Delete Cells). Returns undefined for references it leaves alone.
 */
export function shiftMapper(axis: "rows" | "cols", at: number, count: number, band?: [number, number]): RefMapper {
  const size = axis === "rows" ? MAX_ROWS : MAX_COLS
  return (ref, onSheet) => {
    if (!onSheet) return undefined
    const span = spanOf(ref)
    const whole = axis === "rows" ? ref.from.row === null : ref.from.col === null
    if (whole) return undefined
    let lo = axis === "rows" ? span.r0 : span.c0
    let hi = axis === "rows" ? span.r1 : span.c1
    if (band) {
      const [b0, b1] = band
      const o0 = axis === "rows" ? span.c0 : span.r0
      const o1 = axis === "rows" ? span.c1 : span.r1
      // Only references inside the band's columns (rows) move; ones that straddle its edge stay as they are.
      if (o0 < b0 || o1 > b1) return undefined
    }
    if (count > 0) {
      if (hi < at) return undefined
      if (lo >= at) lo += count
      hi += count
      if (lo > size - 1) return null
      hi = Math.min(hi, size - 1)
    } else {
      const n = -count
      const end = at + n // first line after the deleted ones
      if (hi < at) return undefined
      if (lo >= at && hi < end) return null
      lo = lo < at ? lo : lo >= end ? lo - n : at
      hi = hi >= end ? hi - n : at - 1
    }
    const next = axis === "rows" ? { ...span, r0: lo, r1: hi } : { ...span, c0: lo, c1: hi }
    return sameSpan(next, span) ? undefined : withSpan(ref, next)
  }
}

/** Copying a formula `dr` rows and `dc` columns away: relative parts move, absolute ones stay. */
export function copyMapper(dr: number, dc: number): RefMapper {
  return (ref) => {
    const move = (p: RefPoint): RefPoint | null => {
      const row = p.row === null || p.rowAbs ? p.row : p.row + dr
      const col = p.col === null || p.colAbs ? p.col : p.col + dc
      if (row !== null && (row < 0 || row >= MAX_ROWS)) return null
      if (col !== null && (col < 0 || col >= MAX_COLS)) return null
      return { ...p, row, col }
    }
    const from = move(ref.from)
    const to = ref.to ? move(ref.to) : null
    if (!from || (ref.to && !to)) return null
    return { ...ref, from, to }
  }
}

/** Copies a formula from one cell to another, adjusting relative references. */
export function copyFormula(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula
  return rewriteFormula(formula, "", null, copyMapper(dr, dc)) ?? formula
}

/**
 * Cut and paste of `source` by (dr, dc) on the target sheet: references that lie wholly inside the source follow it;
 * references to cells the paste overwrote (and that did not move with it) become #REF!.
 */
export function moveMapper(source: Range, dr: number, dc: number): RefMapper {
  const dest = { r0: source.r0 + dr, r1: source.r1 + dr, c0: source.c0 + dc, c1: source.c1 + dc }
  const inside = (s: Span, r: Range) => s.r0 >= r.r0 && s.r1 <= r.r1 && s.c0 >= r.c0 && s.c1 <= r.c1
  const overlaps = (s: Span, r: Range) => s.r0 <= r.r1 && s.r1 >= r.r0 && s.c0 <= r.c1 && s.c1 >= r.c0
  return (ref, onSheet) => {
    if (!onSheet) return undefined
    const span = spanOf(ref)
    if (ref.from.row === null || ref.from.col === null) return undefined
    if (inside(span, source)) {
      return withSpan(ref, { r0: span.r0 + dr, r1: span.r1 + dr, c0: span.c0 + dc, c1: span.c1 + dc })
    }
    if (inside(span, dest) && !overlaps(span, source)) return null
    return undefined
  }
}

/**
 * Cut on one sheet and pasted on another (`toSheet`), moved by (dr, dc): references wholly inside the source follow
 * the cells there, as in Excel (=A1 on the source sheet becomes =Dst!C3).
 */
export function moveToSheetMapper(source: Range, dr: number, dc: number, toSheet: string): RefMapper {
  return (ref, onSheet) => {
    if (!onSheet || ref.from.row === null || ref.from.col === null) return undefined
    const span = spanOf(ref)
    if (span.r0 < source.r0 || span.r1 > source.r1 || span.c0 < source.c0 || span.c1 > source.c1) return undefined
    return { ...withSpan(ref, { r0: span.r0 + dr, r1: span.r1 + dr, c0: span.c0 + dc, c1: span.c1 + dc }), toSheet }
  }
}

/** References wholly inside `area` (cells a paste from another sheet overwrote) become #REF!. */
export function overwrittenMapper(area: Range): RefMapper {
  return (ref, onSheet) => {
    if (!onSheet || ref.from.row === null || ref.from.col === null) return undefined
    const span = spanOf(ref)
    return span.r0 >= area.r0 && span.r1 <= area.r1 && span.c0 >= area.c0 && span.c1 <= area.c1 ? null : undefined
  }
}

/**
 * A formula cut from `fromSheet` to `toSheet`, as Excel rewrites it: references to cells that moved with it follow
 * them (=A5*2 → =G7*2); the others keep pointing at the old sheet (=A2+1 → =Src!A2+1).
 */
export function movedFormulaMapper(
  source: Range,
  dr: number,
  dc: number,
  fromSheet: string,
  toSheet: string,
): RefMapper {
  return (ref, onSheet) => {
    if (!onSheet) return undefined
    const span = spanOf(ref)
    const whole = ref.from.row === null || ref.from.col === null
    if (!whole && span.r0 >= source.r0 && span.r1 <= source.r1 && span.c0 >= source.c0 && span.c1 <= source.c1) {
      const shifted = withSpan(ref, { r0: span.r0 + dr, r1: span.r1 + dr, c0: span.c0 + dc, c1: span.c1 + dc })
      return ref.sheet === null ? shifted : { ...shifted, toSheet }
    }
    return ref.sheet === null ? { ...ref, toSheet: fromSheet } : undefined
  }
}

/**
 * A sheet renamed (`to` a name) or deleted (`to` null): references that name it take the new name, or become #REF!
 * (as Excel writes =#REF!A1 for a reference to a deleted sheet). Returns null when nothing changed.
 */
export function renameSheetRefs(formula: string, from: string, to: string | null): string | null {
  const eq = formula.startsWith("=") ? 1 : 0
  const tokens = tokenize(formula.slice(eq))
  const name = from.toLowerCase()
  let changed = false
  for (const token of tokens) {
    if (token.kind !== "ref") continue
    const ref = parseRefText(token.text)
    if (!ref || ref.invalid || ref.sheet === null) continue
    const hitFrom = ref.sheet.toLowerCase() === name
    const hitTo = ref.sheetTo?.toLowerCase() === name
    if (!hitFrom && !hitTo) continue
    const bang = token.text.lastIndexOf("!")
    const rest = token.text.slice(bang + 1)
    if (to === null) token.text = `#REF!${rest}`
    else
      token.text = sheetPrefix(hitFrom ? to : ref.sheet, ref.sheetTo === null ? null : hitTo ? to : ref.sheetTo) + rest
    changed = true
  }
  if (!changed) return null
  return (eq ? "=" : "") + tokens.map((t) => t.text).join("")
}
