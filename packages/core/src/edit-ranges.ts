import type { Range } from "./address"
import { shiftRange } from "./conditional"
import type { PasswordHash } from "./protection"

/**
 * Review › Allow Edit Ranges: ranges of a protected sheet that can still be edited, each with a title and an
 * optional password. A range with a password asks for it once (Unlock Range); it then stays open until the
 * workbook closes. Only the cells' contents open up: formatting still follows Protect Sheet's options.
 */
export type EditRange = {
  /** The title (Excel's "Range1"); unique on the sheet, compared without case. */
  name: string
  ranges: Range[]
  /** Excel's legacy hash of the password, as written for a new one; absent without a password. */
  password?: string
  /** The modern hash, when the file had one. */
  hash?: PasswordHash
  /** Windows users and groups allowed without a password (Permissions...), kept as read. */
  securityDescriptor?: string
}

export function copyEditRanges(list: readonly EditRange[]): EditRange[] {
  return list.map((e) => ({ ...e, ranges: e.ranges.map((r) => ({ ...r })) }))
}

export const hasEditPassword = (e: EditRange) => Boolean(e.password || e.hash)

/** The edit range a cell is in (the first, when ranges overlap), or null. */
export function editRangeAt(list: readonly EditRange[], row: number, col: number): EditRange | null {
  for (const e of list) if (e.ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) return e
  return null
}

/** Whether a cell's edit range lets it change now: one without a password, or one already unlocked. */
export function editRangeOpen(
  list: readonly EditRange[],
  unlocked: ReadonlySet<string>,
  row: number,
  col: number,
): boolean {
  for (const e of list) {
    if (!e.ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) continue
    if (!hasEditPassword(e) || unlocked.has(e.name.toLowerCase())) return true
  }
  return false
}

/** Ranges follow inserted and deleted rows and columns; one whose cells all go is dropped. */
export function shiftEditRanges(list: readonly EditRange[], axis: "rows" | "cols", at: number, count: number) {
  const out: EditRange[] = []
  for (const e of list) {
    const ranges = e.ranges.flatMap((r) => shiftRange(r, axis, at, count) ?? [])
    if (ranges.length) out.push({ ...e, ranges })
  }
  return out
}

/** The next free title: Range1, Range2… (Excel's default). */
export function nextEditRangeName(list: readonly EditRange[]): string {
  const taken = new Set(list.map((e) => e.name.toLowerCase()))
  let n = 1
  while (taken.has(`range${n}`)) n++
  return `Range${n}`
}
