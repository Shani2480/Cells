import {
  canEditCell,
  copyEditRanges,
  editRangeAt,
  hasEditPassword,
  isLocked,
  parseRange,
  passwordMatches,
  columnName,
  type EditRange,
  type Op,
  type Range,
} from "@cells/core"
import type { SheetController } from "./controller"

/** Excel's messages for Allow Edit Ranges (Cells' wording where noted in verify-list). */
export const EDIT_RANGE_MESSAGES = {
  wrongPassword:
    "The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.",
  badReference: "The reference isn't valid. Enter cells on this sheet, such as A1:B5 or A1:A3,C1.",
  takenTitle: "A range with this title already exists. Enter a different title.",
  noTitle: "Enter a title for the range.",
}

/** "=$B$2:$B$5,$D$1" (a sheet name in front is allowed) as ranges; null when any part isn't a reference. */
export function parseRefersTo(text: string): Range[] | null {
  const parts = text.trim().replace(/^=/, "").split(",")
  const out: Range[] = []
  for (const part of parts) {
    const ref = part
      .slice(part.lastIndexOf("!") + 1)
      .replace(/\$/g, "")
      .trim()
    const r = ref ? parseRange(ref) : null
    if (!r) return null
    out.push(r)
  }
  return out.length ? out : null
}

/** Ranges as Excel shows them in the dialog: =$B$2:$B$5,$D$1. */
export function refersToText(ranges: readonly Range[]): string {
  const cell = (row: number, col: number) => "$" + columnName(col) + "$" + (row + 1)
  const one = (r: Range) =>
    r.r0 === r.r1 && r.c0 === r.c1 ? cell(r.r0, r.c0) : cell(r.r0, r.c0) + ":" + cell(r.r1, r.c1)
  return "=" + ranges.map(one).join(",")
}

/** Review › Allow Users to Edit Ranges on the sheet shown. */
export class EditRangeCommands {
  constructor(private readonly c: SheetController) {}

  list(): readonly EditRange[] {
    return this.c.sheet.editRanges
  }

  /** Replaces the sheet's list (the dialog's OK or Apply); one undo step. */
  save(list: readonly EditRange[]) {
    const c = this.c
    const op: Op = { kind: "editRanges", sheet: c.sheetIndex, ranges: copyEditRanges(list) }
    c.transact([op], { label: "Allow Edit Ranges", before: c.selection, after: c.selection, perSheet: true })
  }

  /** The edit range with a password that keeps a locked cell shut now (null: none, or the cell is free). */
  shutRangeAt(row: number, col: number): EditRange | null {
    const sheet = this.c.sheet
    if (canEditCell(sheet, this.c.wb.styles, row, col)) return null
    if (!isLocked(sheet, this.c.wb.styles, row, col)) return null
    for (const e of sheet.editRanges)
      if (hasEditPassword(e) && e.ranges.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) return e
    return null
  }

  /** Unlock Range: the range's password opens it until the workbook closes. */
  unlock(name: string, password: string): boolean {
    const sheet = this.c.sheet
    const e = sheet.editRanges.find((r) => r.name.toLowerCase() === name.toLowerCase())
    if (!e || !passwordMatches(e, password)) return false
    sheet.unlockedEditRanges.add(e.name.toLowerCase())
    return true
  }

  /** The edit range under a cell (for tests and the dialog's default). */
  at(row: number, col: number): EditRange | null {
    return editRangeAt(this.c.sheet.editRanges, row, col)
  }
}
