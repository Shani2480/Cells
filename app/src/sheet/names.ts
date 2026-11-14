import {
  MAX_COLS,
  MAX_ROWS,
  nameProblem,
  parseRange,
  parseRefText,
  quoteSheet,
  rangeName,
  sameRange,
  type DefinedName,
  type Range,
} from "@cells/core"
import type { SheetController } from "./controller"

/** Excel's message when the Name Box gets neither a reference nor a valid name. */
export const NAME_BOX_INVALID =
  "You must enter a valid reference you want to go to, or type a valid name for the selection."

export type Target = { sheet: number; range: Range }

/** A reference as `$A$1:$B$2` (whole rows and columns as `$1:$3`, `$A:$C`). */
export function absoluteRef(r: Range): string {
  const name = rangeName(r)
  return name
    .replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")
    .replace(/(^|:)([A-Z]+)(?=:|$)/g, "$1$$$2")
    .replace(/(^|:)(\d+)(?=:|$)/g, "$1$$$2")
}

/**
 * Defined names in the UI: the Name Box (go to a name, or name the selection by typing one), Define Name, the
 * names Go To lists.
 */
export class NameCommands {
  constructor(private readonly c: SheetController) {}

  all(): readonly DefinedName[] {
    return this.c.wb.names.all()
  }

  /** What a reference points at: `B7`, `Sheet2!A1:C3`, or a defined name. Null if it isn't one. */
  resolve(text: string): Target | null {
    const trimmed = text.trim()
    if (!trimmed) return null
    const plain = parseRange(trimmed)
    if (plain) return { sheet: this.c.sheetIndex, range: plain }
    const def = this.c.wb.names.get(trimmed)
    const ref = parseRefText(def ? def.ref : trimmed)
    if (!ref || ref.invalid || ref.sheetTo) return null
    const sheet = ref.sheet === null ? this.c.sheetIndex : this.sheetIndex(ref.sheet)
    if (sheet < 0) return null
    const to = ref.to ?? ref.from
    const r0 = Math.min(ref.from.row ?? 0, to.row ?? 0)
    const r1 = ref.from.row === null ? MAX_ROWS - 1 : Math.max(ref.from.row, to.row ?? ref.from.row)
    const c0 = Math.min(ref.from.col ?? 0, to.col ?? 0)
    const c1 = ref.from.col === null ? MAX_COLS - 1 : Math.max(ref.from.col, to.col ?? ref.from.col)
    return { sheet, range: { r0, r1, c0, c1 } }
  }

  private sheetIndex(name: string): number {
    const upper = name.toUpperCase()
    return this.c.wb.sheets.findIndex((s) => s.name.toUpperCase() === upper)
  }

  /** The selection as a name's "Refers to": `=Sheet1!$A$1:$B$5`. */
  selectionRef(): string {
    const ranges = this.c.selection.ranges
    const prefix = quoteSheet(this.c.sheet.name) + "!"
    return ranges.map((r) => prefix + absoluteRef(r)).join(",")
  }

  /** The name that refers to exactly the selection (the Name Box shows it instead of the cell, as Excel does). */
  nameOfSelection(): string | null {
    const ranges = this.c.selection.ranges
    if (ranges.length !== 1) return null
    for (const n of this.all()) {
      const t = this.resolve(n.name)
      if (t && t.sheet === this.c.sheetIndex && sameRange(t.range, ranges[0])) return n.name
    }
    return null
  }

  goTo(target: Target) {
    if (target.sheet !== this.c.sheetIndex) this.c.activate(target.sheet)
    this.c.goTo(target.range)
  }

  /** Define Name: a new name, or a changed one (`old`); returns Excel's problem text, or null when done. */
  define(name: string, refersTo: string, comment = "", old?: string): string | null {
    const others = this.all().filter((n) => n.name.toUpperCase() !== old?.toUpperCase())
    const problem = nameProblem(name.trim(), others)
    if (problem) return problem
    const ref = refersTo.trim().replace(/^=/, "")
    if (!ref) return "The formula contains an error."
    const next: DefinedName[] = [...others, { name: name.trim(), ref, ...(comment ? { comment } : {}) }]
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    this.c.transact([{ kind: "names", names: next }], { label: "Define Name", before: sel, after: sel })
    return null
  }

  /** The Name Box's Enter: go to a reference or name, else name the selection, else Excel's message. */
  fromNameBox(text: string) {
    const target = this.resolve(text)
    if (target) {
      this.goTo(target)
      return
    }
    const name = text.trim()
    if (!name) return
    const problem = nameProblem(name, this.all())
    if (problem) {
      this.c.shell.alert(NAME_BOX_INVALID)
      return
    }
    this.define(name, this.selectionRef())
  }
}
