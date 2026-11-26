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

/**
 * Where a reference or name points: a range on a sheet, and every area when it's a list (`B1,D3:E4`, or a name made
 * from a selection of several ranges); `range` is then the first.
 */
export type Target = { sheet: number; range: Range; areas?: Range[] }

/** A list of references split at its commas (or semicolons, the list separator where commas are decimal points). */
function splitList(text: string): string[] {
  const parts: string[] = []
  let current = ""
  let quoted = false
  for (const ch of text) {
    if (ch === "'") quoted = !quoted
    if (!quoted && (ch === "," || ch === ";")) {
      parts.push(current)
      current = ""
    } else current += ch
  }
  parts.push(current)
  return parts.map((p) => p.trim())
}

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

  /**
   * What a reference points at: `B7`, `Sheet2!A1:C3`, a defined name, a table (its data rows, as Excel goes to them),
   * or a list of these on one sheet (`B1,D3:E4`). Null if it isn't one.
   */
  resolve(text: string, depth = 0): Target | null {
    const parts = splitList(text)
    if (parts.length === 1) return this.resolveOne(parts[0], depth)
    if (parts.some((p) => !p)) return null
    let sheet = -1
    const areas: Range[] = []
    for (const part of parts) {
      const t = this.resolveOne(part, depth)
      if (!t || (sheet >= 0 && t.sheet !== sheet)) return null
      sheet = t.sheet
      areas.push(...(t.areas ?? [t.range]))
    }
    return { sheet, range: areas[0], areas }
  }

  private resolveOne(text: string, depth: number): Target | null {
    const trimmed = text.trim()
    if (!trimmed) return null
    const plain = parseRange(trimmed)
    if (plain) return { sheet: this.c.sheetIndex, range: plain }
    const def = this.c.wb.names.get(trimmed)
    // A name made from several ranges refers to a list (a name of a name is followed a few steps, no further).
    if (def && splitList(def.ref).length > 1) return depth < 3 ? this.resolve(def.ref, depth + 1) : null
    if (!def) {
      const table = this.table(trimmed)
      if (table) return table
    }
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

  /** A table by name: its data rows (without the header and total rows), or the whole table when it has none. */
  private table(name: string): Target | null {
    const upper = name.toUpperCase()
    for (const [sheet, s] of this.c.wb.sheets.entries()) {
      const t = s.tables.find((x) => x.name.toUpperCase() === upper)
      if (!t) continue
      const r0 = t.range.r0 + (t.header ? 1 : 0)
      const r1 = t.range.r1 - (t.totals ? 1 : 0)
      return { sheet, range: r1 >= r0 ? { ...t.range, r0, r1 } : t.range }
    }
    return null
  }

  /**
   * The Name Box's list, as Excel's: the defined names that point at cells (not constants or formulas) and the
   * tables, A to Z.
   */
  boxNames(): string[] {
    const names = this.all()
      .filter((n) => this.resolve(n.name) !== null)
      .map((n) => n.name)
    const tables = this.c.wb.sheets.flatMap((s) => s.tables.map((t) => t.name))
    return [...names, ...tables].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }))
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
    for (const n of this.all()) {
      const t = this.resolve(n.name)
      if (!t || t.sheet !== this.c.sheetIndex) continue
      const areas = t.areas ?? [t.range]
      if (areas.length === ranges.length && areas.every((a, i) => sameRange(a, ranges[i]))) return n.name
    }
    return null
  }

  /** Selects the target, every area of a list, the last area's first cell active (as Excel goes to one). */
  goTo(target: Target) {
    if (target.sheet !== this.c.sheetIndex) this.c.activate(target.sheet)
    if (!target.areas || target.areas.length < 2) {
      this.c.goTo(target.range)
      return
    }
    const ranges = target.areas
    const last = ranges[ranges.length - 1]
    const active = { row: last.r0, col: last.c0 }
    this.c.setSelection(
      { ranges, active, activeRange: ranges.length - 1, anchor: active, focus: { row: last.r1, col: last.c1 } },
      active,
    )
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

  /** Name Manager's Delete. */
  remove(names: string[]) {
    const drop = new Set(names.map((n) => n.toUpperCase()))
    const next = this.all().filter((n) => !drop.has(n.name.toUpperCase()))
    if (next.length === this.all().length) return
    const sel = this.c.selection
    this.c.transact([{ kind: "names", names: next }], { label: "Delete Name", before: sel, after: sel })
  }

  /**
   * Create from Selection: which edges of the selection hold labels, as Excel guesses (text in the first row, and in
   * the first column when its first cell is empty or the row isn't all labels).
   */
  guessEdges(): CreateEdges {
    const r = this.c.selection.ranges[0]
    const text = (row: number, col: number) => typeof this.c.sheet.cells.value(row, col) === "string"
    const top =
      r.r1 > r.r0 &&
      Array.from({ length: r.c1 - r.c0 + 1 }, (_, i) => r.c0 + i).every(
        (c) => !this.c.sheet.cells.value(r.r0, c) || text(r.r0, c),
      )
    const left =
      r.c1 > r.c0 &&
      Array.from({ length: r.r1 - r.r0 + 1 }, (_, i) => r.r0 + i)
        .slice(top ? 1 : 0)
        .every((row) => text(row, r.c0))
    return { top: top && r.c1 - r.c0 >= 0, left, bottom: false, right: false }
  }

  /** Create from Selection: a name for each row or column of the selection, from the labels on the chosen edges. */
  createFromSelection(edges: CreateEdges): string | null {
    const r = this.c.selection.ranges[0]
    if (!r) return null
    const inner = {
      r0: r.r0 + (edges.top ? 1 : 0),
      r1: r.r1 - (edges.bottom ? 1 : 0),
      c0: r.c0 + (edges.left ? 1 : 0),
      c1: r.c1 - (edges.right ? 1 : 0),
    }
    if (inner.r0 > inner.r1 || inner.c0 > inner.c1) return "Cells can't create names from this selection."
    const prefix = quoteSheet(this.c.sheet.name) + "!"
    const made = new Map<string, string>()
    const add = (label: unknown, range: Range) => {
      const name = nameFromLabel(label)
      if (name) made.set(name.toUpperCase(), `${name}\u0000${prefix}${absoluteRef(range)}`)
    }
    if (edges.top)
      for (let c = inner.c0; c <= inner.c1; c++) add(this.c.sheet.cells.value(r.r0, c), { ...inner, c0: c, c1: c })
    if (edges.bottom)
      for (let c = inner.c0; c <= inner.c1; c++) add(this.c.sheet.cells.value(r.r1, c), { ...inner, c0: c, c1: c })
    if (edges.left)
      for (let row = inner.r0; row <= inner.r1; row++)
        add(this.c.sheet.cells.value(row, r.c0), { ...inner, r0: row, r1: row })
    if (edges.right)
      for (let row = inner.r0; row <= inner.r1; row++)
        add(this.c.sheet.cells.value(row, r.c1), { ...inner, r0: row, r1: row })
    if (!made.size) return "Cells found no labels to make names from."
    const next: DefinedName[] = this.all().filter((n) => !made.has(n.name.toUpperCase()))
    for (const entry of made.values()) {
      const [name, ref] = entry.split("\u0000")
      next.push({ name, ref })
    }
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    this.c.transact([{ kind: "names", names: next }], { label: "Create Names", before: sel, after: sel })
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

export type CreateEdges = { top: boolean; left: boolean; bottom: boolean; right: boolean }

/** A label as a name, as Excel makes one: spaces and other characters become "_", a leading digit gets "_". */
export function nameFromLabel(label: unknown): string | null {
  if (label === undefined || label === null || label === "") return null
  let name = String(label)
    .trim()
    .replace(/[^A-Za-z0-9_.\\]/g, "_")
  if (!name) return null
  if (!/^[A-Za-z_\\]/.test(name)) name = "_" + name
  return nameProblem(name) ? "_" + name : name
}
