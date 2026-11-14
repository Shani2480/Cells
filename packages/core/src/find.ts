import { sameValidation, validationAt } from "./validation"
import { MAX_COLS, MAX_ROWS, type Range } from "./address"
import type { CellData } from "./cell-store"
import { wildcardRegExp } from "./filter"
import { displayText, editTextFor } from "./format"
import type { Node, RefNode } from "./formula/ast"
import { parseFormula } from "./formula/parser"
import { currentRegion } from "./navigation"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"
import { isError } from "./values"

/**
 * Find and Replace (Ctrl+F, Ctrl+H) and Go To Special, as Excel does them: wildcards (`*`, `?`, `~`), case and whole
 * cell matching, looking in formulas (what was typed), values (what shows), notes or comments, by rows or by columns.
 */

export type CellAt = { row: number; col: number }

export type FindOptions = {
  text: string
  matchCase?: boolean
  /** Match entire cell contents. */
  entire?: boolean
  /** Search by columns (down each column, then the next) instead of by rows. */
  byColumns?: boolean
  lookIn: "formulas" | "values" | "notes" | "comments"
}

/** The text a cell offers to Find: what was typed (formulas), what shows (values), or its note or comment. */
export function searchText(
  sheet: Sheet,
  styles: StyleTable,
  row: number,
  col: number,
  lookIn: FindOptions["lookIn"],
  data = sheet.cells.get(row, col),
): string | null {
  if (lookIn === "notes" || lookIn === "comments") {
    const a = sheet.annotations.get(row, col)
    if (!a) return null
    if (a.kind === "note") return lookIn === "notes" ? a.text : null
    return lookIn === "comments" ? a.posts.map((p) => p.text).join("\n") : null
  }
  if (!data || data.value === undefined) return null
  if (lookIn === "values") return displayText(data.value, styles.get(sheet.styleAt(row, col)).numberFormat)
  // A spilled value isn't the cell's own content: Excel finds it only among values.
  if (data.spill) return null
  return data.input ?? editTextFor(data.value, styles.get(sheet.styleAt(row, col)).numberFormat)
}

/** The pattern of a search: the whole text (Match entire cell contents) or anywhere in it. */
export function findRegExp(o: Pick<FindOptions, "text" | "matchCase" | "entire">, global = false): RegExp {
  const source = wildcardRegExp(o.entire ? o.text : `*${o.text}*`).source
  if (global) {
    // For replacing: the pattern itself, every occurrence (Excel's * takes as much as it can).
    const inner = wildcardRegExp(o.text).source.slice(1, -1)
    return new RegExp(o.entire ? `^${inner}$` : inner || "(?:)", o.matchCase ? "g" : "gi")
  }
  return new RegExp(source, o.matchCase ? "" : "i")
}

/** Whether a cell matches. */
export function cellMatches(
  sheet: Sheet,
  styles: StyleTable,
  row: number,
  col: number,
  o: FindOptions,
  re = findRegExp(o),
) {
  const text = searchText(sheet, styles, row, col, o.lookIn)
  return text !== null && re.test(text)
}

/** Where a search looks: the selection when it's more than one cell, else the whole sheet's used cells. */
export function searchAreas(sheet: Sheet, selection: readonly Range[]): Range[] {
  const single = selection.length === 1 && selection[0].r0 === selection[0].r1 && selection[0].c0 === selection[0].c1
  if (!single) return selection.map((r) => ({ ...r }))
  const used = sheet.cells.usedBounds()
  const notes = sheet.annotations.all()
  if (!used && !notes.length) return []
  let r0 = used?.r0 ?? MAX_ROWS
  let c0 = used?.c0 ?? MAX_COLS
  let r1 = used?.r1 ?? -1
  let c1 = used?.c1 ?? -1
  for (const a of notes) {
    r0 = Math.min(r0, a.row)
    c0 = Math.min(c0, a.col)
    r1 = Math.max(r1, a.row)
    c1 = Math.max(c1, a.col)
  }
  return [{ r0, c0, r1, c1 }]
}

const orderOf = (byColumns: boolean | undefined) => (a: CellAt, b: CellAt) =>
  byColumns ? a.col - b.col || a.row - b.row : a.row - b.row || a.col - b.col

/** Cells of `areas` (each once) that could hold text: cells with content, and annotated cells. */
function candidates(sheet: Sheet, areas: readonly Range[], annotations: boolean): CellAt[] {
  const seen = new Set<string>()
  const out: CellAt[] = []
  const add = (row: number, col: number) => {
    const key = `${row},${col}`
    if (seen.has(key)) return
    seen.add(key)
    out.push({ row, col })
  }
  for (const area of areas) {
    if (annotations) {
      for (const a of sheet.annotations.all())
        if (a.row >= area.r0 && a.row <= area.r1 && a.col >= area.c0 && a.col <= area.c1) add(a.row, a.col)
    } else sheet.cells.forEach(area, (row, col) => add(row, col), true)
  }
  return out
}

/** Find All: every matching cell, in search order. */
export function findAll(sheet: Sheet, styles: StyleTable, o: FindOptions, areas: readonly Range[]): CellAt[] {
  if (!o.text) return []
  const re = findRegExp(o)
  const annotations = o.lookIn === "notes" || o.lookIn === "comments"
  return candidates(sheet, areas, annotations)
    .filter((c) => cellMatches(sheet, styles, c.row, c.col, o, re))
    .sort(orderOf(o.byColumns))
}

/** Find Next (Shift+Find Next: previous): the match after `from` in search order, going round. */
export function findNext(
  sheet: Sheet,
  styles: StyleTable,
  o: FindOptions,
  areas: readonly Range[],
  from: CellAt,
  backwards = false,
): CellAt | null {
  const all = findAll(sheet, styles, o, areas)
  if (!all.length) return null
  const order = orderOf(o.byColumns)
  if (backwards) {
    for (let i = all.length - 1; i >= 0; i--) if (order(all[i], from) < 0) return all[i]
    return all[all.length - 1]
  }
  return all.find((c) => order(c, from) > 0) ?? all[0]
}

/** The text after Replace, or null when the cell doesn't match. */
export function replacedText(text: string, o: FindOptions, replacement: string): string | null {
  const re = findRegExp(o, true)
  if (o.entire) return findRegExp(o).test(text) ? replacement : null
  if (!findRegExp(o).test(text)) return null
  // Replacement text is literal (no $1).
  return text.replace(re, () => replacement)
}

// ── Go To Special ──

export type ValueKinds = { numbers: boolean; text: boolean; logicals: boolean; errors: boolean }

export type SpecialKind =
  | { kind: "notes" }
  | { kind: "constants"; types: ValueKinds }
  | { kind: "formulas"; types: ValueKinds }
  | { kind: "blanks" }
  | { kind: "currentRegion" }
  | { kind: "currentArray" }
  | { kind: "rowDifferences" }
  | { kind: "columnDifferences" }
  | { kind: "precedents"; all: boolean }
  | { kind: "dependents"; all: boolean }
  | { kind: "lastCell" }
  | { kind: "visible" }
  | { kind: "conditionalFormats"; same: boolean }
  | { kind: "validation"; same: boolean }

const typeOf = (data: CellData): keyof ValueKinds =>
  typeof data.value === "number"
    ? "numbers"
    : typeof data.value === "boolean"
      ? "logicals"
      : isError(data.value)
        ? "errors"
        : "text"

/** The references a formula makes to cells of its own sheet. */
function sameSheetRefs(input: string, sheetName: string): Range[] {
  const parsed = parseFormula(input.slice(1))
  if (!parsed.ok) return []
  const out: Range[] = []
  const walk = (n: Node) => {
    switch (n.type) {
      case "ref":
        add(n)
        break
      case "func":
        n.args.forEach(walk)
        break
      case "unary":
      case "percent":
      case "paren":
        walk(n.arg)
        break
      case "binary":
        walk(n.left)
        walk(n.right)
        break
      case "array":
        n.rows.forEach((r) => r.forEach(walk))
        break
      case "spill":
        walk(n.ref)
        break
    }
  }
  const add = (n: RefNode) => {
    if (n.invalid || n.sheetTo) return
    if (n.sheet !== null && n.sheet.toUpperCase() !== sheetName.toUpperCase()) return
    const to = n.to ?? n.from
    const r0 = Math.min(n.from.row ?? 0, to.row ?? 0)
    const r1 = n.from.row === null ? MAX_ROWS - 1 : Math.max(n.from.row, to.row ?? n.from.row)
    const c0 = Math.min(n.from.col ?? 0, to.col ?? 0)
    const c1 = n.from.col === null ? MAX_COLS - 1 : Math.max(n.from.col, to.col ?? n.from.col)
    out.push({ r0, r1, c0, c1 })
  }
  walk(parsed.ast)
  return out
}

const inRange = (r: Range, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1

/**
 * The cells Go To Special picks within `areas` (the selection when it's more than one cell, else the used range).
 * Current region, current array, row and column differences, precedents and dependents work from the active cell.
 */
export function specialCells(
  sheet: Sheet,
  styles: StyleTable,
  special: SpecialKind,
  selection: readonly Range[],
  active: CellAt,
): Range[] {
  const areas = searchAreas(sheet, selection)
  const used = sheet.cells.usedBounds()
  const pick = (test: (row: number, col: number, data: CellData) => boolean) => {
    const cells: CellAt[] = []
    for (const area of areas)
      sheet.cells.forEach(area, (row, col, data) => test(row, col, data) && cells.push({ row, col }))
    return cellsToRanges(cells)
  }
  switch (special.kind) {
    case "notes":
      return cellsToRanges(candidates(sheet, areas, true).filter((c) => sheet.annotations.get(c.row, c.col)))
    case "constants":
      return pick(
        (_r, _c, d) => d.value !== undefined && !d.spill && !d.input?.startsWith("=") && special.types[typeOf(d)],
      )
    case "formulas":
      return pick(
        (_r, _c, d) =>
          d.value !== undefined && !d.spill && Boolean(d.input?.startsWith("=")) && special.types[typeOf(d)],
      )
    case "blanks": {
      if (!used) return []
      const cells: CellAt[] = []
      for (const area of areas) {
        const r = {
          r0: Math.max(area.r0, 0),
          r1: Math.min(area.r1, used.r1),
          c0: Math.max(area.c0, 0),
          c1: Math.min(area.c1, used.c1),
        }
        for (let row = r.r0; row <= r.r1; row++)
          for (let col = r.c0; col <= r.c1; col++) {
            const v = sheet.cells.value(row, col)
            if (v === undefined) cells.push({ row, col })
          }
      }
      return cellsToRanges(cells)
    }
    case "currentRegion":
      return [currentRegion(sheet, active)]
    case "currentArray": {
      // The dynamic array the active cell belongs to: the formula and the cells it spilled into.
      const extent = (row: number, col: number): Range => {
        let r1 = row
        let c1 = col
        while (r1 + 1 < MAX_ROWS && sheet.cells.isSpill(r1 + 1, col)) r1++
        while (c1 + 1 < MAX_COLS && sheet.cells.isSpill(row, c1 + 1)) c1++
        return { r0: row, c0: col, r1, c1 }
      }
      const own = sheet.cells.get(active.row, active.col)
      if (own?.input?.startsWith("=") && !own.spill) return [extent(active.row, active.col)]
      if (!own?.spill || !used) return []
      let hit: Range | null = null
      sheet.cells.forEach({ r0: used.r0, c0: used.c0, r1: active.row, c1: active.col }, (row, col, d) => {
        if (hit || d.spill || !d.input?.startsWith("=")) return
        const r = extent(row, col)
        if (inRange(r, active.row, active.col)) hit = r
      })
      return hit ? [hit] : []
    }
    case "rowDifferences":
    case "columnDifferences":
      return differences(sheet, styles, areas, active, special.kind === "rowDifferences")
    case "precedents": {
      const found: Range[] = []
      const queue: CellAt[] = [active]
      const seen = new Set<string>([`${active.row},${active.col}`])
      while (queue.length) {
        const at = queue.shift()!
        const input = sheet.cells.input(at.row, at.col)
        if (!input?.startsWith("=")) continue
        for (const r of sameSheetRefs(input, sheet.name)) {
          found.push(r)
          if (!special.all) continue
          sheet.cells.forEach(r, (row, col) => {
            const key = `${row},${col}`
            if (!seen.has(key)) {
              seen.add(key)
              queue.push({ row, col })
            }
          })
        }
        if (!special.all) break
      }
      return found
    }
    case "dependents": {
      if (!used) return []
      const formulas: { at: CellAt; refs: Range[] }[] = []
      sheet.cells.forEach(used, (row, col, d) => {
        if (d.input?.startsWith("=")) formulas.push({ at: { row, col }, refs: sameSheetRefs(d.input, sheet.name) })
      })
      const found = new Map<string, CellAt>()
      let wave: CellAt[] = [active]
      while (wave.length) {
        const next: CellAt[] = []
        for (const f of formulas) {
          const key = `${f.at.row},${f.at.col}`
          if (found.has(key)) continue
          if (wave.some((w) => f.refs.some((r) => inRange(r, w.row, w.col)))) {
            found.set(key, f.at)
            next.push(f.at)
          }
        }
        wave = special.all ? next : []
      }
      return cellsToRanges([...found.values()])
    }
    case "lastCell": {
      const b = used ?? { r1: 0, c1: 0 }
      return [{ r0: b.r1, c0: b.c1, r1: b.r1, c1: b.c1 }]
    }
    case "visible": {
      const out: Range[] = []
      for (const area of areas.length ? selection : []) {
        const rows = runs(area.r0, area.r1, (r) => !sheet.rows.isHidden(r))
        const cols = runs(area.c0, area.c1, (c) => !sheet.cols.isHidden(c))
        for (const [r0, r1] of rows) for (const [c0, c1] of cols) out.push({ r0, r1, c0, c1 })
      }
      return out
    }
    case "conditionalFormats":
    case "validation": {
      let entries: { ranges: Range[] }[] = special.kind === "validation" ? sheet.validations : sheet.conditional
      if (special.same && special.kind === "validation") {
        // Same: the cells whose rule is the active cell's (Excel compares the settings, not the rule's identity).
        const own = validationAt(sheet, active.row, active.col)
        entries = own ? sheet.validations.filter((e) => sameValidation(e.rule, own.rule)) : []
      } else if (special.same) entries = entries.filter((e) => e.ranges.some((r) => inRange(r, active.row, active.col)))
      const whole = selection.length === 1 && selection[0].r0 === selection[0].r1 && selection[0].c0 === selection[0].c1
      const out: Range[] = []
      for (const e of entries)
        for (const r of e.ranges) {
          if (whole) out.push({ ...r })
          else
            for (const a of selection) {
              const x = {
                r0: Math.max(a.r0, r.r0),
                r1: Math.min(a.r1, r.r1),
                c0: Math.max(a.c0, r.c0),
                c1: Math.min(a.c1, r.c1),
              }
              if (x.r0 <= x.r1 && x.c0 <= x.c1) out.push(x)
            }
        }
      return out
    }
  }
}

function runs(from: number, to: number, keep: (i: number) => boolean): [number, number][] {
  const out: [number, number][] = []
  let start = -1
  for (let i = from; i <= to + 1; i++) {
    const ok = i <= to && keep(i)
    if (ok && start < 0) start = i
    if (!ok && start >= 0) {
      out.push([start, i - 1])
      start = -1
    }
  }
  return out
}

/**
 * Row differences: in each row of the selection, the cells that differ from the cell in the active cell's column
 * (column differences: in each column, from the active cell's row). Formulas compare by their pattern (as if
 * written relative to their own cell), values by value.
 */
function differences(
  sheet: Sheet,
  styles: StyleTable,
  areas: readonly Range[],
  active: CellAt,
  rows: boolean,
): Range[] {
  void styles
  const key = (row: number, col: number) => {
    const d = sheet.cells.get(row, col)
    if (!d || d.value === undefined) return "\u0000blank"
    if (d.input?.startsWith("=")) return "f" + relativeFormula(d.input, row, col)
    return "v" + typeof d.value + ":" + String(d.value)
  }
  const cells: CellAt[] = []
  for (const area of areas) {
    if (rows)
      for (let row = area.r0; row <= area.r1; row++) {
        const base = key(row, active.col)
        for (let col = area.c0; col <= area.c1; col++)
          if (col !== active.col && key(row, col) !== base) cells.push({ row, col })
      }
    else
      for (let col = area.c0; col <= area.c1; col++) {
        const base = key(active.row, col)
        for (let row = area.r0; row <= area.r1; row++)
          if (row !== active.row && key(row, col) !== base) cells.push({ row, col })
      }
  }
  return cellsToRanges(cells)
}

/** A formula with its relative references written as offsets from its cell (R1C1-like), for comparing patterns. */
function relativeFormula(input: string, row: number, col: number): string {
  return input.replace(
    /(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
    (_m, ca: string, letters: string, ra: string, digits: string) => {
      let c = 0
      for (const ch of letters.toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64
      const cPart = ca ? `C${c - 1}` : `C[${c - 1 - col}]`
      const rPart = ra ? `R${Number(digits) - 1}` : `R[${Number(digits) - 1 - row}]`
      return rPart + cPart
    },
  )
}

/** Cells as few rectangles: runs along each row, then runs of equal runs down the rows. */
export function cellsToRanges(cells: readonly CellAt[]): Range[] {
  const byRow = new Map<number, number[]>()
  for (const c of cells) {
    const list = byRow.get(c.row)
    if (list) list.push(c.col)
    else byRow.set(c.row, [c.col])
  }
  const open = new Map<string, Range>()
  const out: Range[] = []
  for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
    const cols = [...new Set(byRow.get(row))].sort((a, b) => a - b)
    const spans: [number, number][] = []
    for (const c of cols) {
      const last = spans[spans.length - 1]
      if (last && last[1] === c - 1) last[1] = c
      else spans.push([c, c])
    }
    const seen = new Set<string>()
    for (const [c0, c1] of spans) {
      const k = `${c0}:${c1}`
      seen.add(k)
      const r = open.get(k)
      if (r && r.r1 === row - 1) r.r1 = row
      else {
        const range = { r0: row, r1: row, c0, c1 }
        open.set(k, range)
        out.push(range)
      }
    }
  }
  return out.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0)
}
