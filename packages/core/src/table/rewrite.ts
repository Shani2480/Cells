import { columnName } from "../address"
import { tokenize } from "../formula/lexer"
import { quoteSheet } from "../formula/print"
import type { SheetList } from "../names"
import type { CellWrite, Op } from "../ops"
import {
  allTables,
  copyTables,
  parseStructured,
  printStructured,
  structuredArea,
  tableAt,
  type StructuredSpec,
  type Table,
} from "./table"

/**
 * Formula changes that follow a table: renaming it or a column (every structured reference that names it follows),
 * and Convert to Range (its references become ordinary ones, as Excel does).
 */

/** A structured reference's replacement text (null: unchanged). `own`: the formula is inside the referenced table. */
export type StructuredChange = (spec: StructuredSpec, at: { sheet: number; row: number; col: number }) => string | null

function rewriteText(input: string, change: (text: string) => string | null): string | null {
  if (!input.includes("[")) return null
  const eq = input.startsWith("=")
  let out = eq ? "=" : ""
  let changed = false
  for (const t of tokenize(eq ? input.slice(1) : input)) {
    if (t.kind !== "structured") {
      out += t.text
      continue
    }
    const next = change(t.text)
    if (next !== null && next !== t.text) changed = true
    out += next ?? t.text
  }
  return changed ? out : null
}

/** The operations that rewrite every formula (cells, calculated columns, defined names) through `change`. */
export function rewriteStructuredOps(sheets: SheetList, change: StructuredChange): Op[] {
  const ops: Op[] = []
  sheets.forEach((sheet, index) => {
    const writes: CellWrite[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      const next = rewriteText(input, (text) => {
        const spec = parseStructured(text)
        return spec ? change(spec, { sheet: index, row, col }) : null
      })
      if (next !== null) writes.push({ row, col, data: { ...sheet.cells.get(row, col)!, input: next } })
    })
    if (writes.length) ops.push({ kind: "setCells", sheet: index, cells: writes })
    let tablesChanged = false
    const tables = copyTables(sheet.tables)
    for (const t of tables)
      for (const c of t.columns) {
        if (!c.formula) continue
        const at = { sheet: index, row: t.range.r0 + (t.header ? 1 : 0), col: t.range.c0 + t.columns.indexOf(c) }
        const next = rewriteText(c.formula, (text) => {
          const spec = parseStructured(text)
          return spec ? change(spec, at) : null
        })
        if (next !== null) {
          c.formula = next
          tablesChanged = true
        }
      }
    if (tablesChanged) ops.push({ kind: "tables", sheet: index, tables })
  })
  const names = sheets.names
  if (names?.all().length) {
    let changed = false
    const next = names.copy().map((n) => {
      const ref = rewriteText(n.ref, (text) => {
        const spec = parseStructured(text)
        return spec ? change(spec, { sheet: -1, row: -1, col: -1 }) : null
      })
      if (ref === null) return n
      changed = true
      return { ...n, ref }
    })
    if (changed) ops.push({ kind: "names", names: next })
  }
  return ops
}

/** One formula with references naming table `from` renamed to `to` (null: unchanged). Copy Sheet uses it. */
export function renameTableInFormula(input: string, from: string, to: string): string | null {
  const upper = from.toUpperCase()
  return rewriteText(input, (text) => {
    const spec = parseStructured(text)
    if (!spec || spec.table?.toUpperCase() !== upper) return null
    return printStructured({ ...spec, table: to }, to, false)
  })
}

/** Which table a reference points at, from where it is written. */
function target(
  sheets: SheetList,
  spec: StructuredSpec,
  at: { sheet: number; row: number; col: number },
): Table | null {
  if (spec.table !== null) {
    const upper = spec.table.toUpperCase()
    return allTables(sheets).find((t) => t.table.name.toUpperCase() === upper)?.table ?? null
  }
  return at.sheet >= 0 ? tableAt(sheets[at.sheet], at.row, at.col) : null
}

/** Rename a table (`newName`) and/or one of its columns (`column`: old and new name). */
export function renameTableOps(
  sheets: SheetList,
  table: Table,
  newName: string,
  column?: { from: string; to: string },
): Op[] {
  const upper = (s: string | null) => s?.toUpperCase()
  return rewriteStructuredOps(sheets, (spec, at) => {
    if (target(sheets, spec, at) !== table) return null
    const next = { ...spec }
    if (spec.table !== null) next.table = newName
    if (column) {
      if (upper(spec.c0) === upper(column.from)) next.c0 = column.to
      if (upper(spec.c1) === upper(column.from)) next.c1 = column.to
    }
    if (next.table === spec.table && next.c0 === spec.c0 && next.c1 === spec.c1) return null
    return printStructured(next, next.table, spec.table === null)
  })
}

/**
 * Convert to Range: references to the table become ordinary ones ([@Qty] → B2, Sales[Qty] → Sales!$B$2:$B$5 from
 * another sheet, $B$2:$B$5 on its own).
 */
export function unstructureOps(sheets: SheetList, table: Table): Op[] {
  const home = sheets.findIndex((s) => s.tables.includes(table))
  return rewriteStructuredOps(sheets, (spec, at) => {
    if (target(sheets, spec, at) !== table) return null
    const row = at.row >= 0 ? at.row : 0
    const area = structuredArea({ ...spec, table: table.name }, [{ sheet: home, table }], home, row, at.col)
    if (typeof area === "string") return "#REF!"
    const abs = !spec.thisRow
    const cell = (r: number, c: number) => `${abs ? "$" : ""}${columnName(c)}${abs ? "$" : ""}${r + 1}`
    let ref = cell(area.r0, area.c0)
    if (area.r1 !== area.r0 || area.c1 !== area.c0) ref += ":" + cell(area.r1, area.c1)
    return at.sheet === home ? ref : `${quoteSheet(sheets[home].name)}!${ref}`
  })
}
