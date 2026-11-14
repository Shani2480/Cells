import {
  parseRange,
  parseStructured,
  printStructured,
  type Range,
  type Table,
  type TableColumn,
  type TotalsFunction,
  tokenize,
} from "@cells/core"
import { escapeXmlText } from "../xml"
import { autoFilterXml, readAutoFilter } from "./autofilter"
import { NS_MAIN } from "./constants"
import { at, attrText, kid, kids, parseTree } from "./dom"
import type { StylesIn, StylesOut } from "./styles"

/** A table part (`xl/tables/tableN.xml`), as Excel writes them. */

const FUNCTIONS = new Set<TotalsFunction>([
  "sum",
  "min",
  "max",
  "average",
  "count",
  "countNums",
  "stdDev",
  "var",
  "custom",
])
const flag = (v: string | undefined, fallback: boolean) => (v === undefined ? fallback : v === "1" || v === "true")

/** A table part's table (null when it doesn't read); its formulas are still as the file writes them. */
export function readTable(xml: string, styles: StylesIn | null, warnings: string[]): Table | null {
  const root = parseTree(xml).root
  const range = parseRange(at(root, "ref") ?? "")
  const name = at(root, "displayName") ?? at(root, "name")
  if (!range || !name) return null
  const header = at(root, "headerRowCount") !== "0"
  const totals = Number(at(root, "totalsRowCount") ?? 0) > 0
  const columns: TableColumn[] = kids(kid(root, "tableColumns"), "tableColumn").map((c) => {
    const col: TableColumn = { name: at(c, "name") ?? "" }
    const fn = at(c, "totalsRowFunction") as TotalsFunction | undefined
    if (fn && FUNCTIONS.has(fn)) col.totals = fn
    const label = at(c, "totalsRowLabel")
    if (label) col.totalsLabel = label
    const formula = kid(c, "calculatedColumnFormula")?.text
    if (formula) col.formula = formula
    return col
  })
  if (columns.length !== range.c1 - range.c0 + 1) return null
  const info = kid(root, "tableStyleInfo")
  const filterNode = kid(root, "autoFilter")
  return {
    id: Number(at(root, "id") ?? 0),
    name,
    range,
    header,
    totals,
    columns,
    style: info ? (at(info, "name") ?? null) : null,
    firstColumn: flag(at(info, "showFirstColumn"), false),
    lastColumn: flag(at(info, "showLastColumn"), false),
    rowStripes: flag(at(info, "showRowStripes"), true),
    columnStripes: flag(at(info, "showColumnStripes"), false),
    autoFilter: filterNode ? readAutoFilter(filterNode, styles, warnings) : null,
  }
}

/** `xl/tables/tableN.xml` for a table. Formulas go through `toFile` (structured references spelled out). */
export function tableXml(
  t: Table,
  styles: StylesOut,
  rangeRef: (r: Range) => string,
  toFile: (formula: string) => string,
): string {
  const cols = t.columns
    .map((c, i) => {
      const attrs = attrText({
        id: i + 1,
        name: c.name,
        totalsRowFunction: c.totals && c.totals !== "none" ? c.totals : undefined,
        totalsRowLabel: !c.totals || c.totals === "none" ? c.totalsLabel || undefined : undefined,
      })
      const body = c.formula
        ? `<calculatedColumnFormula>${escapeXmlText(toFile(c.formula))}</calculatedColumnFormula>`
        : ""
      return body ? `<tableColumn${attrs}>${body}</tableColumn>` : `<tableColumn${attrs}/>`
    })
    .join("")
  const filter =
    t.header && t.autoFilter
      ? autoFilterXml({ ...t.autoFilter, range: { ...t.range, r1: t.range.r1 - (t.totals ? 1 : 0) } }, styles, rangeRef)
      : ""
  const style = attrText({
    name: t.style ?? undefined,
    showFirstColumn: t.firstColumn ? "1" : "0",
    showLastColumn: t.lastColumn ? "1" : "0",
    showRowStripes: t.rowStripes ? "1" : "0",
    showColumnStripes: t.columnStripes ? "1" : "0",
  })
  const root = attrText({
    id: t.id,
    name: t.name,
    displayName: t.name,
    ref: rangeRef(t.range),
    headerRowCount: t.header ? undefined : "0",
    totalsRowCount: t.totals ? "1" : undefined,
    totalsRowShown: t.totals ? undefined : "0",
  })
  return (
    `<table xmlns="${NS_MAIN}"${root}>${filter}<tableColumns count="${t.columns.length}">${cols}</tableColumns>` +
    `<tableStyleInfo${style}/></table>`
  )
}

/**
 * A formula's structured references as Excel shows them: [@Qty] for this row, and no table name for the formula's
 * own table (the file always writes Sales[[#This Row],[Qty]]).
 */
export function structuredFromFile(formula: string, own: string | null): string {
  if (!formula.includes("[")) return formula
  return rewrite(formula, (text) => {
    const spec = parseStructured(text)
    if (!spec || spec.table === null) return text
    const mine = own !== null && spec.table.toUpperCase() === own.toUpperCase()
    return printStructured(spec, spec.table, mine)
  })
}

/** The other way: every reference names its table, and this row is #This Row. */
export function structuredToFile(formula: string, own: string | null): string {
  if (!formula.includes("[")) return formula
  return rewrite(formula, (text) => {
    const spec = parseStructured(text)
    if (!spec) return text
    const table = spec.table ?? own
    if (!table) return text
    return printStructured(spec, table, false, true)
  })
}

function rewrite(formula: string, change: (text: string) => string): string {
  const eq = formula.startsWith("=")
  let out = eq ? "=" : ""
  for (const t of tokenize(eq ? formula.slice(1) : formula)) out += t.kind === "structured" ? change(t.text) : t.text
  return out
}
