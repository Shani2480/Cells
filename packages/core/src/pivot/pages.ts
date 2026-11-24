import type { CellValue } from "../values"
import { pivotChoices } from "./filter"
import type { PivotCache, PivotTable } from "./model"

/**
 * PivotTable Analyze › Options › Show Report Filter Pages: the items of report filter `f` that get a sheet each, in
 * the filter's order. With Select Multiple Items on, only the ticked ones (Excel's); otherwise every item.
 */
export function reportFilterPages(
  table: PivotTable,
  cache: PivotCache,
  f: number,
  format?: (v: CellValue, numFmt: string | undefined) => string,
): { x: number; label: string }[] {
  if (!table.pages.some((p) => p.field === f)) return []
  const choices = pivotChoices(table, cache, f, format)
  return choices.items.filter((i) => !choices.multipleItems || i.checked).map(({ x, label }) => ({ x, label }))
}

/**
 * The name of a report filter page's sheet: the item's label, or `fallback` (the new sheet's own "SheetN") when the
 * label can't name a sheet (longer than 31 characters, or with : \ / ? * [ ]); a name already taken gets " (2)",
 * " (3)"… as Excel gives.
 */
export function reportPageSheetName(label: string, taken: readonly string[], fallback: string): string {
  const bad = (n: string) => !n || n.length > 31 || /[:\\/?*[\]]/.test(n) || n.startsWith("'") || n.endsWith("'")
  if (bad(label)) return fallback
  const used = new Set(taken.map((n) => n.toLowerCase()))
  if (!used.has(label.toLowerCase())) return label
  for (let k = 2; ; k++) {
    const name = `${label} (${k})`
    if (bad(name)) return fallback
    if (!used.has(name.toLowerCase())) return name
  }
}
