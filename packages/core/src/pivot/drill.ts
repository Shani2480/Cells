import type { PivotCellItem } from "./engine"
import { VALUES_FIELD, type PivotCache, type PivotTable } from "./model"

/**
 * Expand and collapse (the +/− buttons, double-clicking an item, the cell menu's Expand/Collapse): which items a
 * command changes, and the PivotTable with their collapsed flags set. Laying the cells out again is the caller's.
 */

/** The fields on a label's axis that sit inside its level (the Values field aside). */
export function innerFields(table: PivotTable, item: PivotCellItem): number[] {
  const fields = item.axis === "row" ? table.rows : table.cols
  return fields.slice(item.level + 1).filter((f) => f !== VALUES_FIELD)
}

/** The field of the level above a label's, on its axis (the Values field aside), or null at the outermost level. */
export function outerField(table: PivotTable, item: PivotCellItem): number | null {
  const fields = item.axis === "row" ? table.rows : table.cols
  const above = fields.slice(0, item.level).filter((f) => f !== VALUES_FIELD)
  return above.length ? above[above.length - 1] : null
}

/**
 * The table with items of `field` collapsed or expanded: those with cache items `xs`, or all of the field's (and the
 * cache's items it doesn't list yet) when `xs` is null. The same object when nothing changes.
 */
export function withCollapsed(
  table: PivotTable,
  cache: PivotCache,
  field: number,
  xs: readonly number[] | null,
  collapsed: boolean,
): PivotTable {
  const f = table.fields[field]
  if (!f) return table
  const wanted = xs ? new Set(xs) : null
  const touches = (x: number) => !wanted || wanted.has(x)
  let changed = false
  const items = f.items.map((item) => {
    if (!touches(item.x) || Boolean(item.collapsed) === collapsed) return item
    changed = true
    const next = { ...item }
    if (collapsed) next.collapsed = true
    else delete next.collapsed
    return next
  })
  if (collapsed) {
    // Items the field doesn't list yet (shared items added since) join it, collapsed.
    const listed = new Set(items.map((i) => i.x))
    const count = cache.fields[field]?.items?.length ?? 0
    const add = wanted ? [...wanted] : Array.from({ length: count }, (_, x) => x)
    for (const x of add)
      if (!listed.has(x) && x >= 0 && x < count) {
        items.push({ x, collapsed: true })
        changed = true
      }
  }
  if (!changed) return table
  const fields = [...table.fields]
  fields[field] = { ...f, items }
  return { ...table, fields }
}
