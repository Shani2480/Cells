import {
  isChart,
  slicerFilter,
  slicerItems,
  slicerTarget,
  tableFilterRange,
  type AutoFilter,
  type SheetSlicer,
} from "@cells/core"
import { setSlicerSource, type SlicerView } from "@/drawing/slicer"
import { todaySerial, type SheetController } from "./controller"

export type { SlicerView }

/** How a click picks: this button only, this one on or off (Ctrl, Multi-Select), or a run of them (Shift). */
export type SlicerPick = "only" | "toggle" | "range"

/**
 * Slicers on the sheet: their buttons (from the table they filter) and picking them. A pick is the table's filter on
 * the slicer's column, set as Excel does, so it undoes as one step and the Filter Button's dropdown agrees.
 */
export class SlicerCommands {
  /** Slicers with Multi-Select on (the header's button), by id: a view state, not saved. */
  private readonly multi = new Set<number>()
  /** The button a Shift+click runs from, per slicer. */
  private readonly anchor = new Map<number, string>()
  private cache: { version: number; views: Map<number, SlicerView> } | null = null

  constructor(private readonly c: SheetController) {
    setSlicerSource((o) => this.view(o))
  }

  /** The selected slicer (the Slicer tab's), or null. */
  active(): SheetSlicer | null {
    const list = this.c.objects.drawings()
    return list.length === 1 && list[0].kind === "slicer" ? list[0] : null
  }

  /** The selected slicers (the Slicer tab changes them all). */
  selected(): SheetSlicer[] {
    return this.c.objects.drawings().filter((o): o is SheetSlicer => o.kind === "slicer")
  }

  view(o: SheetSlicer): SlicerView {
    if (!this.cache || this.cache.version !== this.c.wb.version)
      this.cache = { version: this.c.wb.version, views: new Map() }
    let v = this.cache.views.get(o.id)
    if (!v) {
      const target = slicerTarget(this.c.wb.sheets, o)
      const look = target?.sheet === this.c.sheetIndex ? this.c.data.look : undefined
      const filter = target?.table.autoFilter
      const offset = target && filter ? target.col - filter.range.c0 : -1
      v = {
        items: slicerItems(this.c.wb.sheets, this.c.wb.styles, o, todaySerial(), look),
        filtered: Boolean(filter && filter.columns[offset]),
        multi: false,
      }
      this.cache.views.set(o.id, v)
    }
    return { ...v, multi: this.multi.has(o.id) }
  }

  isMulti(o: SheetSlicer): boolean {
    return this.multi.has(o.id)
  }

  /** The header's Multi-Select button (Alt+S): clicks then pick buttons on and off. */
  toggleMulti(o: SheetSlicer) {
    if (this.multi.has(o.id)) this.multi.delete(o.id)
    else this.multi.add(o.id)
    this.c.changed()
  }

  /** A click on a button (`key`, as slicerItems gives it). */
  pick(o: SheetSlicer, key: string, how: SlicerPick): boolean {
    const items = this.view(o).items
    let picked = new Set(items.filter((i) => i.selected).map((i) => i.key))
    const from = this.anchor.get(o.id)
    if (how === "range" && from !== undefined && items.some((i) => i.key === from)) {
      const a = items.findIndex((i) => i.key === from)
      const b = items.findIndex((i) => i.key === key)
      picked = new Set(items.slice(Math.min(a, b), Math.max(a, b) + 1).map((i) => i.key))
    } else if (how === "toggle" || this.multi.has(o.id)) {
      // With nothing filtered, the first toggle leaves just that one out, as in Excel.
      if (picked.has(key)) {
        picked.delete(key)
        // Excel keeps at least one button picked.
        if (!picked.size) return false
      } else picked.add(key)
      this.anchor.set(o.id, key)
    } else {
      picked = new Set([key])
      this.anchor.set(o.id, key)
    }
    return this.apply(o, slicerFilter(items, picked))
  }

  /** The header's Clear Filter button (Alt+C). */
  clear(o: SheetSlicer): boolean {
    this.anchor.delete(o.id)
    return this.apply(o, null)
  }

  private apply(o: SheetSlicer, column: ReturnType<typeof slicerFilter>): boolean {
    const target = slicerTarget(this.c.wb.sheets, o)
    if (!target) return false
    if (this.c.edit) this.c.commitEdit()
    const { table } = target
    const before: AutoFilter = table.autoFilter ?? { range: tableFilterRange(table), columns: {} }
    const offset = target.col - before.range.c0
    const columns = { ...before.columns }
    if (column) columns[offset] = column
    else delete columns[offset]
    if (!column && !before.columns[offset]) return false
    const filter: AutoFilter = { ...before, columns }
    const ops = this.c.data.filterOps({ filter: before, table }, filter, target.sheet)
    const sel = this.c.selection
    return this.c.transact(ops, { label: "Slicer", before: sel, after: sel }) !== null
  }

  /** Scrolls a slicer's buttons (the mouse wheel over it): a view change, kept when saved, not undone. */
  scroll(o: SheetSlicer, rows: number, shownRows: number, totalRows: number) {
    const columns = Math.max(1, o.columns)
    const first = Math.max(0, Math.min(totalRows - shownRows, Math.floor((o.startItem ?? 0) / columns) + rows))
    const next = first * columns
    if (next === (o.startItem ?? 0)) return
    if (next) o.startItem = next
    else delete o.startItem
    this.c.changed()
  }

  /** Every slicer of the workbook (with its sheet). */
  all(): { sheet: number; slicer: SheetSlicer }[] {
    return this.c.wb.sheets.flatMap((s, sheet) =>
      s.objects.flatMap((o) => (!isChart(o) && o.kind === "slicer" ? [{ sheet, slicer: o }] : [])),
    )
  }
}
