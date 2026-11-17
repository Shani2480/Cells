import {
  DEFAULT_SLICER_STYLE,
  isChart,
  nextObjectId,
  SLICER_BUTTON_HEIGHT,
  SLICER_SIZE,
  SLICER_STYLES,
  slicerFilter,
  slicerItems,
  slicerTarget,
  tableFilterRange,
  type AutoFilter,
  type SheetSlicer,
} from "@cells/core"
import { setSlicerSource, type SlicerView } from "@/drawing/slicer"
import { anchorOf } from "./charts"
import { todaySerial, type SheetController } from "./controller"

export type { SlicerView }

/** Slicer Settings: what the dialog changes. */
export type SlicerSettings = Pick<
  SheetSlicer,
  "name" | "caption" | "header" | "sort" | "customLists" | "hideNoData" | "crossFilter"
>

/** Room around the buttons: 9 px each side, 3 px between columns (Excel's layout). */
const SIDES = 18
const GAP = 3

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

  // ── Insert › Slicer and the Slicer tab ──

  /** Insert › Slicer, Table Design › Insert Slicer: the Insert Slicers dialog for the active cell's table. */
  openInsert() {
    if (!this.c.tables.active()) {
      this.c.shell.alert("Select a cell in a table to insert a slicer for it.")
      return
    }
    this.c.shell.openDialog?.("insertSlicers")
  }

  /** The columns the Insert Slicers dialog lists (the active cell's table's). */
  insertColumns(): string[] {
    return this.c.tables.active()?.columns.map((c) => c.name) ?? []
  }

  /** A slicer name not taken in the workbook: the column's, then "Region 1", "Region 2"… */
  private freeName(base: string, taken: Set<string>): string {
    let name = base
    for (let n = 1; taken.has(name.toLowerCase()); n++) name = `${base} ${n}`
    taken.add(name.toLowerCase())
    return name
  }

  /** One slicer per column picked, in the middle of the window, each a little below and right of the last. */
  insert(columns: readonly string[]): SheetSlicer[] {
    const table = this.c.tables.active()
    if (!table || !columns.length) return []
    const taken = new Set(this.all().map((s) => s.slicer.name.toLowerCase()))
    let id = nextObjectId(this.c.wb.sheets)
    const start = this.c.objects.middle(SLICER_SIZE.w, SLICER_SIZE.h)
    const made = columns.map((column, i): SheetSlicer => {
      const box = { ...start, x: start.x + i * 24, y: start.y + i * 24 }
      return {
        kind: "slicer",
        id: id++,
        name: this.freeName(column, taken),
        anchor: anchorOf(this.c.sheet, box),
        placement: "fixed",
        table: table.id,
        column,
        caption: column,
        header: true,
        style: DEFAULT_SLICER_STYLE,
        columns: 1,
        buttonHeight: SLICER_BUTTON_HEIGHT,
        sort: "ascending",
        customLists: true,
        crossFilter: "showItemsWithDataAtTop",
      }
    })
    if (!this.c.objects.commit([...this.c.objects.list, ...made], "Insert Slicer")) return []
    this.c.objects.select(made.map((o) => o.id))
    this.c.shell.showTab?.("TabSlicer")
    return made
  }

  /** A change to every selected slicer, as one undo step. */
  private change(label: string, edit: (o: SheetSlicer) => void): boolean {
    const ids = this.selected().map((o) => o.id)
    return this.c.objects.update(
      label,
      (o) => {
        if (!isChart(o) && o.kind === "slicer") edit(o)
      },
      ids,
    )
  }

  /** Slicer Caption (the tab's box). */
  setCaption(caption: string) {
    return this.change("Slicer Caption", (o) => (o.caption = caption))
  }

  styleIndex(): number {
    return SLICER_STYLES.indexOf(this.active()?.style as (typeof SLICER_STYLES)[number])
  }

  setStyle(style: string) {
    return this.change("Slicer Style", (o) => (o.style = style))
  }

  /** Buttons › Columns (1–20000 in Excel; Cells takes 1–100). */
  setColumns(n: number) {
    if (!Number.isFinite(n) || n < 1) return false
    return this.change("Slicer Columns", (o) => {
      o.columns = Math.min(100, Math.round(n))
      delete o.startItem
    })
  }

  /** Buttons › Height (px). */
  setButtonHeight(px: number) {
    if (!Number.isFinite(px) || px < 4) return false
    return this.change("Button Height", (o) => (o.buttonHeight = Math.round(px)))
  }

  /** A slicer's button width (px), from its width and columns. */
  buttonWidth(o: SheetSlicer): number {
    const w = this.c.objects.box(o).w
    const columns = Math.max(1, o.columns)
    return Math.max(4, (w - SIDES - GAP * (columns - 1)) / columns)
  }

  /** Buttons › Width: the slicer grows or shrinks to fit its buttons. */
  setButtonWidth(px: number) {
    if (!Number.isFinite(px) || px < 4) return false
    return this.change("Button Width", (o) => {
      const box = this.c.objects.box(o)
      const columns = Math.max(1, o.columns)
      const w = px * columns + GAP * (columns - 1) + SIDES
      o.anchor = anchorOf(this.c.sheet, { ...box, w })
    })
  }

  /** Slicer Settings. Returns Excel's message when the name is taken, else null. */
  applySettings(o: SheetSlicer, s: SlicerSettings): string | null {
    const name = s.name.trim()
    if (!name) return "Type a name for the slicer."
    const clash = this.all().some((x) => x.slicer.id !== o.id && x.slicer.name.toLowerCase() === name.toLowerCase())
    if (clash) return "The slicer name is already in use. Type a different name."
    this.c.objects.update(
      "Slicer Settings",
      (x) => {
        if (isChart(x) || x.kind !== "slicer") return
        Object.assign(x, { ...s, name })
        if (!s.hideNoData) delete x.hideNoData
      },
      [o.id],
    )
    return null
  }

  /** Every slicer of the workbook (with its sheet). */
  all(): { sheet: number; slicer: SheetSlicer }[] {
    return this.c.wb.sheets.flatMap((s, sheet) =>
      s.objects.flatMap((o) => (!isChart(o) && o.kind === "slicer" ? [{ sheet, slicer: o }] : [])),
    )
  }
}
