import {
  DEFAULT_TIMELINE_STYLE,
  groupKind,
  inRecords,
  isChart,
  nextObjectId,
  pivotSlicerTarget,
  timelineBounds,
  timelineLabel,
  timelinePeriods,
  timelineSelection,
  TIMELINE_PERIOD_WIDTH,
  TIMELINE_SIZE,
  TIMELINE_STYLES,
  withTimelineDates,
  type Op,
  type PivotCache,
  type PivotTable,
  type SheetObject,
  type SheetTimeline,
  type TimelineLevel,
  type TimelinePeriod,
} from "@cells/core"
import { setTimelineSource, type TimelineView } from "@/drawing/timeline"
import { anchorOf } from "./charts"
import type { SheetController } from "./controller"

export type { TimelineView }

/** Timeline › Show: the parts a timeline shows. */
export type TimelineShowKey = "header" | "selectionLabel" | "scrollbar" | "timeLevel"

/** A PivotTable change: a timeline's dates on field `f` (null: none). */
const datesIn = (f: number, range: { from: number; to: number } | null) => (t: PivotTable) =>
  withTimelineDates(t, f, range)

/**
 * Timelines on the sheet: their periods (from the PivotTables' date field) and picking them. A pick is a whole-day
 * date filter on that field of each connected PivotTable, set as Excel does, so it undoes as one step.
 */
export class TimelineCommands {
  /** The period a Shift+click runs from, per timeline. */
  private readonly anchor = new Map<number, number>()
  private cache: { version: number; views: Map<number, TimelineView & { all: TimelinePeriod[] }> } | null = null

  constructor(private readonly c: SheetController) {
    setTimelineSource((o) => this.view(o))
  }

  /** The selected timeline (the Timeline tab's), or null. */
  active(): SheetTimeline | null {
    const list = this.c.objects.drawings()
    return list.length === 1 && list[0].kind === "timeline" ? list[0] : null
  }

  selected(): SheetTimeline[] {
    return this.c.objects.drawings().filter((o): o is SheetTimeline => o.kind === "timeline")
  }

  private target(o: SheetTimeline) {
    return pivotSlicerTarget(this.c.wb.sheets, this.c.wb.pivotCaches, o)
  }

  view(o: SheetTimeline): TimelineView {
    if (!this.cache || this.cache.version !== this.c.wb.version)
      this.cache = { version: this.c.wb.version, views: new Map() }
    const key = o.id
    const had = this.cache.views.get(key)
    if (had && had.first === this.firstOf(o, had.periods, had.picked)) return had
    const target = this.target(o)
    const bounds = target ? timelineBounds(target.cache, target.field) : null
    const picked = target ? timelineSelection(target) : null
    const periods = bounds ? timelinePeriods(bounds, o.level) : []
    const v = {
      periods,
      picked,
      first: this.firstOf(o, periods, picked),
      label: timelineLabel(picked, o.selectionLevel),
      all: periods,
    }
    this.cache.views.set(key, v)
    return v
  }

  /** The first period shown: the one at the scroll position, else the first picked, else the first. */
  private firstOf(o: SheetTimeline, periods: readonly TimelinePeriod[], picked: { from: number } | null): number {
    const at = o.scrollPosition ?? picked?.from
    if (at === undefined) return 0
    const i = periods.findIndex((p) => at >= p.start && at < p.end)
    return i < 0 ? 0 : i
  }

  /**
   * A click on period `index` picks it; Shift+click (or a drag) picks the run from the last one clicked. The timeline
   * remembers the level it was picked at, for its label.
   */
  pick(o: SheetTimeline, index: number, extend: boolean): boolean {
    const periods = this.view(o).periods
    const p = periods[index]
    if (!p) return false
    const from = extend ? this.anchor.get(o.id) : undefined
    let range = { from: p.start, to: p.end - 1 }
    if (from !== undefined && periods[from]) {
      const a = periods[Math.min(from, index)]
      const b = periods[Math.max(from, index)]
      range = { from: a.start, to: b.end - 1 }
    } else this.anchor.set(o.id, index)
    return this.apply(o, range, "Timeline")
  }

  /** The header's Clear Filter button. */
  clear(o: SheetTimeline): boolean {
    this.anchor.delete(o.id)
    return this.apply(o, null, "Clear Filter")
  }

  private apply(o: SheetTimeline, range: { from: number; to: number } | null, label: string): boolean {
    const target = this.target(o)
    if (!target) return false
    const changes = target.tables.map((t) => ({ ...t, apply: datesIn(target.field, range) }))
    // The label reads at the level the dates were picked at.
    const extra: Op[] = []
    if (range && o.selectionLevel !== o.level) extra.push(this.objectsOp(o, { ...o, selectionLevel: o.level }))
    return this.c.pivots.slicerPick(changes, target.cache, label, extra)
  }

  /** The op that puts `next` in place of timeline `o` on its sheet. */
  private objectsOp(o: SheetTimeline, next: SheetTimeline): Op {
    const sheet = this.c.wb.sheets.findIndex((s) => s.objects.some((x) => x.id === o.id))
    const objects = this.c.wb.sheets[sheet].objects.map((x) => (x.id === o.id ? next : x))
    return { kind: "objects", sheet, objects }
  }

  /** The time level (the dropdown): the periods shown. */
  setLevel(o: SheetTimeline, level: TimelineLevel) {
    if (o.level === level) return false
    return this.c.objects.update(
      "Time Level",
      (x) => {
        if (isChart(x) || x.kind !== "timeline") return
        x.level = level
        delete x.scrollPosition
      },
      [o.id],
    )
  }

  /** Scrolls the strip by `periods` (the arrows, the wheel): a view change, kept when saved, not undone. */
  scroll(o: SheetTimeline, periods: number) {
    const v = this.view(o)
    const fit = Math.max(1, Math.floor((this.c.objects.box(o).w - 16) / TIMELINE_PERIOD_WIDTH[o.level]))
    const first = Math.max(0, Math.min(v.periods.length - fit, v.first + periods))
    if (first === v.first || !v.periods[first]) return
    o.scrollPosition = v.periods[first].start
    this.c.changed()
  }

  // ── Insert › Timeline and the Timeline tab ──

  /** Insert › Timeline, PivotTable Analyze › Insert Timeline: the Insert Timelines dialog for the PivotTable. */
  openInsert() {
    if (!this.c.pivots.active()) {
      this.c.shell.alert("Select a cell in a PivotTable to insert a timeline for it.")
      return
    }
    if (!this.insertFields().length) {
      this.c.shell.alert(
        "This PivotTable doesn't have any fields formatted as dates. To use a timeline, add a field of dates to its data.",
      )
      return
    }
    this.c.shell.openDialog?.("insertTimelines")
  }

  /** The date fields the Insert Timelines dialog lists (the active PivotTable's data's). */
  insertFields(): string[] {
    const at = this.c.pivots.activeWithCache()
    if (!at) return []
    const { cache } = at
    return cache.fields.flatMap((f, i) =>
      inRecords(f, i) && !f.group && (f.dates || groupKind(cache, i) === "dates") ? [f.name] : [],
    )
  }

  /** One timeline per field picked, in the middle of the window, each a little below and right of the last. */
  insert(fields: readonly string[]): SheetTimeline[] {
    const at = this.c.pivots.activeWithCache()
    if (!at || !fields.length) return []
    const taken = new Set(
      this.c.wb.sheets.flatMap((s) => s.objects.flatMap((o) => (isChart(o) ? [] : [o.name.toLowerCase()]))),
    )
    let id = nextObjectId(this.c.wb.sheets)
    const start = this.c.objects.middle(TIMELINE_SIZE.w, TIMELINE_SIZE.h)
    const made = fields.map((field, i): SheetTimeline => {
      let name = field
      for (let n = 1; taken.has(name.toLowerCase()); n++) name = `${field} ${n}`
      taken.add(name.toLowerCase())
      const box = { ...start, x: start.x + i * 24, y: start.y + i * 24 }
      return {
        kind: "timeline",
        id: id++,
        name,
        anchor: anchorOf(this.c.sheet, box),
        placement: "fixed",
        pivot: { cacheId: at.cache.id, field, tables: [{ sheet: this.c.sheet.name, name: at.table.name }] },
        caption: field,
        header: true,
        selectionLabel: true,
        scrollbar: true,
        timeLevel: true,
        level: "months",
        selectionLevel: "months",
        style: DEFAULT_TIMELINE_STYLE,
      }
    })
    if (!this.c.objects.commit([...this.c.objects.list, ...made], "Insert Timeline")) return []
    this.c.objects.select(made.map((o) => o.id))
    this.c.shell.showTab?.("TabTimeline")
    return made
  }

  /** A change to every selected timeline, as one undo step. */
  private change(label: string, edit: (o: SheetTimeline) => void): boolean {
    const ids = this.selected().map((o) => o.id)
    return this.c.objects.update(
      label,
      (o) => {
        if (!isChart(o) && o.kind === "timeline") edit(o)
      },
      ids,
    )
  }

  setCaption(caption: string) {
    return this.change("Timeline Caption", (o) => (o.caption = caption))
  }

  styleIndex(): number {
    return TIMELINE_STYLES.indexOf(this.active()?.style as (typeof TIMELINE_STYLES)[number])
  }

  setStyle(style: string) {
    return this.change("Timeline Style", (o) => (o.style = style))
  }

  /** Timeline › Show: Header, Selection Label, Scrollbar, Time Level. */
  setShown(key: TimelineShowKey, on: boolean) {
    return this.change("Timeline", (o) => (o[key] = on))
  }

  shown(key: TimelineShowKey): boolean {
    return this.active()?.[key] ?? false
  }

  // ── Report Connections ──

  /** The PivotTables a timeline can filter (those of its cache), ticked when connected. */
  reportConnections(o: SheetTimeline): { key: string; name: string; sheet: string; on: boolean }[] {
    const link = o.pivot
    const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase()
    return this.c.wb.sheets.flatMap((s) =>
      s.pivots
        .filter((t) => t.cacheId === link.cacheId)
        .map((t) => ({
          key: `${s.name}!${t.name}`,
          name: t.name,
          sheet: s.name,
          on: link.tables.some((x) => same(x.sheet, s.name) && x.name === t.name),
        })),
    )
  }

  /** Report Connections' OK: the timeline filters the PivotTables `keys`; new ones take its dates, old ones lose them. */
  setReportConnections(o: SheetTimeline, keys: readonly string[]): string | null {
    const picked = this.reportConnections(o).filter((x) => keys.includes(x.key))
    if (!picked.length) return "Select at least one PivotTable for the timeline to filter."
    this.connect([{ timeline: o, tables: picked.map((x) => ({ sheet: x.sheet, name: x.name })) }], "Report Connections")
    return null
  }

  /** Timelines connected to these PivotTables, as one undo step (as SlicerCommands.connect does for slicers). */
  connect(changes: readonly { timeline: SheetTimeline; tables: { sheet: string; name: string }[] }[], label: string) {
    if (!changes.length) return false
    const sheets = this.c.wb.sheets
    const picks: { sheet: number; table: PivotTable; apply: (t: PivotTable) => PivotTable }[] = []
    const bySheet = new Map<number, SheetObject[]>()
    let cache: PivotCache | null = null
    const same = (a: { sheet: string; name: string }, b: { sheet: string; name: string }) =>
      a.sheet.toLowerCase() === b.sheet.toLowerCase() && a.name === b.name
    for (const { timeline, tables } of changes) {
      const where = sheets.findIndex((s) => s.objects.some((x) => x.id === timeline.id))
      const list = bySheet.get(where) ?? [...sheets[where].objects]
      bySheet.set(
        where,
        list.map((x) => (x.id === timeline.id ? { ...timeline, pivot: { ...timeline.pivot, tables } } : x)),
      )
      const target = this.target(timeline)
      if (!target) continue
      cache = target.cache
      const dates = timelineSelection(target)
      const put = (t: { sheet: string; name: string }, range: typeof dates) => {
        const sheet = sheets.findIndex((s) => s.name.toLowerCase() === t.sheet.toLowerCase())
        const table = sheets[sheet]?.pivots.find((p) => p.name === t.name)
        if (table) picks.push({ sheet, table, apply: datesIn(target.field, range) })
      }
      for (const t of tables) if (!timeline.pivot.tables.some((x) => same(x, t))) put(t, dates)
      for (const t of timeline.pivot.tables) if (!tables.some((x) => same(x, t))) put(t, null)
    }
    const extra: Op[] = [...bySheet].map(([sheet, objects]) => ({ kind: "objects", sheet, objects }))
    if (!cache) return this.c.transact(extra, { label, before: this.c.selection, after: this.c.selection }) !== null
    return this.c.pivots.slicerPick(picks, cache, label, extra)
  }
}
