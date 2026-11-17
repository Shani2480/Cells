import {
  copySparklineGroups,
  newSparklineGroup,
  sparkStyleColors,
  sparkStyleIndex,
  sparklinesFor,
  sparkRefText,
  type Op,
  type Range,
  type SparkAxisScale,
  type SparkColor,
  type SparkColors,
  type Sparkline,
  type SparklineGroup,
  type SparklineType,
} from "@cells/core"
import type { SheetController } from "./controller"
import { parseSheetRange } from "@/shell/AdvancedFilterDialog"

/** The Show group's check boxes (and Marker Color's entries). */
export type SparkFlag = "high" | "low" | "negative" | "first" | "last" | "markers"

/** What Create / Edit Sparklines opens with. */
export type SparklineDialogInit =
  | { kind: "create"; type: SparklineType; data: string; location: string }
  | { kind: "editGroup"; data: string; location: string }
  | { kind: "editSingle"; data: string; cell: string }

export const TYPE_LABEL: Record<SparklineType, string> = { line: "Line", column: "Column", stacked: "Win/Loss" }

/**
 * Insert › Sparklines and the Sparkline tab: sparklines on the active sheet, the groups the selection touches, and
 * every change Excel's tab makes to them (each one undo step).
 */
export class SparklineCommands {
  constructor(private readonly c: SheetController) {}

  private get groups(): SparklineGroup[] {
    return this.c.sheet.sparklines
  }

  /** The group of the active cell's sparkline, or null (the Sparkline tab shows while there is one). */
  active(): SparklineGroup | null {
    const { row, col } = this.c.selection.active
    return this.groups.find((g) => g.sparklines.some((s) => s.row === row && s.col === col)) ?? null
  }

  private inSelection(s: Sparkline) {
    return this.c.selection.ranges.some((r) => s.row >= r.r0 && s.row <= r.r1 && s.col >= r.c0 && s.col <= r.c1)
  }

  /** The groups with a sparkline in the selection (what the tab's commands change, as in Excel). */
  selectedGroups(): SparklineGroup[] {
    return this.groups.filter((g) => g.sparklines.some((s) => this.inSelection(s)))
  }

  /** Writes the sheet's groups, `change` applied to copies of the selected ones. */
  private edit(label: string, change: (g: SparklineGroup) => void): boolean {
    const chosen = new Set(this.selectedGroups())
    if (!chosen.size) return false
    const groups = this.groups.map((g) => {
      if (!chosen.has(g)) return g
      const copy = copySparklineGroups([g])[0]
      change(copy)
      return copy
    })
    return this.commit(groups, label)
  }

  private commit(groups: SparklineGroup[], label: string): boolean {
    const sel = this.c.selection
    return (
      this.c.transact([{ kind: "sparklines", sheet: this.c.sheetIndex, groups }], {
        label,
        before: sel,
        after: sel,
      }) !== null
    )
  }

  // ── Insert › Sparklines and Edit Data ──

  /** A range as the dialogs show it: "A2:E4" on this sheet, "Data!A2:E4" on another. */
  private rangeText(range: Range, sheet = this.c.sheetIndex) {
    const text = sparkRefText(this.c.wb.sheets[sheet].name, range)
    return sheet === this.c.sheetIndex ? text.slice(text.lastIndexOf("!") + 1) : text
  }

  /** Create Sparklines opens with the selection as the data (Excel's way: the location is then typed or picked). */
  openCreate(type: SparklineType) {
    const sel = this.c.selection.ranges[0]
    this.c.shell.sparklineDialog?.({ kind: "create", type, data: this.rangeText(sel), location: "" })
  }

  openEditGroup() {
    const g = this.active()
    if (!g) return
    const loc = boundsOf(g.sparklines.map((s) => ({ r0: s.row, c0: s.col, r1: s.row, c1: s.col })))
    const data = boundsOf(
      g.sparklines.flatMap((s) => {
        const at = parseSheetRange(this.c, s.source)
        return at ? [at.range] : []
      }),
    )
    const first = g.sparklines.map((s) => parseSheetRange(this.c, s.source)).find(Boolean)
    this.c.shell.sparklineDialog?.({
      kind: "editGroup",
      data: data && first ? this.rangeText(data, first.sheet) : "",
      location: loc ? this.rangeText(loc) : "",
    })
  }

  openEditSingle() {
    const { row, col } = this.c.selection.active
    const s = this.active()?.sparklines.find((x) => x.row === row && x.col === col)
    if (!s) return
    const at = parseSheetRange(this.c, s.source)
    this.c.shell.sparklineDialog?.({
      kind: "editSingle",
      data: at ? this.rangeText(at.range, at.sheet) : s.source,
      cell: this.rangeText({ r0: row, c0: col, r1: row, c1: col }),
    })
  }

  /** The sparklines for typed data and location ranges, or Excel's message. */
  private build(dataText: string, locationText: string): Sparkline[] | string {
    const data = parseSheetRange(this.c, dataText)
    if (!data) return "Data source reference is not valid."
    const location = parseSheetRange(this.c, locationText)
    if (!location || location.sheet !== this.c.sheetIndex)
      return "Location reference is not valid because the cells are not all in the same column or same row. Select cells that are all in a single row or column."
    const made = sparklinesFor(this.c.wb.sheets[data.sheet].name, data.range, location.range)
    return "error" in made ? made.error : made.sparklines
  }

  /** Create Sparklines: a new group; sparklines already in its cells leave their old groups. */
  create(type: SparklineType, dataText: string, locationText: string): boolean {
    const made = this.build(dataText, locationText)
    if (typeof made === "string") {
      this.c.shell.alert(made)
      return false
    }
    const groups = withoutCells(this.groups, made)
    groups.push(newSparklineGroup(type, made))
    const ok = this.commit(groups, "Insert Sparklines")
    if (ok) {
      const s = made[0]
      this.c.goTo({ r0: s.row, c0: s.col, r1: s.row, c1: s.col })
      this.c.shell.showTab?.("TabSparkline")
    }
    return ok
  }

  /**
   * The change that adds a group of sparklines of `data` (one per row or column) in `location` on this sheet (Quick
   * Analysis previews it), or Excel's message.
   */
  createOps(type: SparklineType, data: Range, location: Range): Op[] | string {
    const made = sparklinesFor(this.c.sheet.name, data, location)
    if ("error" in made) return made.error
    const groups = withoutCells(this.groups, made.sparklines)
    groups.push(newSparklineGroup(type, made.sparklines))
    return [{ kind: "sparklines", sheet: this.c.sheetIndex, groups }]
  }

  /** Edit Group Location & Data: the active group's sparklines made anew from the ranges, its options kept. */
  editGroup(dataText: string, locationText: string): boolean {
    const g = this.active()
    if (!g) return false
    const made = this.build(dataText, locationText)
    if (typeof made === "string") {
      this.c.shell.alert(made)
      return false
    }
    const groups = withoutCells(
      this.groups.filter((x) => x !== g),
      made,
    )
    groups.push({ ...copySparklineGroups([g])[0], sparklines: made })
    return this.commit(groups, "Edit Sparklines")
  }

  /** Edit Single Sparkline's Data. */
  editSingle(dataText: string): boolean {
    const data = parseSheetRange(this.c, dataText)
    const { row, col } = this.c.selection.active
    if (!data || (data.range.r0 !== data.range.r1 && data.range.c0 !== data.range.c1)) {
      this.c.shell.alert("Data source reference is not valid.")
      return false
    }
    const source = sparkRefText(this.c.wb.sheets[data.sheet].name, data.range)
    const g = this.active()
    if (!g) return false
    const groups = this.groups.map((x) =>
      x === g
        ? {
            ...copySparklineGroups([x])[0],
            sparklines: x.sparklines.map((s) => (s.row === row && s.col === col ? { ...s, source } : { ...s })),
          }
        : x,
    )
    return this.commit(groups, "Edit Sparkline Data")
  }

  // ── Type, Show, Style ──

  type(): SparklineType | null {
    return this.active()?.type ?? null
  }

  setType(type: SparklineType) {
    this.edit("Sparkline Type", (g) => (g.type = type))
  }

  flag(key: SparkFlag): boolean {
    return Boolean(this.active()?.[key])
  }

  toggleFlag(key: SparkFlag) {
    const on = !this.flag(key)
    this.edit("Sparkline Show", (g) => (g[key] = on))
  }

  styleIndex(): number {
    const g = this.active()
    return g ? sparkStyleIndex(g.colors) : -1
  }

  setStyle(index: number) {
    this.edit("Sparkline Style", (g) => (g.colors = { ...sparkStyleColors(index), axis: g.colors.axis }))
  }

  color(key: keyof SparkColors): SparkColor | null {
    return this.active()?.colors[key] ?? null
  }

  setColor(key: keyof SparkColors, color: SparkColor) {
    this.edit("Sparkline Color", (g) => (g.colors[key] = color))
  }

  weight(): number {
    return this.active()?.lineWeight ?? 0.75
  }

  setWeight(points: number) {
    this.edit("Sparkline Weight", (g) => (g.lineWeight = points))
  }

  // ── Hidden & Empty Cells, Axis ──

  setEmpty(emptyAs: SparklineGroup["emptyAs"], showHidden: boolean) {
    this.edit("Hidden & Empty Cells", (g) => {
      g.emptyAs = emptyAs
      g.showHidden = showHidden
    })
  }

  setAxisOption(patch: Partial<Pick<SparklineGroup, "displayXAxis" | "rightToLeft">>) {
    this.edit("Sparkline Axis", (g) => Object.assign(g, patch))
  }

  setScale(which: "min" | "max", scale: SparkAxisScale, value?: number) {
    this.edit("Sparkline Axis", (g) => {
      if (which === "min") {
        g.minAxis = scale
        if (scale === "custom") g.manualMin = value ?? g.manualMin ?? 0
      } else {
        g.maxAxis = scale
        if (scale === "custom") g.manualMax = value ?? g.manualMax ?? 0
      }
    })
  }

  /** Date Axis Type with a range of dates (null: back to General Axis Type). */
  setDateRange(text: string | null): boolean {
    if (text === null) return this.edit("Sparkline Axis", (g) => delete g.dateRange)
    const at = parseSheetRange(this.c, text)
    if (!at || (at.range.r0 !== at.range.r1 && at.range.c0 !== at.range.c1)) {
      this.c.shell.alert("The date range isn't valid. Select a single row or column of dates.")
      return false
    }
    const ref = sparkRefText(this.c.wb.sheets[at.sheet].name, at.range)
    return this.edit("Sparkline Axis", (g) => (g.dateRange = ref))
  }

  /** The active group's date range for the dialog, as typed ("A1:E1", "Sheet2!A1:E1"). */
  dateRangeText(): string {
    const ref = this.active()?.dateRange
    const at = ref ? parseSheetRange(this.c, ref) : null
    return at ? this.rangeText(at.range, at.sheet) : ""
  }

  // ── Group, Ungroup, Clear ──

  /** Group: the selected sparklines become one group, with the active cell's group's look. */
  group() {
    const chosen = this.selectedGroups()
    const lead = this.active() ?? chosen[0]
    if (!lead || chosen.length < 2) return
    const picked = chosen.flatMap((g) => g.sparklines.filter((s) => this.inSelection(s)))
    const groups = withoutCells(this.groups, picked)
    groups.push({ ...copySparklineGroups([lead])[0], sparklines: picked.map((s) => ({ ...s })) })
    this.commit(groups, "Group Sparklines")
  }

  /** Ungroup: every selected sparkline its own group, keeping its look. */
  ungroup() {
    const chosen = this.selectedGroups()
    if (!chosen.length) return
    const groups: SparklineGroup[] = []
    for (const g of this.groups) {
      if (!chosen.includes(g)) {
        groups.push(g)
        continue
      }
      for (const s of g.sparklines) groups.push({ ...copySparklineGroups([g])[0], sparklines: [{ ...s }] })
    }
    this.commit(groups, "Ungroup Sparklines")
  }

  canUngroup(): boolean {
    return this.selectedGroups().some((g) => g.sparklines.length > 1)
  }

  canGroup(): boolean {
    return this.selectedGroups().length > 1
  }

  /** Clear Selected Sparklines (or, `whole`, Clear Selected Sparkline Groups). */
  clear(whole = false) {
    const chosen = new Set(this.selectedGroups())
    if (!chosen.size) return
    const groups = whole
      ? this.groups.filter((g) => !chosen.has(g))
      : this.groups.flatMap((g) => {
          if (!chosen.has(g)) return [g]
          const left = g.sparklines.filter((s) => !this.inSelection(s))
          return left.length ? [{ ...copySparklineGroups([g])[0], sparklines: left.map((s) => ({ ...s })) }] : []
        })
    this.commit(groups, "Clear Sparklines")
  }
}

/** The groups without sparklines in `cells` (a new group takes them over); groups left empty go. */
function withoutCells(groups: readonly SparklineGroup[], cells: readonly Sparkline[]): SparklineGroup[] {
  const taken = new Set(cells.map((s) => `${s.row},${s.col}`))
  return groups.flatMap((g) => {
    const left = g.sparklines.filter((s) => !taken.has(`${s.row},${s.col}`))
    if (left.length === g.sparklines.length) return [g]
    return left.length ? [{ ...copySparklineGroups([g])[0], sparklines: left.map((s) => ({ ...s })) }] : []
  })
}

function boundsOf(ranges: Range[]): Range | null {
  if (!ranges.length) return null
  return ranges.reduce((b, r) => ({
    r0: Math.min(b.r0, r.r0),
    c0: Math.min(b.c0, r.c0),
    r1: Math.max(b.r1, r.r1),
    c1: Math.max(b.c1, r.c1),
  }))
}
