import { columnName, type Range } from "./address"
import { MAX_OUTLINE, type Axis } from "./axis"
import { contentForInput } from "./edit"
import { formulaRefs } from "./find"
import type { AxisEntry, AxisName, CellWrite, Op } from "./ops"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"

/**
 * Data › Outline: groups of rows or columns (Group, Ungroup, Clear Outline), showing and hiding their detail (the
 * outline bar's + and −, its level buttons, Show Detail and Hide Detail) and Subtotal. Each function returns the
 * operations to apply, so one Undo takes the whole change back.
 */

/** One group: entries `start`–`end` at `level` or deeper, summarised by the entry `summary` beside them. */
export type OutlineGroup = { level: number; start: number; end: number; summary: number }

const groupCache = new WeakMap<Axis, { key: string; groups: OutlineGroup[] }>()

/** Every group on the axis, outermost first where they start together; `after`: summaries below / to the right. */
export function outlineGroups(axis: Axis, after: boolean): OutlineGroup[] {
  if (!axis.maxLevel) return []
  const entries = axis.overrides(0, axis.count - 1).filter((e) => e.level)
  const key = `${after}|${entries.map((e) => `${e.index}:${e.level}`).join(",")}`
  const cached = groupCache.get(axis)
  if (cached?.key === key) return cached.groups
  const groups: OutlineGroup[] = []
  const open = new Array<number>(MAX_OUTLINE + 1).fill(-1)
  const close = (level: number, end: number) => {
    const start = open[level]
    groups.push({ level, start, end, summary: after ? end + 1 : start - 1 })
    open[level] = -1
  }
  let previous = -1
  for (const e of entries) {
    // A gap of level-0 entries closes every open group.
    if (e.index !== previous + 1) for (let l = MAX_OUTLINE; l >= 1; l--) if (open[l] >= 0) close(l, previous)
    const level = e.level ?? 0
    for (let l = MAX_OUTLINE; l > level; l--) if (open[l] >= 0) close(l, e.index - 1)
    for (let l = 1; l <= level; l++) if (open[l] < 0) open[l] = e.index
    previous = e.index
  }
  for (let l = MAX_OUTLINE; l >= 1; l--) if (open[l] >= 0) close(l, previous)
  groups.sort((a, b) => a.start - b.start || a.level - b.level)
  groupCache.set(axis, { key, groups })
  return groups
}

/** An entry as it is, with changes: sizes kept, so the axis op only moves what it means to. */
function entry(axis: Axis, index: number, change: Partial<AxisEntry>): AxisEntry {
  return {
    index,
    size: axis.hasCustomSize(index) ? axis.storedSize(index) : null,
    hidden: axis.isHidden(index),
    ...(axis.isAuto(index) ? { auto: true as const } : {}),
    level: axis.level(index),
    ...(axis.isCollapsed(index) ? { collapsed: true as const } : {}),
    ...change,
  }
}

function axisOp(sheet: number, name: AxisName, entries: AxisEntry[]): Op[] {
  return entries.length ? [{ kind: "axis", sheet, axis: name, entries }] : []
}

const axisOf = (sheet: Sheet, name: AxisName) => (name === "rows" ? sheet.rows : sheet.cols)
const summaryAfter = (sheet: Sheet, name: AxisName) =>
  name === "rows" ? sheet.outline.summaryBelow : sheet.outline.summaryRight

/** Group: entries `from`–`to` one level deeper. Null when they are already at Excel's deepest level. */
export function groupOps(sheet: Sheet, index: number, name: AxisName, from: number, to: number): Op[] | null {
  const axis = axisOf(sheet, name)
  const entries: AxisEntry[] = []
  for (let i = from; i <= to; i++) {
    if (axis.level(i) >= MAX_OUTLINE) return null
    entries.push(entry(axis, i, { level: axis.level(i) + 1 }))
  }
  return axisOp(index, name, entries)
}

/**
 * Ungroup: entries `from`–`to` one level up. Detail hidden by a collapsed group that goes away is shown again, and
 * a summary left without a group loses its collapsed mark. Null when nothing there is grouped.
 */
export function ungroupOps(sheet: Sheet, index: number, name: AxisName, from: number, to: number): Op[] | null {
  const axis = axisOf(sheet, name)
  const after = summaryAfter(sheet, name)
  const groups = outlineGroups(axis, after)
  const changes = new Map<number, Partial<AxisEntry>>()
  for (let i = from; i <= to; i++) {
    const level = axis.level(i)
    if (!level) continue
    const collapsedBy = groups.some(
      (g) => g.level === level && g.start <= i && i <= g.end && axis.isCollapsed(g.summary) && axis.isHidden(i),
    )
    changes.set(i, { level: level - 1, ...(collapsedBy ? { hidden: false } : {}) })
  }
  if (!changes.size) return null
  // Summaries of groups that no longer exist: one level of nesting goes from a range ungrouped whole, so a summary
  // shared by nested groups of the same lines keeps its mark.
  for (const g of groups) {
    const whole = from <= g.start && g.end <= to
    const sharing = groups.filter((h) => h.summary === g.summary).length
    if (whole && sharing === 1 && g.summary >= 0 && g.summary < axis.count && axis.isCollapsed(g.summary))
      changes.set(g.summary, { ...changes.get(g.summary), collapsed: undefined })
  }
  const entries = [...changes].map(([i, c]) => {
    const e = entry(axis, i, c)
    if ("collapsed" in c && !c.collapsed) delete e.collapsed
    return e
  })
  return axisOp(index, name, entries)
}

/** Ungroup › Clear Outline: no groups left on the axis (hidden detail stays hidden, as in Excel). */
export function clearOutlineOps(sheet: Sheet, index: number, name: AxisName): Op[] {
  const axis = axisOf(sheet, name)
  const entries = axis
    .overrides(0, axis.count - 1)
    .filter((e) => e.level || e.collapsed)
    .map((e) => {
      const next = entry(axis, e.index, { level: 0 })
      delete next.collapsed
      return next
    })
  return axisOp(index, name, entries)
}

/** Hide (−) or show (+) one group's detail. Inner groups that are collapsed stay so when their parent opens. */
export function toggleGroupOps(sheet: Sheet, index: number, name: AxisName, group: OutlineGroup, show: boolean): Op[] {
  const axis = axisOf(sheet, name)
  const keepHidden = new Set<number>()
  if (show)
    for (const inner of outlineGroups(axis, summaryAfter(sheet, name)))
      if (inner.level > group.level && inner.start >= group.start && inner.end <= group.end && isCollapsed(axis, inner))
        for (let i = inner.start; i <= inner.end; i++) keepHidden.add(i)
  const entries: AxisEntry[] = []
  for (let i = group.start; i <= group.end; i++) {
    const hidden = show ? keepHidden.has(i) : true
    if (axis.isHidden(i) !== hidden) entries.push(entry(axis, i, { hidden }))
  }
  if (group.summary >= 0 && group.summary < axis.count) {
    const e = entry(axis, group.summary, show ? {} : { collapsed: true })
    if (show) delete e.collapsed
    entries.push(e)
  }
  return axisOp(index, name, entries)
}

/** Whether a group is shown collapsed: its summary is marked, or (no summary line) its detail is all hidden. */
export function isCollapsed(axis: Axis, group: OutlineGroup): boolean {
  if (group.summary >= 0 && group.summary < axis.count && axis.isCollapsed(group.summary)) return true
  for (let i = group.start; i <= group.end; i++) if (!axis.isHidden(i)) return false
  return true
}

/** The outline bar's level button `n` (1 = only the top level): detail deeper than `n` hidden, the rest shown. */
export function showLevelOps(sheet: Sheet, index: number, name: AxisName, n: number): Op[] {
  const axis = axisOf(sheet, name)
  const groups = outlineGroups(axis, summaryAfter(sheet, name))
  const entries = new Map<number, AxisEntry>()
  for (const g of groups) {
    for (let i = g.start; i <= g.end; i++) {
      const hidden = axis.level(i) >= n
      if (axis.isHidden(i) !== hidden) entries.set(i, entry(axis, i, { hidden }))
    }
  }
  for (const g of groups) {
    if (g.summary < 0 || g.summary >= axis.count) continue
    const collapsed = g.level >= n
    if (axis.isCollapsed(g.summary) === collapsed) continue
    const e = entries.get(g.summary) ?? entry(axis, g.summary, {})
    if (collapsed) e.collapsed = true
    else delete e.collapsed
    entries.set(g.summary, e)
  }
  return axisOp(index, name, [...entries.values()])
}

/**
 * Show Detail / Hide Detail for the selection: the innermost group the selected lines are in, or that the selected
 * summary line closes. Null when the selection isn't in or beside a group.
 */
export function detailOps(sheet: Sheet, index: number, name: AxisName, from: number, to: number, show: boolean) {
  const axis = axisOf(sheet, name)
  const groups = outlineGroups(axis, summaryAfter(sheet, name))
  // A selected summary line acts on its own group first.
  const own = groups.filter((g) => g.summary >= from && g.summary <= to)
  const inside = groups.filter((g) => g.start <= from && to <= g.end)
  const pick = show
    ? (own.length ? own : inside).filter((g) => isCollapsed(axis, g)).sort((a, b) => a.level - b.level)[0]
    : (inside.length ? inside : own).sort((a, b) => b.level - a.level)[0]
  return pick ? toggleGroupOps(sheet, index, name, pick, show) : null
}

// ── Subtotal ──────────────────────────────────────────────────────────────────────────────────────────────────

/** Subtotal's "Use function" list: label and SUBTOTAL's function number. */
export const SUBTOTAL_FUNCTIONS: { label: string; code: number; word: string }[] = [
  { label: "Sum", code: 9, word: "Total" },
  { label: "Count", code: 3, word: "Count" },
  { label: "Average", code: 1, word: "Average" },
  { label: "Max", code: 4, word: "Max" },
  { label: "Min", code: 5, word: "Min" },
  { label: "Product", code: 6, word: "Product" },
  { label: "Count Numbers", code: 2, word: "Count" },
  { label: "StdDev", code: 7, word: "StdDev" },
  { label: "StdDevp", code: 8, word: "StdDevp" },
  { label: "Var", code: 10, word: "Var" },
  { label: "Varp", code: 11, word: "Varp" },
]

export type SubtotalOptions = {
  /** "At each change in": the column (sheet index) whose value starts a new group. */
  by: number
  /** Index into SUBTOTAL_FUNCTIONS. */
  fn: number
  /** "Add subtotal to": columns (sheet indexes). */
  cols: number[]
  /** Replace current subtotals. */
  replace: boolean
  /** Page break between groups. */
  pageBreaks: boolean
  /** Summary below data (else above). */
  below: boolean
}

const isSubtotalRow = (sheet: Sheet, row: number, c0: number, c1: number) => {
  for (let c = c0; c <= c1; c++) if (/^=\s*SUBTOTAL\(/i.test(sheet.cells.input(row, c) ?? "")) return true
  return false
}

/** Remove All: the subtotal rows in `range` deleted and its rows' outline cleared. */
export function removeSubtotalsOps(sheet: Sheet, index: number, range: Range): Op[] {
  const ops: Op[] = []
  const entries: AxisEntry[] = []
  for (let r = range.r0; r <= range.r1; r++)
    if (sheet.rows.level(r) || sheet.rows.isCollapsed(r) || (sheet.rows.isHidden(r) && sheet.rows.level(r))) {
      const e = entry(sheet.rows, r, { level: 0, hidden: false })
      delete e.collapsed
      entries.push(e)
    }
  ops.push(...axisOp(index, "rows", entries))
  for (let r = range.r1; r >= range.r0; r--)
    if (isSubtotalRow(sheet, r, range.c0, range.c1)) ops.push({ kind: "deleteRows", sheet: index, at: r, count: 1 })
  return ops
}

/**
 * Subtotal: a row after (or before) each run of equal values in the `by` column, with SUBTOTAL over the run for each
 * chosen column, a grand total, and the rows grouped so the outline's level 1 shows the grand total, 2 the
 * subtotals, 3 everything. `range` includes the header row. Without Replace, existing subtotals stay and the new
 * ones nest inside them. Returns the operations and the range afterwards, or a problem to show.
 */
export function subtotalOps(
  sheet: Sheet,
  styles: StyleTable,
  index: number,
  range: Range,
  o: SubtotalOptions,
): { ops: Op[]; range: Range } | { problem: string } {
  if (range.r1 <= range.r0) return { problem: "Cells can't tell which row has the column labels." }
  if (!o.cols.length) return { problem: "Choose at least one column to add a subtotal to." }
  const fn = SUBTOTAL_FUNCTIONS[o.fn] ?? SUBTOTAL_FUNCTIONS[0]
  const ops: Op[] = []
  const empty = (r: number) => {
    for (let c = range.c0; c <= range.c1; c++) if (sheet.cells.get(r, c)?.value !== undefined) return false
    return true
  }
  // Empty rows at the end aren't part of the list (a selected blank cell under it brings one in).
  while (range.r1 > range.r0 && empty(range.r1)) range = { ...range, r1: range.r1 - 1 }
  // Where each row will be once Replace has deleted the old subtotal rows.
  let removed = 0
  const runs: { start: number; end: number; value: unknown; level: number }[] = []
  for (let r = range.r0 + 1; r <= range.r1; r++) {
    if (isSubtotalRow(sheet, r, range.c0, range.c1)) {
      if (o.replace) removed++
      continue
    }
    if (empty(r)) continue
    const at = r - removed
    const value = sheet.cells.value(r, o.by)
    const level = o.replace ? 0 : sheet.rows.level(r)
    const last = runs[runs.length - 1]
    if (last && last.end === at - 1 && last.level === level && sameValue(last.value, value)) last.end = at
    else runs.push({ start: at, end: at, value, level })
  }
  if (!runs.length) return { problem: "There's no data under the column labels to subtotal." }
  if (o.replace) ops.push(...removeSubtotalsOps(sheet, index, range))
  const r1 = range.r1 - removed
  const nested = runs.some((run) => run.level > 0)
  // Insert the rows from the bottom up, so each insertion leaves the ones above where they were.
  for (let k = runs.length - 1; k >= 0; k--)
    ops.push({ kind: "insertRows", sheet: index, at: o.below ? runs[k].end + 1 : runs[k].start, count: 1 })
  // The grand total: after everything, or (summaries above) right under the labels.
  const top = !nested && !o.below ? 1 : 0
  if (!nested)
    ops.push({ kind: "insertRows", sheet: index, at: o.below ? r1 + 1 + runs.length : range.r0 + 1, count: 1 })
  const cells: CellWrite[] = []
  const levels: [number, number][] = []
  const breaks: number[] = []
  const write = (row: number, col: number, text: string) => {
    const data = contentForInput(styles, undefined, text)
    if (data) cells.push({ row, col, data })
  }
  const label = (value: unknown) => `${value === null || value === undefined ? "" : String(value)} ${fn.word}`.trim()
  const ref = (col: number, a: number, b: number) => `${columnName(col)}${a + 1}:${columnName(col)}${b + 1}`
  let first = Infinity
  let last = -Infinity
  runs.forEach((run, k) => {
    // The run's rows after the k rows inserted above it (one more for its own total when totals go above).
    const start = run.start + k + top + (o.below ? 0 : 1)
    const end = run.end + k + top + (o.below ? 0 : 1)
    const total = o.below ? end + 1 : start - 1
    first = Math.min(first, start, total)
    last = Math.max(last, end, total)
    write(total, o.by, label(run.value))
    for (const col of o.cols) write(total, col, `=SUBTOTAL(${fn.code},${ref(col, start, end)})`)
    const base = nested ? run.level : 1
    levels.push([total, base])
    for (let r = start; r <= end; r++) levels.push([r, base + 1])
    if (k > 0) breaks.push(Math.min(start, total))
  })
  if (!nested) {
    const grand = o.below ? last + 1 : range.r0 + 1
    write(grand, o.by, `Grand ${fn.word}`)
    for (const col of o.cols) write(grand, col, `=SUBTOTAL(${fn.code},${ref(col, first, last)})`)
  }
  ops.push({ kind: "setCells", sheet: index, cells })
  ops.push({ kind: "outlineLevels", sheet: index, axis: "rows", levels })
  if (o.pageBreaks) {
    const rowBreaks = [...new Set([...sheet.pageSetup.rowBreaks, ...breaks])].sort((a, b) => a - b)
    ops.push({ kind: "pageSetup", sheet: index, setup: { ...sheet.pageSetup, rowBreaks } })
  }
  return { ops, range: { ...range, r1: r1 + runs.length + (nested ? 0 : 1) } }
}

function sameValue(a: unknown, b: unknown) {
  if (typeof a === "string" && typeof b === "string") return a.toLowerCase() === b.toLowerCase()
  const empty = (v: unknown) => v === null || v === undefined || v === ""
  return a === b || (empty(a) && empty(b))
}

/**
 * Group › Auto Outline: groups made from the sheet's formulas, as Excel makes them: a formula that sums the lines
 * just before it (rows above with summaries below, columns to the left with summaries to the right) makes those lines
 * a group; groups inside groups go a level deeper. Replaces the outline the sheet had. Null when no formula
 * summarises anything.
 */
export function autoOutlineOps(sheet: Sheet, index: number): Op[] | null {
  const used = sheet.cells.usedBounds()
  if (!used) return null
  const below = sheet.outline.summaryBelow
  const right = sheet.outline.summaryRight
  const spans: Record<AxisName, Map<string, [number, number]>> = { rows: new Map(), cols: new Map() }
  sheet.cells.forEach(used, (row, col, data) => {
    if (!data.input?.startsWith("=")) return
    for (const r of formulaRefs(data.input, sheet.name).local) {
      const inCol = r.c0 <= col && col <= r.c1
      const inRow = r.r0 <= row && row <= r.r1
      if (inCol && !inRow && (below ? r.r1 === row - 1 : r.r0 === row + 1))
        spans.rows.set(`${r.r0},${r.r1}`, [r.r0, r.r1])
      if (inRow && !inCol && (right ? r.c1 === col - 1 : r.c0 === col + 1))
        spans.cols.set(`${r.c0},${r.c1}`, [r.c0, r.c1])
    }
  })
  if (!spans.rows.size && !spans.cols.size) return null
  const ops: Op[] = []
  for (const name of ["rows", "cols"] as const) {
    const axis = axisOf(sheet, name)
    const depth = new Map<number, number>()
    for (const [a, b] of spans[name].values())
      for (let i = a; i <= b; i++) depth.set(i, Math.min(MAX_OUTLINE, (depth.get(i) ?? 0) + 1))
    const entries = new Map<number, AxisEntry>()
    for (const e of axis.overrides(0, axis.count - 1))
      if (e.level || e.collapsed) {
        const next = entry(axis, e.index, { level: 0 })
        delete next.collapsed
        entries.set(e.index, next)
      }
    for (const [i, level] of depth) entries.set(i, { ...(entries.get(i) ?? entry(axis, i, {})), level })
    ops.push(...axisOp(index, name, [...entries.values()]))
  }
  return ops
}
