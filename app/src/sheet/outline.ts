import {
  autoOutlineOps,
  clearOutlineOps,
  currentRegion,
  detailOps,
  groupOps,
  isWholeColumns,
  isWholeRows,
  outlineGroups,
  removeSubtotalsOps,
  showLevelOps,
  subtotalOps,
  toggleGroupOps,
  ungroupOps,
  type AxisName,
  type Op,
  type OutlineGroup,
  type OutlineSettings,
  type Range,
  type SubtotalOptions,
} from "@cells/core"
import type { SheetController } from "./controller"

/**
 * Data › Outline on the active sheet: Group and Ungroup (Alt+Shift+Right / Left), Clear Outline, Show and Hide
 * Detail, the outline bar's + / − and level buttons, Settings, Ctrl+8, and Subtotal.
 */
export class OutlineCommands {
  constructor(private readonly c: SheetController) {}

  private apply(ops: Op[] | null, label: string): boolean {
    if (!ops) return false
    if (this.c.edit) this.c.commitEdit()
    if (!ops.length) return true
    const sel = this.c.selection
    this.c.transact(ops, { label, before: sel, after: sel })
    return true
  }

  /** Which lines the selection groups: whole rows or columns decide; otherwise null (the Group dialog asks). */
  axisOfSelection(): AxisName | null {
    const r = this.c.selection.ranges[0]
    if (isWholeRows(r) && !isWholeColumns(r)) return "rows"
    if (isWholeColumns(r) && !isWholeRows(r)) return "cols"
    return null
  }

  private span(axis: AxisName): [number, number] {
    const r = this.c.selection.ranges[0]
    return axis === "rows" ? [r.r0, r.r1] : [r.c0, r.c1]
  }

  /** Group: false when the lines are already 7 levels deep (Excel's limit). */
  group(axis: AxisName): boolean {
    const [from, to] = this.span(axis)
    const ok = this.apply(groupOps(this.c.sheet, this.c.sheetIndex, axis, from, to), "Group")
    if (!ok) this.c.shell.alert("Cells can't group these any deeper: an outline has at most 8 levels.")
    return ok
  }

  ungroup(axis: AxisName): boolean {
    const [from, to] = this.span(axis)
    const ok = this.apply(ungroupOps(this.c.sheet, this.c.sheetIndex, axis, from, to), "Ungroup")
    if (!ok) this.c.shell.alert("Cannot ungroup. There's no group here.")
    return ok
  }

  /** Group › Auto Outline. */
  autoOutline() {
    if (!this.apply(autoOutlineOps(this.c.sheet, this.c.sheetIndex), "Auto Outline"))
      this.c.shell.alert(
        "Cannot create an outline. Cells found no formulas that add up the rows or columns beside them.",
      )
  }

  /** Ungroup › Clear Outline: rows and columns. */
  clearOutline() {
    const s = this.c.sheet
    if (!s.rows.maxLevel && !s.cols.maxLevel) {
      this.c.shell.alert("Cannot clear the outline. There's no outline on this sheet.")
      return
    }
    this.apply(
      [...clearOutlineOps(s, this.c.sheetIndex, "rows"), ...clearOutlineOps(s, this.c.sheetIndex, "cols")],
      "Clear Outline",
    )
  }

  /** Show Detail / Hide Detail on the selection: rows first, columns when whole columns are selected. */
  detail(show: boolean) {
    const r = this.c.selection.ranges[0]
    const axes: AxisName[] = isWholeColumns(r) && !isWholeRows(r) ? ["cols"] : ["rows", "cols"]
    for (const axis of axes) {
      const [from, to] = axis === "rows" ? [r.r0, r.r1] : [r.c0, r.c1]
      const ops = detailOps(this.c.sheet, this.c.sheetIndex, axis, from, to, show)
      if (ops) return void this.apply(ops, show ? "Show Detail" : "Hide Detail")
    }
    if (!show) this.c.shell.alert("Cannot hide the detail. Select a cell in a group first.")
  }

  /** The outline bar's + and −. */
  toggle(axis: AxisName, group: OutlineGroup, show: boolean) {
    this.apply(toggleGroupOps(this.c.sheet, this.c.sheetIndex, axis, group, show), show ? "Show Detail" : "Hide Detail")
  }

  /** The outline bar's level buttons. */
  showLevel(axis: AxisName, n: number) {
    this.apply(showLevelOps(this.c.sheet, this.c.sheetIndex, axis, n), "Show Level")
  }

  groups(axis: AxisName): OutlineGroup[] {
    const s = this.c.sheet
    return outlineGroups(
      axis === "rows" ? s.rows : s.cols,
      axis === "rows" ? s.outline.summaryBelow : s.outline.summaryRight,
    )
  }

  /** Data › Outline › Settings. */
  setSettings(settings: OutlineSettings) {
    this.apply([{ kind: "outlineSettings", sheet: this.c.sheetIndex, settings }], "Outline Settings")
  }

  /** Ctrl+8: show or hide the outline bar. */
  toggleSymbols() {
    const s = this.c.sheet.outline
    if (!this.c.sheet.rows.maxLevel && !this.c.sheet.cols.maxLevel && s.showSymbols) {
      this.c.shell.alert("Cannot show or hide the outline symbols. There's no outline on this sheet.")
      return
    }
    this.setSettings({ ...s, showSymbols: !s.showSymbols })
  }

  /** The list Subtotal works on: the selection, or the data around the active cell when one cell is selected. */
  subtotalRange(): Range {
    const r = this.c.selection.ranges[0]
    if (r.r0 !== r.r1 || r.c0 !== r.c1) return r
    return currentRegion(this.c.sheet, this.c.selection.active)
  }

  /** Subtotal: null when done, else the problem to show. */
  subtotal(range: Range, o: SubtotalOptions): string | null {
    const result = subtotalOps(this.c.sheet, this.c.wb.styles, this.c.sheetIndex, range, o)
    if ("problem" in result) return result.problem
    if (this.c.edit) this.c.commitEdit()
    const before = this.c.selection
    const after = { ...before, ranges: [result.range] }
    this.c.transact(result.ops, { label: "Subtotal", before, after })
    return null
  }

  removeSubtotals(range: Range) {
    this.apply(removeSubtotalsOps(this.c.sheet, this.c.sheetIndex, range), "Remove Subtotals")
  }
}
