import {
  copyCustomStyle,
  customStyle,
  customStyleNameProblem,
  DEFAULT_PIVOT_STYLE,
  PIVOT_STYLE_GROUPS,
  TABLE_STYLE_GROUPS,
  type CustomTableStyle,
  type Op,
  type TableStyleSet,
} from "@cells/core"
import type { SheetController } from "./controller"

const BUILT_IN = new Set(
  [...PIVOT_STYLE_GROUPS, ...TABLE_STYLE_GROUPS].flatMap((g) => g.names).map((n) => n.toLowerCase()),
)
/** A name Excel's built-in table and PivotTable styles use. */
export const isBuiltInStyle = (name: string) => BUILT_IN.has(name.toLowerCase())

/** What the New / Modify PivotTable Style dialog edits: the style and whether it is the workbook's default. */
export type PivotStyleDraft = { style: CustomTableStyle; isDefault: boolean }

/**
 * The workbook's own PivotTable styles (PivotTable Design › PivotTable Styles › New PivotTable Style, and the
 * gallery's Modify, Duplicate, Delete and Set As Default). Each change is one undo step.
 */
export class PivotStyleCommands {
  constructor(private readonly c: SheetController) {}

  /** The workbook's styles shown with PivotTable styles, in the gallery's Custom group. */
  list(): CustomTableStyle[] {
    return this.c.wb.tableStyles.styles.filter((s) => s.pivot)
  }

  get(name: string): CustomTableStyle | undefined {
    return customStyle(this.c.wb.tableStyles.styles, name, "pivot")
  }

  /** The default for new PivotTables. */
  defaultStyle(): string {
    return this.c.wb.tableStyles.defaultPivot ?? DEFAULT_PIVOT_STYLE
  }

  /** Why the style can't take its name (null: it can); `editing` is the name of the style being modified. */
  nameProblem(name: string, editing?: string): string | null {
    return customStyleNameProblem(name, this.c.wb.tableStyles.styles, isBuiltInStyle, editing)
  }

  /**
   * New PivotTable Style's OK (Modify's, with `editing`): the style saved, renamed PivotTables following a new name,
   * and Set as default PivotTable style for this document. A problem with the name comes back as Excel's message.
   */
  save(draft: PivotStyleDraft, editing?: string): string | null {
    const name = draft.style.name.trim()
    const problem = this.nameProblem(name, editing)
    if (problem) return problem
    const set = this.c.wb.tableStyles
    const style = { ...copyCustomStyle(draft.style), name }
    const old = editing ? set.styles.findIndex((s) => s.name.toLowerCase() === editing.toLowerCase()) : -1
    const styles = [...set.styles]
    if (old >= 0) styles[old] = style
    else styles.push(style)
    const wasDefault = editing !== undefined && set.defaultPivot?.toLowerCase() === editing.toLowerCase()
    const next: TableStyleSet = {
      ...set,
      styles,
      defaultPivot: draft.isDefault ? name : wasDefault ? null : set.defaultPivot,
    }
    const renamed = editing && editing !== name ? this.restyled(editing, name) : []
    this.commit(
      [{ kind: "tableStyles", set: next }, ...renamed],
      editing ? "Modify PivotTable Style" : "New PivotTable Style",
    )
    return null
  }

  /** Delete: PivotTables with the style go back to Light 16, and the default with it, as in Excel. */
  remove(name: string): boolean {
    const set = this.c.wb.tableStyles
    const styles = set.styles.filter((s) => s.name.toLowerCase() !== name.toLowerCase())
    if (styles.length === set.styles.length) return false
    const isDefault = set.defaultPivot?.toLowerCase() === name.toLowerCase()
    const next: TableStyleSet = { ...set, styles, defaultPivot: isDefault ? null : set.defaultPivot }
    this.commit([{ kind: "tableStyles", set: next }, ...this.restyled(name, DEFAULT_PIVOT_STYLE)], "Delete Style")
    return true
  }

  /** Set As Default (built-in or the workbook's own): new PivotTables in this workbook take it. */
  setDefault(name: string): boolean {
    const set = this.c.wb.tableStyles
    const defaultPivot = name === DEFAULT_PIVOT_STYLE ? null : name
    if (set.defaultPivot === defaultPivot) return false
    this.commit([{ kind: "tableStyles", set: { ...set, defaultPivot } }], "Set As Default")
    return true
  }

  /** Every PivotTable with style `from` given style `to` (ops per sheet). */
  private restyled(from: string, to: string): Op[] {
    const ops: Op[] = []
    this.c.wb.sheets.forEach((sheet, index) => {
      if (!sheet.pivots.some((t) => t.style.name?.toLowerCase() === from.toLowerCase())) return
      const tables = sheet.pivots.map((t) =>
        t.style.name?.toLowerCase() === from.toLowerCase() ? { ...t, style: { ...t.style, name: to } } : t,
      )
      ops.push({ kind: "pivots", sheet: index, tables })
    })
    return ops
  }

  private commit(ops: Op[], label: string) {
    this.c.transact(ops, { label, before: this.c.selection, after: this.c.selection })
  }
}
