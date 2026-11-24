import {
  copyCustomStyle,
  customStyle,
  customStyleNameProblem,
  DEFAULT_PIVOT_STYLE,
  DEFAULT_TABLE_STYLE,
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

/** Table styles or PivotTable styles. */
export type StyleKind = "table" | "pivot"

/** What the New / Modify Table (PivotTable) Style dialog edits: the style and whether it is the workbook's default. */
export type CustomStyleDraft = { style: CustomTableStyle; isDefault: boolean }

const same = (a: string | null | undefined, b: string) => !!a && a.toLowerCase() === b.toLowerCase()

/**
 * The workbook's own table or PivotTable styles (Design › Table Styles › New Table Style, PivotTable Design ›
 * PivotTable Styles › New PivotTable Style, and the galleries' Modify, Duplicate, Delete and Set As Default). Each
 * change is one undo step. Names are shared by both kinds, as in Excel.
 */
export class CustomStyleCommands {
  constructor(
    private readonly c: SheetController,
    readonly kind: StyleKind,
  ) {}

  /** The built-in style new tables (PivotTables) take without a default of the workbook's. */
  get builtInDefault(): string {
    return this.kind === "pivot" ? DEFAULT_PIVOT_STYLE : DEFAULT_TABLE_STYLE
  }

  private get defaultKey(): "defaultPivot" | "defaultTable" {
    return this.kind === "pivot" ? "defaultPivot" : "defaultTable"
  }

  /** The workbook's styles shown with this kind's, in the gallery's Custom group. */
  list(): CustomTableStyle[] {
    return this.c.wb.tableStyles.styles.filter((s) => s[this.kind])
  }

  get(name: string): CustomTableStyle | undefined {
    return customStyle(this.c.wb.tableStyles.styles, name, this.kind)
  }

  /** The default for new tables (PivotTables). */
  defaultStyle(): string {
    return this.c.wb.tableStyles[this.defaultKey] ?? this.builtInDefault
  }

  /** Why the style can't take its name (null: it can); `editing` is the name of the style being modified. */
  nameProblem(name: string, editing?: string): string | null {
    return customStyleNameProblem(name, this.c.wb.tableStyles.styles, isBuiltInStyle, editing)
  }

  /**
   * New Style's OK (Modify's, with `editing`): the style saved, renamed tables following a new name, and "Set as
   * default … style for this document". A problem with the name comes back as Excel's message.
   */
  save(draft: CustomStyleDraft, editing?: string): string | null {
    const name = draft.style.name.trim()
    const problem = this.nameProblem(name, editing)
    if (problem) return problem
    const set = this.c.wb.tableStyles
    const style = { ...copyCustomStyle(draft.style), name }
    const old = editing ? set.styles.findIndex((s) => same(s.name, editing)) : -1
    const styles = [...set.styles]
    if (old >= 0) styles[old] = style
    else styles.push(style)
    const current = set[this.defaultKey]
    const wasDefault = editing !== undefined && same(current, editing)
    const next: TableStyleSet = {
      ...set,
      styles,
      [this.defaultKey]: draft.isDefault ? name : wasDefault ? null : current,
    }
    const renamed = editing && editing !== name ? this.restyled(editing, name) : []
    const what = this.kind === "pivot" ? "PivotTable Style" : "Table Style"
    this.commit([{ kind: "tableStyles", set: next }, ...renamed], editing ? `Modify ${what}` : `New ${what}`)
    return null
  }

  /** Delete: tables (PivotTables) with the style go back to Medium 2 (Light 16), and the default with it, as Excel. */
  remove(name: string): boolean {
    const set = this.c.wb.tableStyles
    const styles = set.styles.filter((s) => !same(s.name, name))
    if (styles.length === set.styles.length) return false
    const next: TableStyleSet = {
      ...set,
      styles,
      defaultTable: same(set.defaultTable, name) ? null : set.defaultTable,
      defaultPivot: same(set.defaultPivot, name) ? null : set.defaultPivot,
    }
    const back = [
      ...new CustomStyleCommands(this.c, "table").restyled(name, DEFAULT_TABLE_STYLE),
      ...new CustomStyleCommands(this.c, "pivot").restyled(name, DEFAULT_PIVOT_STYLE),
    ]
    this.commit([{ kind: "tableStyles", set: next }, ...back], "Delete Style")
    return true
  }

  /** Set As Default (built-in or the workbook's own): new tables (PivotTables) in this workbook take it. */
  setDefault(name: string): boolean {
    const set = this.c.wb.tableStyles
    const value = name === this.builtInDefault ? null : name
    if (set[this.defaultKey] === value) return false
    this.commit([{ kind: "tableStyles", set: { ...set, [this.defaultKey]: value } }], "Set As Default")
    return true
  }

  /** Every table (PivotTable) with style `from` given style `to` (ops per sheet). */
  private restyled(from: string, to: string): Op[] {
    const ops: Op[] = []
    this.c.wb.sheets.forEach((sheet, index) => {
      if (this.kind === "pivot") {
        if (!sheet.pivots.some((t) => same(t.style.name, from))) return
        const tables = sheet.pivots.map((t) =>
          same(t.style.name, from) ? { ...t, style: { ...t.style, name: to } } : t,
        )
        ops.push({ kind: "pivots", sheet: index, tables })
      } else {
        if (!sheet.tables.some((t) => same(t.style, from))) return
        const tables = sheet.tables.map((t) => (same(t.style, from) ? { ...t, style: to } : t))
        ops.push({ kind: "tables", sheet: index, tables })
      }
    })
    return ops
  }

  private commit(ops: Op[], label: string) {
    this.c.transact(ops, { label, before: this.c.selection, after: this.c.selection })
  }
}
