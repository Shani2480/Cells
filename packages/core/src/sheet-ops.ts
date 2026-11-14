import type { Range } from "./address"
import type { AnnotationWrite } from "./annotations"
import { copyEntries, type CFEntry } from "./conditional"
import { copyValidations, type ValidationEntry } from "./validation"
import { copyTables, type Table } from "./table/table"
import { rewriteNames, type SheetList } from "./names"
import { copyAutoFilter, type AutoFilter } from "./filter"
import type { SheetProtection } from "./protection"
import { renameSheetRefs } from "./formula/adjust"
import type { Applied, AxisEntry, CellWrite, Op } from "./ops"
import { Sheet } from "./sheet"

/**
 * Sheet-level operations: add, delete (with everything on it, for undo), rename, move, tab colour and hiding.
 * Formulas name sheets, so a rename rewrites the references to the old name everywhere and a delete turns them into
 * #REF! (as Excel). Sheet indices shift, so these count as structural changes: the engine re-reads every formula.
 */

/** Everything on a sheet, as data (the undo of a delete puts it back exactly). */
export type SheetSnapshot = {
  name: string
  cells: CellWrite[]
  rows: AxisEntry[]
  cols: AxisEntry[]
  rowStyles: number[]
  colStyles: number[]
  merges: Range[]
  conditional?: CFEntry[]
  validations?: ValidationEntry[]
  tables?: Table[]
  annotations?: AnnotationWrite[]
  protection?: SheetProtection | null
  autoFilter?: AutoFilter | null
  tabColor?: string
  hidden: boolean
}

export type SheetOp =
  | { kind: "addSheet"; sheet: number; name: string; snapshot?: SheetSnapshot }
  | { kind: "removeSheet"; sheet: number }
  | { kind: "renameSheet"; sheet: number; name: string }
  | { kind: "moveSheet"; sheet: number; to: number }
  | { kind: "sheetProps"; sheet: number; tabColor?: string | null; hidden?: boolean }

export function snapshotSheet(sheet: Sheet): SheetSnapshot {
  const used = sheet.cells.usedBounds()
  return {
    name: sheet.name,
    cells: used ? sheet.cells.entries(used).map((e) => ({ row: e.row, col: e.col, data: { ...e.data } })) : [],
    rows: sheet.rows.overrides(0, sheet.rows.count - 1),
    cols: sheet.cols.overrides(0, sheet.cols.count - 1),
    rowStyles: sheet.rowStyles.entries().flat(),
    colStyles: sheet.colStyles.entries().flat(),
    merges: sheet.merges.all.map((m) => ({ ...m })),
    conditional: copyEntries(sheet.conditional),
    validations: copyValidations(sheet.validations),
    tables: copyTables(sheet.tables),
    annotations: sheet.annotations.all().map((a) => ({ row: a.row, col: a.col, value: a.value })),
    protection: sheet.protection ? { ...sheet.protection } : null,
    autoFilter: copyAutoFilter(sheet.autoFilter),
    tabColor: sheet.tabColor,
    hidden: sheet.hidden,
  }
}

function sheetFrom(s: SheetSnapshot): Sheet {
  const sheet = new Sheet(s.name)
  for (const w of s.cells) sheet.cells.set(w.row, w.col, w.data)
  for (const e of s.rows) {
    if (e.size !== null) sheet.rows.setSize(e.index, e.size, e.auto)
    if (e.hidden) sheet.rows.setHidden(e.index, true)
  }
  for (const e of s.cols) {
    if (e.size !== null) sheet.cols.setSize(e.index, e.size, e.auto)
    if (e.hidden) sheet.cols.setHidden(e.index, true)
  }
  for (let i = 0; i < s.rowStyles.length; i += 2) sheet.rowStyles.set(s.rowStyles[i], s.rowStyles[i + 1])
  for (let i = 0; i < s.colStyles.length; i += 2) sheet.colStyles.set(s.colStyles[i], s.colStyles[i + 1])
  sheet.merges.replace(s.merges.map((m) => ({ ...m })))
  sheet.conditional = copyEntries(s.conditional ?? [])
  sheet.validations = copyValidations(s.validations ?? [])
  sheet.tables = copyTables(s.tables ?? [])
  for (const a of s.annotations ?? []) sheet.annotations.set(a.row, a.col, a.value)
  sheet.protection = s.protection ? { ...s.protection } : null
  sheet.autoFilter = copyAutoFilter(s.autoFilter ?? null)
  sheet.tabColor = s.tabColor
  sheet.hidden = s.hidden
  return sheet
}

/** Rewrites every formula that names `from`; returns the setCells ops that put the old texts back. */
function renameRefs(sheets: Sheet[], from: string, to: string | null): Op[] {
  const undo: Op[] = []
  sheets.forEach((sheet, index) => {
    const previous: CellWrite[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      if (!input.includes("!")) return
      const text = renameSheetRefs(input, from, to)
      if (text === null) return
      const data = sheet.cells.get(row, col)!
      previous.push({ row, col, data })
      sheet.cells.set(row, col, { ...data, input: text })
    })
    if (previous.length) undo.push({ kind: "setCells", sheet: index, cells: previous })
  })
  const names = rewriteNames(sheets as SheetList, (ref) => renameSheetRefs(ref, from, to))
  if (names) undo.push({ kind: "names", names })
  return undo
}

const structural = (sheet: number) => ({ sheet, structural: true })

export function applySheetOp(sheets: Sheet[], op: SheetOp): Applied {
  switch (op.kind) {
    case "addSheet": {
      const sheet = op.snapshot ? sheetFrom(op.snapshot) : new Sheet(op.name)
      sheets.splice(op.sheet, 0, sheet)
      // References that were #REF! stay so; a re-added sheet (undo of a delete) gets its texts back from the delete's
      // inverse, which also holds the setCells that restore them.
      return { inverse: [{ kind: "removeSheet", sheet: op.sheet }], change: structural(op.sheet) }
    }
    case "removeSheet": {
      const sheet = sheets[op.sheet]
      const snapshot = snapshotSheet(sheet)
      sheets.splice(op.sheet, 1)
      // Indices of later sheets drop by one: the undo texts are recorded after the splice, so they use the new ones.
      const texts = renameRefs(sheets, sheet.name, null)
      return {
        // Undo: put the sheet back first (indices as before), then the texts (shifted back by the re-insert).
        inverse: [
          { kind: "addSheet", sheet: op.sheet, name: sheet.name, snapshot },
          ...texts.map((t) => (t.kind === "setCells" && t.sheet >= op.sheet ? { ...t, sheet: t.sheet + 1 } : t)),
        ],
        change: structural(op.sheet),
      }
    }
    case "renameSheet": {
      const sheet = sheets[op.sheet]
      const old = sheet.name
      const texts = renameRefs(sheets, old, op.name)
      sheet.name = op.name
      return {
        inverse: [{ kind: "renameSheet", sheet: op.sheet, name: old }, ...texts],
        change: structural(op.sheet),
      }
    }
    case "moveSheet": {
      const [sheet] = sheets.splice(op.sheet, 1)
      sheets.splice(op.to, 0, sheet)
      return { inverse: [{ kind: "moveSheet", sheet: op.to, to: op.sheet }], change: structural(op.to) }
    }
    case "sheetProps": {
      const sheet = sheets[op.sheet]
      const before: SheetOp = { kind: "sheetProps", sheet: op.sheet }
      if (op.tabColor !== undefined) {
        before.tabColor = sheet.tabColor ?? null
        sheet.tabColor = op.tabColor ?? undefined
      }
      if (op.hidden !== undefined) {
        before.hidden = sheet.hidden
        sheet.hidden = op.hidden
      }
      return { inverse: [before], change: { sheet: op.sheet, structural: false, format: true } }
    }
  }
}

/** A name Excel accepts for a sheet: 1–31 characters, none of : \ / ? * [ ], not starting or ending with '. */
export function sheetNameProblem(name: string, others: string[]): string | null {
  if (!name.trim()) return "You typed an invalid name for a sheet or chart."
  if (name.length > 31 || /[:\\/?*[\]]/.test(name) || name.startsWith("'") || name.endsWith("'"))
    return "You typed an invalid name for a sheet or chart. Make sure that: The name that you type does not exceed 31 characters; The name does not contain any of the following characters: : \\ / ? * [ or ]; You did not leave the name blank."
  if (name.toLowerCase() === "history") return "This name is reserved by Excel. Choose a different name."
  if (others.some((o) => o.toLowerCase() === name.toLowerCase()))
    return "That name is already taken. Try a different one."
  return null
}

/** The next free "SheetN" (Excel counts up from the number of sheets). */
export function nextSheetName(names: string[]): string {
  const taken = new Set(names.map((n) => n.toLowerCase()))
  for (let n = names.length + 1; ; n++) if (!taken.has(`sheet${n}`)) return `Sheet${n}`
}
