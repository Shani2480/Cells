import {
  copiedSheetObjects,
  nextTableId,
  renameTableInFormula,
  type SheetSnapshot,
  hashPassword,
  hasPassword,
  passwordMatches,
  nextSheetName,
  sheetNameProblem,
  snapshotSheet,
  type CellWrite,
  type Op,
  type SheetProtection,
} from "@cells/core"
import type { SheetController } from "./controller"

/** Excel's messages (en-US). */
export const SHEET_MESSAGES = {
  lastVisible:
    "A workbook must contain at least one visible worksheet. To hide, delete, or move the selected sheets, you must first insert a new sheet or unhide a sheet that is already hidden.",
  deleteData: "Microsoft Excel will permanently delete this sheet. Do you want to continue?",
  passwordMismatch: "The confirmation password is not identical.",
  /** Shown when a change needs the workbook's structure (Review › Protect Workbook). */
  structureLocked: "Workbook is protected and cannot be changed.",
  wrongPassword:
    "The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.",
}

/**
 * Sheet tabs: switching, adding, deleting, renaming, moving, copying, colouring and hiding sheets. Each change is one
 * undo step that shows the sheet it touched again.
 */
export class SheetCommands {
  /** The sheet whose tab is being renamed in place (Format › Rename Sheet, a double-click on the tab). */
  renaming: number | null = null

  constructor(private readonly c: SheetController) {}

  beginRename(index = this.c.sheetIndex) {
    if (this.isWorkbookProtected()) return
    this.c.activate(index)
    this.renaming = index
    this.c.changed()
  }

  endRename() {
    if (this.renaming === null) return
    this.renaming = null
    this.c.changed()
  }

  get names(): string[] {
    return this.c.wb.sheets.map((s) => s.name)
  }

  /** Indices of the sheets with a tab. */
  visible(): number[] {
    const out: number[] = []
    this.c.wb.sheets.forEach((s, i) => !s.hidden && out.push(i))
    return out
  }

  hidden(): number[] {
    const out: number[] = []
    this.c.wb.sheets.forEach((s, i) => s.hidden && out.push(i))
    return out
  }

  /** The visible sheet at or after `index`, else before it. */
  nearestVisible(index: number): number {
    const sheets = this.c.wb.sheets
    for (let i = index; i < sheets.length; i++) if (!sheets[i].hidden) return i
    for (let i = index - 1; i >= 0; i--) if (!sheets[i].hidden) return i
    return 0
  }

  /** Ctrl+PgDn / Ctrl+PgUp: the next or previous visible sheet (no wrapping, as Excel). */
  step(dir: 1 | -1) {
    const sheets = this.c.wb.sheets
    for (let i = this.c.sheetIndex + dir; i >= 0 && i < sheets.length; i += dir) {
      if (!sheets[i].hidden) {
        this.c.activate(i)
        return
      }
    }
  }

  private run(ops: Op[], label: string, sheetAfter: number) {
    if (this.c.edit) this.c.commitEdit()
    this.c.transact(ops, { label, before: this.c.selection, sheetAfter })
    this.c.activate(sheetAfter)
  }

  /**
   * A new blank sheet. Insert › Insert Sheet and Shift+F11 put it before the active sheet; the + button after it.
   */
  insert(at = this.c.sheetIndex) {
    const name = nextSheetName(this.names)
    this.run([{ kind: "addSheet", sheet: at, name }], "Insert Sheet", at)
  }

  /** A new sheet after the active one holding imported cells (Data › From Text/CSV); a taken name gets " (2)". */
  importSheet(name: string, cells: CellWrite[], extra: (sheet: number) => Op[] = () => []) {
    const names = this.names.map((n) => n.toLowerCase())
    let unique = name
    for (let n = 2; names.includes(unique.toLowerCase()); n++) {
      const suffix = ` (${n})`
      unique = name.slice(0, 31 - suffix.length) + suffix
    }
    const at = this.c.sheetIndex + 1
    const ops: Op[] = [{ kind: "addSheet", sheet: at, name: unique }]
    if (cells.length) ops.push({ kind: "setCells", sheet: at, cells })
    ops.push(...extra(at))
    this.run(ops, "Import Text", at)
  }

  canRemove(index = this.c.sheetIndex): boolean {
    const sheets = this.c.wb.sheets
    return sheets.length > 1 && (sheets[index].hidden || this.visible().length > 1)
  }

  /**
   * Deletes a sheet (every grouped sheet, when `index` is in the group), asking first when one holds anything
   * (Excel). References to them become #REF!.
   */
  remove(index = this.c.sheetIndex) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet) return
    const targets = this.c.group.active && this.c.group.has(index) ? this.c.group.indexes() : [index]
    const sheets = this.c.wb.sheets
    const visibleLeft = this.visible().filter((i) => !targets.includes(i)).length
    if (sheets.length <= targets.length || (visibleLeft === 0 && targets.some((i) => !sheets[i].hidden))) {
      this.c.shell.alert(SHEET_MESSAGES.lastVisible)
      return
    }
    const go = () => {
      // The sheet after the first deleted one takes its place; at the end, the one before.
      const left = this.c.wb.sheets.filter((_, i) => !targets.includes(i))
      const first = Math.min(...targets)
      let next = first < left.length ? first : left.length - 1
      while (left[next]?.hidden && next < left.length - 1) next++
      while (left[next]?.hidden && next > 0) next--
      this.c.group.clear(false)
      const ops: Op[] = [...targets].sort((a, b) => b - a).map((i) => ({ kind: "removeSheet", sheet: i }))
      this.run(ops, "Delete Sheet", next)
    }
    const empty = targets.every((i) => !sheets[i].cells.usedBounds() && !sheets[i].merges.size)
    if (empty) go()
    else this.c.shell.confirm(SHEET_MESSAGES.deleteData, go)
  }

  /** Renames a sheet; returns Excel's message when the name is not allowed (nothing changes then). */
  rename(index: number, name: string): string | null {
    const sheet = this.c.wb.sheets[index]
    if (!sheet || name === sheet.name) return null
    const problem = sheetNameProblem(
      name,
      this.names.filter((_, i) => i !== index),
    )
    if (problem) return problem
    this.run([{ kind: "renameSheet", sheet: index, name }], "Rename Sheet", index)
    return null
  }

  /** Moves a sheet so it ends up before the sheet now at `before` (sheets.length: to the end). */
  move(index: number, before: number) {
    const to = before > index ? before - 1 : before
    if (to === index) return
    this.run([{ kind: "moveSheet", sheet: index, to }], "Move Sheet", to)
  }

  /** A copy of a sheet ("Sheet1 (2)"), placed before the sheet now at `before`. */
  copy(index: number, before: number) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet) return
    const snapshot = snapshotSheet(sheet)
    const base = sheet.name.replace(/ \(\d+\)$/, "")
    const taken = new Set(this.names.map((n) => n.toLowerCase()))
    const named = (n: number) => base.slice(0, 31 - ` (${n})`.length) + ` (${n})`
    let n = 2
    while (taken.has(named(n).toLowerCase())) n++
    const name = named(n)
    snapshot.name = name
    snapshot.hidden = false
    this.renameCopiedTables(snapshot)
    // Its text links get connections of their own (Excel: sales → sales1), linked to the same files.
    const connections = this.c.wb.sheets.flatMap((s) => s.queryTables.map((q) => q.connection))
    snapshot.queryTables = (snapshot.queryTables ?? []).map((q) => {
      const taken = new Set(connections.map((c) => c.toLowerCase()))
      let n = 1
      while (taken.has(`${q.connection}${n}`.toLowerCase())) n++
      const connection = `${q.connection}${n}`
      connections.push(connection)
      return { ...q, connection }
    })
    snapshot.objects = copiedSheetObjects(snapshot.objects ?? snapshot.charts ?? [], this.c.wb.sheets, sheet, name)
    delete snapshot.charts
    this.run([{ kind: "addSheet", sheet: before, name, snapshot }], "Copy Sheet", before)
  }

  /**
   * Tables in a copied sheet get names of their own (Sales → Sales_2) and new ids; the copy's formulas that name them
   * follow.
   */
  private renameCopiedTables(snapshot: SheetSnapshot) {
    const tables = snapshot.tables ?? []
    if (!tables.length) return
    const wb = this.c.wb
    const taken = new Set([
      ...wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toUpperCase())),
      ...wb.names.all().map((n) => n.name.toUpperCase()),
    ])
    let id = nextTableId(wb.sheets)
    for (const t of tables) {
      const base = t.name.replace(/_\d+$/, "")
      let n = 2
      while (taken.has(`${base}_${n}`.toUpperCase())) n++
      const name = `${base}_${n}`
      taken.add(name.toUpperCase())
      for (const w of snapshot.cells) {
        const input = w.data?.input
        const next = input ? renameTableInFormula(input, t.name, name) : null
        if (next !== null && w.data) w.data = { ...w.data, input: next }
      }
      for (const c of t.columns) if (c.formula) c.formula = renameTableInFormula(c.formula, t.name, name) ?? c.formula
      t.name = name
      t.id = id++
    }
  }

  /** Tab colour (null: No Color). */
  setColor(index: number, color: string | null) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet || (sheet.tabColor ?? null) === color) return
    this.run([{ kind: "sheetProps", sheet: index, tabColor: color }], "Tab Color", this.c.sheetIndex)
  }

  /** Page Layout › Background: a picture tiled behind the active sheet's cells (null: Delete Background). */
  setBackground(src: string | null) {
    const index = this.c.sheetIndex
    if ((this.c.sheet.background ?? null) === src) return
    this.run(
      [{ kind: "sheetProps", sheet: index, background: src }],
      src ? "Sheet Background" : "Delete Background",
      index,
    )
  }

  /** Hides a sheet (every grouped sheet, when `index` is in the group). */
  hide(index = this.c.sheetIndex) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet || sheet.hidden) return
    const targets = this.c.group.active && this.c.group.has(index) ? this.c.group.indexes() : [index]
    const visible = this.visible().filter((i) => !targets.includes(i))
    if (!visible.length) {
      this.c.shell.alert(SHEET_MESSAGES.lastVisible)
      return
    }
    const ops: Op[] = targets.map((i) => ({ kind: "sheetProps", sheet: i, hidden: true }))
    // Hiding the sheet shown moves to the next one (the one before when it was the last).
    let next = this.c.sheetIndex
    if (targets.includes(next)) next = visible.find((i) => i > index) ?? visible[visible.length - 1]
    this.c.group.clear(false)
    this.run(ops, "Hide Sheet", next)
  }

  unhide(index: number) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet?.hidden) return
    this.run([{ kind: "sheetProps", sheet: index, hidden: false }], "Unhide Sheet", index)
  }

  // ── Protection ───────────────────────────────────────────────────────────────────────────────────────────

  isProtected(index = this.c.sheetIndex): boolean {
    return Boolean(this.c.wb.sheets[index]?.protection)
  }

  /** Protect Sheet: the options as the dialog left them, with the password (empty: none). */
  protect(options: Omit<SheetProtection, "password">, password: string) {
    const protection: SheetProtection = { ...options, ...(password ? { password: hashPassword(password) } : {}) }
    this.run([{ kind: "protect", sheet: this.c.sheetIndex, protection }], "Protect Sheet", this.c.sheetIndex)
    this.c.keepSelectable()
  }

  /** Whether Unprotect Sheet needs to ask for a password. */
  needsPassword(index = this.c.sheetIndex): boolean {
    return hasPassword(this.c.wb.sheets[index]?.protection)
  }

  /** Unprotect Sheet: false (and Excel's message) when the password is wrong. */
  unprotect(password = ""): boolean {
    const sheet = this.c.sheet
    if (!sheet.protection) return true
    if (!passwordMatches(sheet.protection, password)) {
      this.c.shell.alert(SHEET_MESSAGES.wrongPassword)
      return false
    }
    this.run([{ kind: "protect", sheet: this.c.sheetIndex, protection: null }], "Unprotect Sheet", this.c.sheetIndex)
    return true
  }

  // ── Protect Workbook (Structure) ─────────────────────────────────────────────────────────────────────────

  isWorkbookProtected(): boolean {
    return this.c.wb.structure.protection !== null
  }

  workbookNeedsPassword(): boolean {
    return hasPassword(this.c.wb.structure.protection)
  }

  protectWorkbook(password: string) {
    const sel = this.c.selection
    this.c.transact([{ kind: "protectWorkbook", protection: password ? { password: hashPassword(password) } : {} }], {
      label: "Protect Workbook",
      before: sel,
      after: sel,
    })
  }

  /** Unprotect Workbook: false (and Excel's message) when the password is wrong. */
  unprotectWorkbook(password = ""): boolean {
    const p = this.c.wb.structure.protection
    if (!p) return true
    if (!passwordMatches(p, password)) {
      this.c.shell.alert(SHEET_MESSAGES.wrongPassword)
      return false
    }
    const sel = this.c.selection
    this.c.transact([{ kind: "protectWorkbook", protection: null }], {
      label: "Unprotect Workbook",
      before: sel,
      after: sel,
    })
    return true
  }
}
