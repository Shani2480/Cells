import {
  hashPassword,
  hasPassword,
  passwordMatches,
  nextSheetName,
  sheetNameProblem,
  snapshotSheet,
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

  canRemove(index = this.c.sheetIndex): boolean {
    const sheets = this.c.wb.sheets
    return sheets.length > 1 && (sheets[index].hidden || this.visible().length > 1)
  }

  /** Deletes a sheet, asking first when it holds anything (Excel). References to it become #REF!. */
  remove(index = this.c.sheetIndex) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet) return
    if (!this.canRemove(index)) {
      this.c.shell.alert(SHEET_MESSAGES.lastVisible)
      return
    }
    const go = () => {
      // The sheet after it takes its place; the last one hands over to the one before.
      const left = this.c.wb.sheets.filter((_, i) => i !== index)
      let next = index < left.length ? index : left.length - 1
      while (left[next]?.hidden && next < left.length - 1) next++
      while (left[next]?.hidden && next > 0) next--
      this.run([{ kind: "removeSheet", sheet: index }], "Delete Sheet", next)
    }
    const empty = !sheet.cells.usedBounds() && !sheet.merges.size
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
    this.run([{ kind: "addSheet", sheet: before, name, snapshot }], "Copy Sheet", before)
  }

  /** Tab colour (null: No Color). */
  setColor(index: number, color: string | null) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet || (sheet.tabColor ?? null) === color) return
    this.run([{ kind: "sheetProps", sheet: index, tabColor: color }], "Tab Color", this.c.sheetIndex)
  }

  hide(index = this.c.sheetIndex) {
    const sheet = this.c.wb.sheets[index]
    if (!sheet || sheet.hidden) return
    if (this.visible().length <= 1) {
      this.c.shell.alert(SHEET_MESSAGES.lastVisible)
      return
    }
    const ops: Op[] = [{ kind: "sheetProps", sheet: index, hidden: true }]
    // Hiding the sheet shown moves to the next one (the one before when it was the last).
    let next = this.c.sheetIndex
    if (index === next) {
      const visible = this.visible().filter((i) => i !== index)
      next = visible.find((i) => i > index) ?? visible[visible.length - 1]
    }
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
