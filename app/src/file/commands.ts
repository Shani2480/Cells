import { allQueryTables, type Op } from "@cells/core"
import { openXls, openXlsx, sheetNameFor, type TextImport } from "@cells/io"
import type { FilePort, FileRef } from "@cells/platform"
import {
  classifyOpenError,
  extensionOf,
  FileProblem,
  isTextFile,
  newDocument,
  openDocument,
  OPEN_TYPES,
  packageName,
  saveTypes,
  TEXT_EXTENSIONS,
  TEXT_TYPES,
  titleOf,
  type WorkbookDocument,
} from "./document"
import { fileNameOf } from "@/platform/android-files"
import { breakLinkOps, linkPaths, readLink, updateLinkOps, type SourceBook } from "@/sheet/links"
import { newQueryTable, refreshOps, sourceName, textCells } from "@/sheet/queries"
import { pinRecent, recentFiles, removeRecent, touchRecent, type RecentFile } from "./recent"
import { AUTORECOVER_MS, dropRecovery, keepRecovery, recoveredFiles, type RecoveredFile } from "./recovery"

/** "Want to save your changes to Book1?": Save, Don't Save or Cancel. */
export type SaveChoice = "save" | "discard" | "cancel"

/** What the commands need from the window: the shown workbook, questions, messages, progress. */
export interface FileUi {
  show(doc: WorkbookDocument): void
  askSave(title: string): Promise<SaveChoice>
  problem(problem: FileProblem): void
  /** null: done. */
  progress(state: { label: string; done: number | null } | null): void
  /** Things the person should know about a file that opened (features left out). */
  notice(lines: string[], fileName: string): void
  recentChanged(list: RecentFile[]): void
  /** The Text Import Wizard for a text file; null when cancelled. */
  textImport(bytes: Uint8Array, fileName: string): Promise<TextImport | null>
  /** A question with OK and Cancel; true for OK. */
  confirm(message: string): Promise<boolean>
}

/** Excel's message when a workbook of several sheets is saved as CSV or text. */
/** How long after a change a chart's workbook (Edit Data) is saved back for Inkspire. */
export const EMBEDDED_SAVE_MS = 800

export const ONE_SHEET_MESSAGE =
  "The selected file type does not support workbooks that contain multiple sheets. To save only the active sheet, click OK. To save all sheets, save them one at a time under different file names, or choose a file type that supports multiple sheets."

/**
 * File › New, Open, Save, Save As and Close, plus AutoRecover. One workbook per window: opening another file asks to
 * save unsaved changes first, then replaces the workbook (a new, untouched Book1 is simply replaced, as in Excel).
 */
export class FileCommands {
  doc: WorkbookDocument
  private busy = false
  /** New workbooks are numbered per window, as Excel's Book1, Book2… */
  private books = 1
  private recoverTimer: ReturnType<typeof setInterval> | null = null
  private embedded: { stop: () => void } | null = null
  /** Files picked for text links this session, by the link's source: a browser can read them again by handle. */
  private readonly linked = new Map<string, FileRef>()

  constructor(
    private readonly files: FilePort,
    private readonly ui: FileUi,
    first: WorkbookDocument,
  ) {
    this.doc = first
  }

  get recent(): RecentFile[] {
    return recentFiles(this.files.remembersFiles)
  }

  /** Starts AutoRecover (every two minutes while there are unsaved changes). */
  startAutoRecover(every = AUTORECOVER_MS) {
    this.recoverTimer ??= setInterval(() => void this.autoRecover(), every)
  }

  stop() {
    if (this.recoverTimer) clearInterval(this.recoverTimer)
    this.recoverTimer = null
    this.embedded?.stop()
    this.embedded = null
  }

  /** Keeps a copy of the workbook if it changed since the last copy. */
  async autoRecover(): Promise<boolean> {
    const doc = this.doc
    if (!doc.dirty || doc.recoveredEdits === doc.edits || this.busy) return false
    const edits = doc.edits
    // Always a whole workbook (a CSV's copy keeps every sheet and format).
    const bytes = await doc.bytes(packageName(doc.fileName))
    await keepRecovery({
      id: doc.id,
      title: doc.title,
      fileName: packageName(doc.fileName),
      path: doc.ref?.path,
      saved: Date.now(),
      bytes,
    })
    doc.recoveredEdits = edits
    return true
  }

  /** Asks about unsaved changes; false when the person cancels (or the save didn't happen). */
  async confirmClose(): Promise<boolean> {
    const doc = this.doc
    if (!doc.dirty) return true
    // A chart's workbook goes back to Inkspire: saved, never asked about.
    if (doc.embeddedFor && doc.ref) return this.write(doc.ref, true)
    const choice = await this.ui.askSave(doc.title)
    if (choice === "cancel") return false
    if (choice === "save") return this.save()
    return true
  }

  private replace(next: WorkbookDocument) {
    const old = this.doc
    this.embedded?.stop()
    this.embedded = null
    this.doc = next
    old.close()
    void dropRecovery(old.id)
    this.ui.show(next)
    void this.refreshOnOpen()
  }

  /** Ctrl+N: a new blank workbook (this window shows one workbook at a time). */
  async newWorkbook(): Promise<boolean> {
    if (!(await this.confirmClose())) return false
    this.replace(newDocument(undefined, `Book${++this.books}.xlsx`))
    return true
  }

  /** Ctrl+O: the system's Open dialog. */
  async open(): Promise<boolean> {
    let picked: { ref: FileRef; bytes: Uint8Array } | null
    try {
      picked = await this.files.pickOpen(OPEN_TYPES)
    } catch (error) {
      this.ui.problem(classifyOpenError(error, null, "the file"))
      return false
    }
    if (!picked) return false
    return this.openBytes(picked.bytes, picked.ref)
  }

  /** A file from the Recent list. */
  async openRecent(entry: Pick<RecentFile, "name" | "path" | "handle">, remember = true): Promise<boolean> {
    const ref: FileRef = { name: entry.name, path: entry.path, handle: entry.handle }
    let bytes: Uint8Array
    try {
      bytes = await this.files.read(ref)
    } catch (error) {
      this.ui.problem(classifyOpenError(error, null, entry.name))
      return false
    }
    return this.openBytes(bytes, ref, remember)
  }

  /**
   * A file by its full path (a cells:// link, a file opened from the file manager). The workbook already shown is
   * not opened again.
   */
  async openPath(path: string): Promise<boolean> {
    if (this.doc.ref?.path === path) return true
    const name = fileNameOf(path)
    return this.openRecent({ name, path })
  }

  /**
   * Inkspire's Edit Data: the workbook of a chart in `documentName`, which Inkspire copied out to `path` and reads
   * back. It isn't added to Recent; every change is saved to it a moment later, and closing doesn't ask.
   */
  async openEmbedded(path: string, documentName: string): Promise<boolean> {
    if (this.doc.ref?.path !== path) {
      const name = fileNameOf(path)
      if (!(await this.openRecent({ name, path }, false))) return false
    }
    const doc = this.doc
    doc.embeddedFor = documentName
    this.embedded?.stop()
    let timer: ReturnType<typeof setTimeout> | null = null
    const later = () => {
      if (timer) clearTimeout(timer)
      timer = setTimeout(() => {
        timer = null
        if (doc !== this.doc || !doc.embeddedFor || !doc.ref || !doc.dirty) return
        // Not while a cell is being typed in (saving would enter it) or another save runs.
        if (doc.controller.edit || this.busy) return later()
        void this.write(doc.ref, true)
      }, EMBEDDED_SAVE_MS)
    }
    const unsubscribe = doc.controller.wb.subscribe((_changes, ops) => {
      if (ops.length) later()
    })
    // Edits made before (the workbook was already open here) go back too.
    if (doc.dirty) later()
    this.embedded = {
      stop: () => {
        unsubscribe()
        if (timer) clearTimeout(timer)
      },
    }
    doc.controller.changed()
    return true
  }

  /** A file dropped on the window or passed by the system. A .txt or .tsv goes through the Text Import Wizard. */
  async openBytes(bytes: Uint8Array, ref: FileRef, remember = true): Promise<boolean> {
    if (this.busy) return false
    if (!(await this.confirmClose())) return false
    let text: TextImport | undefined
    if (isTextFile(ref.name) && extensionOf(ref.name) !== "csv") {
      const choices = await this.ui.textImport(bytes, ref.name)
      if (!choices) return false
      text = choices
    }
    this.busy = true
    const label = `Opening ${ref.name}`
    this.ui.progress({ label, done: 0 })
    try {
      const doc = await openDocument(bytes, ref, (done) => this.ui.progress({ label, done }), true, text)
      this.replace(doc)
      if (remember && doc.ref) this.ui.recentChanged(touchRecent(ref, this.files.remembersFiles))
      if (doc.warnings.length) this.ui.notice(doc.warnings, ref.name)
      return true
    } catch (error) {
      this.ui.problem(classifyOpenError(error, bytes, ref.name))
      return false
    } finally {
      this.busy = false
      this.ui.progress(null)
    }
  }

  /** Ctrl+S. A workbook never saved (or one the platform can't write back) goes to Save As. */
  async save(): Promise<boolean> {
    const doc = this.doc
    if (!doc.ref || !this.files.saveInPlace) return this.saveAs()
    return this.write(doc.ref)
  }

  /** F12: the system's Save As dialog. */
  async saveAs(): Promise<boolean> {
    const doc = this.doc
    let ref: FileRef | null
    try {
      ref = await this.files.pickSave(doc.fileName, saveTypes(doc.fileName))
    } catch (error) {
      this.ui.problem(new FileProblem("write", doc.fileName, error instanceof Error ? error.message : String(error)))
      return false
    }
    if (!ref) return false
    return this.write(ref)
  }

  /** `quiet`: no progress shown (a chart's workbook saved back as it changes). */
  private async write(ref: FileRef, quiet = false): Promise<boolean> {
    if (this.busy) return false
    const doc = this.doc
    // CSV and text hold one sheet: Excel saves the active one after asking.
    if (isTextFile(ref.name) && doc.controller.wb.sheets.length > 1 && !(await this.ui.confirm(ONE_SHEET_MESSAGE)))
      return false
    this.busy = true
    if (!quiet) this.ui.progress({ label: `Saving ${ref.name}`, done: null })
    try {
      const edits = doc.edits
      const bytes = await doc.bytes(ref.name)
      await this.files.write(ref, bytes)
      doc.saved(ref, edits)
      void dropRecovery(doc.id)
      if ((ref.path || ref.handle) && !doc.embeddedFor)
        this.ui.recentChanged(touchRecent(ref, this.files.remembersFiles))
      doc.controller.changed()
      return true
    } catch (error) {
      this.ui.problem(new FileProblem("write", ref.name, error instanceof Error ? error.message : String(error)))
      return false
    } finally {
      this.busy = false
      this.ui.progress(null)
    }
  }

  /**
   * File › Export › Create PDF (and Save As PDF): the PDF `make` produces, saved where the person picks. The workbook
   * stays the file it was.
   */
  async exportPdf(make: (progress: (done: number) => void) => Promise<Uint8Array>): Promise<boolean> {
    if (this.busy) return false
    const name = this.doc.fileName.replace(/\.[^.]+$/, "") + ".pdf"
    let ref: FileRef | null
    try {
      ref = await this.files.pickSave(name, [{ name: "PDF", extensions: ["pdf"] }])
    } catch (error) {
      this.ui.problem(new FileProblem("write", name, error instanceof Error ? error.message : String(error)))
      return false
    }
    if (!ref) return false
    this.busy = true
    const label = `Publishing ${ref.name}`
    this.ui.progress({ label, done: 0 })
    try {
      const bytes = await make((done) => this.ui.progress({ label, done }))
      await this.files.write(ref, bytes)
      return true
    } catch (error) {
      this.ui.problem(new FileProblem("write", ref.name, error instanceof Error ? error.message : String(error)))
      return false
    } finally {
      this.busy = false
      this.ui.progress(null)
    }
  }

  /** File › Close: the window then shows a new blank workbook. */
  async close(): Promise<boolean> {
    if (!(await this.confirmClose())) return false
    this.replace(newDocument(undefined, `Book${++this.books}.xlsx`))
    return true
  }

  removeRecent(entry: RecentFile) {
    this.ui.recentChanged(removeRecent(entry, this.files.remembersFiles))
  }

  pinRecent(entry: RecentFile, pinned: boolean) {
    this.ui.recentChanged(pinRecent(entry, pinned, this.files.remembersFiles))
  }

  /** Copies AutoRecover kept from a session that ended unexpectedly. */
  recovered(): Promise<RecoveredFile[]> {
    return recoveredFiles()
  }

  /** Opens a recovered copy: unsaved, named after the original ("Budget (Recovered)"), saved through Save As. */
  async openRecovered(file: RecoveredFile): Promise<boolean> {
    if (this.busy) return false
    if (!(await this.confirmClose())) return false
    try {
      const doc = await openDocument(file.bytes, { name: packageName(file.fileName) })
      doc.ref = null
      doc.fileName = `${titleOf(file.fileName)} (Recovered).xlsx`
      // Unsaved until the person saves it.
      doc.edits = 1
      this.replace(doc)
      await dropRecovery(file.id)
      return true
    } catch (error) {
      this.ui.problem(classifyOpenError(error, file.bytes, file.fileName))
      return false
    }
  }

  /**
   * Data › From Text/CSV: a CSV or text file, through the Text Import Wizard, onto a new sheet after the active one
   * (one step to undo).
   */
  async importText(): Promise<boolean> {
    let picked: { ref: FileRef; bytes: Uint8Array } | null
    try {
      picked = await this.files.pickOpen(TEXT_TYPES)
    } catch (error) {
      this.ui.problem(classifyOpenError(error, null, "the file"))
      return false
    }
    if (!picked) return false
    const choices = await this.ui.textImport(picked.bytes, picked.ref.name)
    if (!choices) return false
    const c = this.doc.controller
    const data = await textCells(c.wb.styles, picked.bytes, choices)
    // The data stays linked to its file (Data › Queries & Connections), so Refresh can read it again.
    const table = newQueryTable(c.wb.sheets, picked.ref.name, picked.ref.path ?? picked.ref.name, choices, data)
    this.linked.set(table.source, picked.ref)
    c.sheets.importSheet(sheetNameFor(picked.ref.name), data.cells, (sheet) => [
      { kind: "queryTables", sheet, tables: [table] },
    ])
    return true
  }

  /**
   * Refresh: reads a linked text file again into its range (one step to undo). A file that can't be read (moved, or
   * a browser that only knows its name) is looked for with the Open dialog, as Excel's Import Text File; false when
   * the person cancels.
   */
  async refreshQuery(sheet: number, name: string): Promise<boolean> {
    const ops = await this.refreshOpsFor(sheet, name)
    if (!ops) return false
    this.transactRefresh(ops, "Refresh")
    return true
  }

  /** Ctrl+Alt+F5: every linked file in the workbook, in one step to undo; how many were refreshed. */
  async refreshAll(): Promise<number> {
    const c = this.doc.controller
    const ops: Op[] = []
    let done = 0
    for (const { sheet, table } of allQueryTables(c.wb.sheets)) {
      const more = await this.refreshOpsFor(sheet, table.name)
      if (!more) continue
      ops.push(...more)
      done++
    }
    if (ops.length) this.transactRefresh(ops, "Refresh All")
    return done
  }

  /** Asks to refresh the links that are set to refresh when the file opens (Excel asks before reading files too). */
  async refreshOnOpen(): Promise<number> {
    const c = this.doc.controller
    const due = allQueryTables(c.wb.sheets).filter((q) => q.table.refreshOnLoad)
    if (!due.length) return 0
    const names = due.map((q) => sourceName(q.table.source)).join(", ")
    if (!(await this.ui.confirm(`This workbook refreshes its data from ${names} when it opens. Refresh it now?`)))
      return 0
    const ops: Op[] = []
    for (const { sheet, table } of due) ops.push(...((await this.refreshOpsFor(sheet, table.name)) ?? []))
    if (ops.length) this.transactRefresh(ops, "Refresh")
    return due.length
  }

  /**
   * Workbook Links › Update Values: reads the linked workbook again (from its path beside this workbook, then its
   * absolute path, then the Open dialog) and recalculates the formulas that use it. False when cancelled.
   */
  async updateLink(name: string): Promise<boolean> {
    const c = this.doc.controller
    const book = c.wb.links.list.find((b) => b.name.toUpperCase() === name.toUpperCase())
    for (const path of linkPaths(book, this.doc.ref?.path)) {
      const known = this.linked.get(path)
      try {
        const bytes = await this.files.read(known ?? { name: fileNameOf(path), path })
        const source = await this.readSourceBook(bytes, fileNameOf(path))
        if (!source) return false
        this.transactRefresh(
          updateLinkOps(c.wb, name, readLink(c.wb, name, source, this.linkLocation(path))),
          "Update Values",
        )
        return true
      } catch {
        // Not there: try the next place.
      }
    }
    return this.pickLink(name, "Update Values")
  }

  /** Workbook Links › Change Source: another workbook takes the link's place; formulas name the new file. */
  changeLinkSource(name: string): Promise<boolean> {
    return this.pickLink(name, "Change Source")
  }

  /** Workbook Links › Break Link, after Excel's warning: formulas that use the workbook become their values. */
  async breakLink(name: string): Promise<boolean> {
    const ok = await this.ui.confirm(
      "Breaking links permanently converts formulas and external references to their existing values. Because this cannot be undone in Excel, you may want to save a version of this file with a new name. Do you want to break the links?",
    )
    if (!ok) return false
    this.transactRefresh(breakLinkOps(this.doc.controller.wb, name), "Break Link")
    return true
  }

  private async pickLink(name: string, label: string): Promise<boolean> {
    let picked: { ref: FileRef; bytes: Uint8Array } | null
    try {
      picked = await this.files.pickOpen(
        OPEN_TYPES.filter((t) => !t.extensions.every((e) => TEXT_EXTENSIONS.includes(e))),
      )
    } catch (error) {
      this.ui.problem(classifyOpenError(error, null, name))
      return false
    }
    if (!picked) return false
    const source = await this.readSourceBook(picked.bytes, picked.ref.name)
    if (!source) return false
    const where = picked.ref.path ?? picked.ref.name
    this.linked.set(where, picked.ref)
    const c = this.doc.controller
    this.transactRefresh(updateLinkOps(c.wb, name, readLink(c.wb, name, source, this.linkLocation(where))), label)
    return true
  }

  /** Where a link points: relative to this workbook's folder when the file is beside it, as Excel stores it. */
  private linkLocation(path: string): { target: string; absolute?: string } {
    const own = this.doc.ref?.path
    const absolute = /^(?:[A-Za-z]:[\\/]|[\\/])/.test(path) ? path : undefined
    if (!absolute) return { target: path }
    const dir = own ? own.slice(0, Math.max(own.lastIndexOf("/"), own.lastIndexOf("\\")) + 1) : ""
    const target = dir && path.startsWith(dir) ? path.slice(dir.length) : path
    return { target, absolute }
  }

  private async readSourceBook(bytes: Uint8Array, name: string): Promise<SourceBook | null> {
    try {
      const ext = extensionOf(name)
      if (ext === "xls") return (await openXls(bytes, name, { calculate: false })).workbook
      return (await openXlsx(bytes, name, { calculate: false })).workbook
    } catch (error) {
      this.ui.problem(classifyOpenError(error, bytes, name))
      return null
    }
  }

  private transactRefresh(ops: Op[], label: string) {
    const c = this.doc.controller
    if (c.edit) c.commitEdit()
    c.transact(ops, { label, before: c.selection })
  }

  private async refreshOpsFor(index: number, name: string): Promise<Op[] | null> {
    const c = this.doc.controller
    const sheet = c.wb.sheets[index]
    const table = sheet?.queryTables.find((q) => q.name === name)
    if (!sheet || !table) return null
    let bytes: Uint8Array | null = null
    let source: string | undefined
    const known = this.linked.get(table.source)
    const ref =
      known ?? (/[\\/]|^content:/.test(table.source) ? { name: fileNameOf(table.source), path: table.source } : null)
    if (ref) {
      try {
        bytes = await this.files.read(ref)
      } catch {
        bytes = null
      }
    }
    if (!bytes) {
      let picked: { ref: FileRef; bytes: Uint8Array } | null
      try {
        picked = await this.files.pickOpen(TEXT_TYPES)
      } catch (error) {
        this.ui.problem(classifyOpenError(error, null, sourceName(table.source)))
        return null
      }
      if (!picked) return null
      bytes = picked.bytes
      source = picked.ref.path ?? picked.ref.name
      this.linked.set(source, picked.ref)
    }
    // The person may have changed the workbook while the file was read: look the table up again.
    const now = c.wb.sheets.indexOf(sheet)
    if (now < 0 || !sheet.queryTables.some((q) => q.name === name)) return null
    return refreshOps(now, sheet, name, await textCells(c.wb.styles, bytes, table.options), source)
  }

  discardRecovered(file: RecoveredFile): Promise<void> {
    return dropRecovery(file.id)
  }
}
