import type { FilePort, FileRef } from "@cells/platform"
import {
  classifyOpenError,
  FileProblem,
  newDocument,
  openDocument,
  OPEN_TYPES,
  saveTypes,
  titleOf,
  type WorkbookDocument,
} from "./document"
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
}

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
  }

  /** Keeps a copy of the workbook if it changed since the last copy. */
  async autoRecover(): Promise<boolean> {
    const doc = this.doc
    if (!doc.dirty || doc.recoveredEdits === doc.edits || this.busy) return false
    const edits = doc.edits
    const bytes = await doc.bytes()
    await keepRecovery({
      id: doc.id,
      title: doc.title,
      fileName: doc.fileName,
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
    const choice = await this.ui.askSave(doc.title)
    if (choice === "cancel") return false
    if (choice === "save") return this.save()
    return true
  }

  private replace(next: WorkbookDocument) {
    const old = this.doc
    this.doc = next
    old.close()
    void dropRecovery(old.id)
    this.ui.show(next)
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
  async openRecent(entry: RecentFile): Promise<boolean> {
    const ref: FileRef = { name: entry.name, path: entry.path, handle: entry.handle }
    let bytes: Uint8Array
    try {
      bytes = await this.files.read(ref)
    } catch (error) {
      this.ui.problem(classifyOpenError(error, null, entry.name))
      return false
    }
    return this.openBytes(bytes, ref)
  }

  /** A file dropped on the window or passed by the system. */
  async openBytes(bytes: Uint8Array, ref: FileRef, remember = true): Promise<boolean> {
    if (this.busy) return false
    if (!(await this.confirmClose())) return false
    this.busy = true
    const label = `Opening ${ref.name}`
    this.ui.progress({ label, done: 0 })
    try {
      const doc = await openDocument(bytes, ref, (done) => this.ui.progress({ label, done }))
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

  private async write(ref: FileRef): Promise<boolean> {
    if (this.busy) return false
    const doc = this.doc
    this.busy = true
    this.ui.progress({ label: `Saving ${ref.name}`, done: null })
    try {
      const edits = doc.edits
      const bytes = await doc.bytes(ref.name)
      await this.files.write(ref, bytes)
      doc.saved(ref, edits)
      void dropRecovery(doc.id)
      if (ref.path || ref.handle) this.ui.recentChanged(touchRecent(ref, this.files.remembersFiles))
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
      const doc = await openDocument(file.bytes, { name: file.fileName })
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

  discardRecovered(file: RecoveredFile): Promise<void> {
    return dropRecovery(file.id)
  }
}
