import {
  copyAnnotation,
  noteToComment,
  type Annotation,
  type AnnotationWrite,
  type CellComment,
  type CellNote,
  type Op,
} from "@cells/core"
import type { SheetController } from "./controller"
import { authorName } from "@/settings/settings"

export type CellAt = { row: number; col: number }

/** Excel's question before Notes › Convert to Comments. */
export const CONVERT_MESSAGE =
  "We will convert all your notes into comments. Comments use a reply thread so you can have conversations with other people. Convert all notes?"

/**
 * Review › Notes and Comments. A note is edited in its yellow box on the sheet; a comment thread opens in a card
 * beside the cell. The UI reads `editing` and `open` to know which box or card to show.
 */
export class AnnotationCommands {
  /** The note whose box is being typed in. */
  editing: CellAt | null = null
  /** The comment card that is open (pinned, not just hovered); `draft` while a new thread is being written. */
  open: (CellAt & { draft: boolean }) | null = null
  /** Show All Notes: every note's box stays on screen (Excel's toggle, per workbook view). */
  showAll = false
  /** A name set for this workbook view only (tests); otherwise File › Options › User name. */
  private authorOverride: string | null = null
  /** The name notes and comments are signed with (File › Options › User name in Excel). */
  get author(): string {
    return this.authorOverride ?? authorName()
  }
  set author(name: string) {
    this.authorOverride = name
  }

  constructor(private readonly c: SheetController) {}

  private get active(): CellAt {
    const sel = this.c.selection
    const merge = this.c.sheet.merges.at(sel.active.row, sel.active.col)
    return merge ? { row: merge.r0, col: merge.c0 } : { ...sel.active }
  }

  at(row: number, col: number): Annotation | undefined {
    return this.c.sheet.annotations.get(row, col)
  }

  noteAt(row: number, col: number): CellNote | undefined {
    const a = this.at(row, col)
    return a?.kind === "note" ? a : undefined
  }

  commentAt(row: number, col: number): CellComment | undefined {
    const a = this.at(row, col)
    return a?.kind === "comment" ? a : undefined
  }

  private write(writes: AnnotationWrite[], label: string, sheet = this.c.sheetIndex): boolean {
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    const ops: Op[] = [{ kind: "annotations", sheet, writes }]
    return this.c.transact(ops, { label, before: sel, after: sel }) !== null
  }

  // ── Notes ──────────────────────────────────────────────────────────────────────────────────────────────────

  /** New Note (Shift+F2): a note on the active cell, opened for typing; on a cell with a note, Edit Note. */
  newNote() {
    const { row, col } = this.active
    const existing = this.at(row, col)
    if (existing?.kind === "comment") {
      this.openThread(row, col)
      return
    }
    if (!existing) {
      const note: CellNote = { kind: "note", author: this.author, text: "", shown: false }
      if (!this.write([{ row, col, value: note }], "New Note")) return
    }
    this.editing = { row, col }
    this.c.changed()
  }

  editNote(row = this.active.row, col = this.active.col) {
    if (!this.noteAt(row, col)) return
    this.editing = { row, col }
    this.c.changed()
  }

  /** The note box loses focus: its text is kept (an unchanged text records nothing). */
  finishNote(text: string, size?: { width: number; height: number }) {
    const at = this.editing
    this.editing = null
    if (!at) return this.c.changed()
    const note = this.noteAt(at.row, at.col)
    if (!note) return this.c.changed()
    const resized = size && (size.width !== (note.width ?? 0) || size.height !== (note.height ?? 0))
    if (note.text === text && !resized) return this.c.changed()
    const next: CellNote = { ...note, text, ...(resized ? size : {}) }
    this.write([{ row: at.row, col: at.col, value: next }], "Edit Note")
  }

  /** Delete Note (context menu): every note in the selection. */
  deleteNotes() {
    this.removeIn("note", "Delete Note")
  }

  /** Show/Hide Note on the active cell. */
  toggleNote() {
    const { row, col } = this.active
    const note = this.noteAt(row, col)
    if (note) this.write([{ row, col, value: { ...note, shown: !note.shown } }], "Show/Hide Note")
  }

  /** Show All Notes: a view toggle; it doesn't change the notes. */
  toggleShowAll() {
    this.showAll = !this.showAll
    this.c.changed()
  }

  /** Notes › Convert to Comments: every note of the workbook becomes a thread (Excel asks first). */
  convertToComments() {
    const time = Date.now()
    const sel = this.c.selection
    const ops: Op[] = []
    this.c.wb.sheets.forEach((sheet, index) => {
      const writes = sheet.annotations
        .all()
        .filter((a) => a.value.kind === "note")
        .map((a) => ({ row: a.row, col: a.col, value: noteToComment(a.value as CellNote, time) }))
      if (writes.length) ops.push({ kind: "annotations", sheet: index, writes })
    })
    if (!ops.length) return
    this.c.shell.confirm(CONVERT_MESSAGE, () => {
      this.c.transact(ops, { label: "Convert to Comments", before: sel, after: sel })
    })
  }

  hasNotes(): boolean {
    return this.c.sheet.annotations.all().some((a) => a.value.kind === "note")
  }

  hasComments(): boolean {
    return this.c.sheet.annotations.all().some((a) => a.value.kind === "comment")
  }

  // ── Comments ───────────────────────────────────────────────────────────────────────────────────────────────

  /** New Comment (Ctrl+Shift+F2): the card opens with an empty box; on a cell with a thread, the thread opens. */
  newComment() {
    const { row, col } = this.active
    const existing = this.at(row, col)
    if (existing?.kind === "note") {
      this.editNote(row, col)
      return
    }
    this.open = { row, col, draft: !existing }
    this.c.changed()
  }

  openThread(row: number, col: number) {
    if (!this.commentAt(row, col)) return
    this.open = { row, col, draft: false }
    this.c.changed()
  }

  closeThread() {
    if (!this.open) return
    this.open = null
    this.c.changed()
  }

  /** Post: the first post starts the thread, later ones are replies. */
  post(text: string) {
    const at = this.open
    if (!at || !text.trim()) return
    const thread = this.commentAt(at.row, at.col)
    const post = { author: this.author, text, time: Date.now() }
    const next: CellComment = thread
      ? { ...copyAnnotation(thread), posts: [...thread.posts, post] }
      : { kind: "comment", resolved: false, posts: [post] }
    if (this.write([{ row: at.row, col: at.col, value: next }], thread ? "Reply" : "New Comment"))
      this.open = { ...at, draft: false }
    this.c.changed()
  }

  editPost(row: number, col: number, index: number, text: string) {
    const thread = this.commentAt(row, col)
    if (!thread || !thread.posts[index] || thread.posts[index].text === text || !text.trim()) return
    const next = copyAnnotation(thread)
    next.posts[index] = { ...next.posts[index], text, edited: true }
    this.write([{ row, col, value: next }], "Edit Comment")
  }

  /** Deleting the first post deletes the thread (Excel's Delete thread); a reply goes on its own. */
  deletePost(row: number, col: number, index: number) {
    const thread = this.commentAt(row, col)
    if (!thread) return
    if (index === 0) {
      this.deleteThread(row, col)
      return
    }
    const next = copyAnnotation(thread)
    next.posts.splice(index, 1)
    this.write([{ row, col, value: next }], "Delete Comment")
  }

  deleteThread(row: number, col: number) {
    if (this.write([{ row, col, value: null }], "Delete Comment") && this.open?.row === row && this.open.col === col)
      this.open = null
    this.c.changed()
  }

  resolve(row: number, col: number, resolved: boolean) {
    const thread = this.commentAt(row, col)
    if (thread)
      this.write([{ row, col, value: { ...thread, resolved } }], resolved ? "Resolve Thread" : "Reopen Thread")
  }

  /** Review › Delete: every comment in the selection. */
  deleteComments() {
    this.removeIn("comment", "Delete Comment")
  }

  private removeIn(kind: Annotation["kind"], label: string) {
    const op = this.removalOp([kind])
    if (op) this.write(op.writes, label)
  }

  /** Removing the selection's notes and/or comments (Clear All, Clear Notes, Clear Comments); null if none. */
  removalOp(kinds: Annotation["kind"][] = ["note", "comment"]): Extract<Op, { kind: "annotations" }> | null {
    const writes: AnnotationWrite[] = []
    for (const r of this.c.selection.ranges)
      for (const a of this.c.sheet.annotations.inRange(r))
        if (kinds.includes(a.value.kind) && !writes.some((w) => w.row === a.row && w.col === a.col))
          writes.push({ row: a.row, col: a.col, value: null })
    return writes.length ? { kind: "annotations", sheet: this.c.sheetIndex, writes } : null
  }

  /** Home › Find & Select › Notes: selects every cell with a note (or comment), as Excel's Go To Special does. */
  selectAll(kind: Annotation["kind"]) {
    const cells = this.c.sheet.annotations.all().filter((a) => a.value.kind === kind)
    if (!cells.length) {
      this.c.shell.alert("No cells were found.")
      return
    }
    const ranges = cells.map((a) => ({ r0: a.row, c0: a.col, r1: a.row, c1: a.col }))
    const first = cells[0]
    this.c.selection = {
      ranges,
      activeRange: 0,
      anchor: { row: first.row, col: first.col },
      focus: { row: first.row, col: first.col },
      active: { row: first.row, col: first.col },
    }
    this.c.scrollIntoView(this.c.selection.active)
    this.c.changed()
  }

  /** Clear Notes / Clear Comments. */
  clear(kind: Annotation["kind"]) {
    this.removeIn(kind, kind === "note" ? "Clear Notes" : "Clear Comments")
  }

  /** Whether the selection holds a note / a comment (what Delete and Show/Hide enable on). */
  selectionHas(kind: Annotation["kind"]): boolean {
    return this.c.selection.ranges.some((r) => this.c.sheet.annotations.inRange(r).some((a) => a.value.kind === kind))
  }

  // ── Previous / Next ───────────────────────────────────────────────────────────────────────────────────────

  /**
   * Previous/Next Note or Comment: the next one in reading order after the active cell, wrapping round the sheet.
   * A note opens for editing (Excel), a comment's card opens.
   */
  step(kind: Annotation["kind"], dir: 1 | -1) {
    const list = this.c.sheet.annotations.all().filter((a) => a.value.kind === kind)
    if (!list.length) return
    const { row, col } = this.active
    const after = (a: CellAt) => a.row > row || (a.row === row && a.col > col)
    const before = (a: CellAt) => a.row < row || (a.row === row && a.col < col)
    const target =
      dir === 1 ? (list.find(after) ?? list[0]) : ([...list].reverse().find(before) ?? list[list.length - 1])
    this.c.goTo({ r0: target.row, c0: target.col, r1: target.row, c1: target.col })
    if (kind === "note") this.editNote(target.row, target.col)
    else this.openThread(target.row, target.col)
  }
}
