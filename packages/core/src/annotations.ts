import { MAX_COLS } from "./address"
import type { Range } from "./address"

/**
 * Notes and comments, as Excel has them since 365: a note is the old yellow sticky note (author and text, shown on
 * hover or kept shown); a comment is a thread of posts with replies that can be resolved. A cell has at most one
 * of the two.
 */
export type CellNote = {
  kind: "note"
  author: string
  text: string
  /** Show/Hide Note: kept on screen instead of only on hover. */
  shown: boolean
  /** The box's size in px at 100% when the user resized it. */
  width?: number
  height?: number
}

export type CommentPost = {
  author: string
  text: string
  /** ms since the epoch. */
  time: number
  edited?: boolean
}

export type CellComment = {
  kind: "comment"
  resolved: boolean
  posts: CommentPost[]
}

export type Annotation = CellNote | CellComment

/** Default note box, as Excel draws a new note (about 2 columns by 4 rows). */
export const NOTE_WIDTH = 144
export const NOTE_HEIGHT = 79

const key = (row: number, col: number) => row * MAX_COLS + col

/** The sheet's notes and comments by cell. */
export class Annotations {
  private readonly map = new Map<number, { row: number; col: number; value: Annotation }>()

  get size() {
    return this.map.size
  }

  get(row: number, col: number): Annotation | undefined {
    return this.map.get(key(row, col))?.value
  }

  set(row: number, col: number, value: Annotation | null | undefined) {
    if (value) this.map.set(key(row, col), { row, col, value: copyAnnotation(value) })
    else this.map.delete(key(row, col))
  }

  /** Every annotation, in reading order (row by row), the order Excel's Next/Previous steps through. */
  all(): { row: number; col: number; value: Annotation }[] {
    return [...this.map.values()].sort((a, b) => a.row - b.row || a.col - b.col)
  }

  /** Annotations inside the range. */
  inRange(r: Range) {
    return this.all().filter((a) => a.row >= r.r0 && a.row <= r.r1 && a.col >= r.c0 && a.col <= r.c1)
  }

  clear() {
    this.map.clear()
  }

  /**
   * Inserting (count > 0) or deleting (count < 0) lines at `at`: annotations below or right move; those in deleted
   * lines go.
   */
  shift(axis: "rows" | "cols", at: number, count: number) {
    if (!this.map.size) return
    const items = [...this.map.values()]
    this.map.clear()
    for (const a of items) {
      let { row, col } = a
      const pos = axis === "rows" ? row : col
      if (count < 0 && pos >= at && pos < at - count) continue
      if (pos >= at) {
        if (axis === "rows") row += count
        else col += count
      }
      this.map.set(key(row, col), { row, col, value: a.value })
    }
  }
}

export function copyAnnotation<T extends Annotation>(a: T): T {
  return JSON.parse(JSON.stringify(a)) as T
}

/** One cell's annotation change; `value: null` removes it. */
export type AnnotationWrite = { row: number; col: number; value: Annotation | null }

/** Notes › Convert to Comments: Excel turns a note into a comment with the note's text as the first post. */
export function noteToComment(note: CellNote, time: number): CellComment {
  return { kind: "comment", resolved: false, posts: [{ author: note.author, text: note.text, time }] }
}
