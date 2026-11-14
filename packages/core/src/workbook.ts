import { applyOp, type Change, type Op } from "./ops"
import { Sheet } from "./sheet"
import { StyleTable } from "./styles"

/** One undoable step: a label for the Undo tooltip, the ops either way, and the caller's state either side. */
export type UndoEntry<Meta = unknown> = {
  label: string
  undo: Op[]
  redo: Op[]
  before?: Meta
  after?: Meta
}

export type TransactionOptions<Meta> = {
  label: string
  /** Caller state to restore on undo (the selection before the change). */
  before?: Meta
  /** Caller state to restore on redo (the selection after the change). */
  after?: Meta
  /** false: apply without an undo step (and without clearing redo). */
  history?: boolean
}

export type WorkbookListener = (changes: readonly Change[]) => void

/** Excel keeps 100 undo levels. */
export const DEFAULT_UNDO_LIMIT = 100

export class Workbook<Meta = unknown> {
  readonly sheets: Sheet[]
  readonly styles = new StyleTable()
  undoLimit: number
  private readonly undoStack: UndoEntry<Meta>[] = []
  private readonly redoStack: UndoEntry<Meta>[] = []
  private readonly listeners = new Set<WorkbookListener>()
  /** Bumped on every change; views compare it to know when to redraw. */
  version = 0

  constructor(sheetNames: string[] = ["Sheet1"], undoLimit = DEFAULT_UNDO_LIMIT) {
    this.sheets = sheetNames.map((name) => new Sheet(name))
    this.undoLimit = undoLimit
  }

  subscribe(listener: WorkbookListener) {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  /** Applies `ops` in order as one undo step. Returns the step (or null when nothing was applied). */
  transact(ops: Op[], options: TransactionOptions<Meta>): UndoEntry<Meta> | null {
    if (ops.length === 0) {
      return null
    }
    const { inverse, changes } = this.run(ops)
    const entry: UndoEntry<Meta> = {
      label: options.label,
      undo: inverse,
      redo: ops,
      before: options.before,
      after: options.after,
    }
    if (options.history !== false) {
      this.undoStack.push(entry)
      if (this.undoStack.length > this.undoLimit) {
        this.undoStack.splice(0, this.undoStack.length - this.undoLimit)
      }
      this.redoStack.length = 0
    }
    this.emit(changes)
    return entry
  }

  get canUndo() {
    return this.undoStack.length > 0
  }

  get canRedo() {
    return this.redoStack.length > 0
  }

  /** The label of the step Undo would take back. */
  get undoLabel(): string | null {
    return this.undoStack.at(-1)?.label ?? null
  }

  get redoLabel(): string | null {
    return this.redoStack.at(-1)?.label ?? null
  }

  get undoDepth() {
    return this.undoStack.length
  }

  undo(): UndoEntry<Meta> | null {
    const entry = this.undoStack.pop()
    if (!entry) {
      return null
    }
    const { inverse, changes } = this.run(entry.undo)
    entry.redo = inverse
    this.redoStack.push(entry)
    this.emit(changes)
    return entry
  }

  redo(): UndoEntry<Meta> | null {
    const entry = this.redoStack.pop()
    if (!entry) {
      return null
    }
    const { inverse, changes } = this.run(entry.redo)
    entry.undo = inverse
    this.undoStack.push(entry)
    this.emit(changes)
    return entry
  }

  clearHistory() {
    this.undoStack.length = 0
    this.redoStack.length = 0
  }

  /** Applies ops; the returned inverse undoes all of them (last op's inverse first). */
  private run(ops: Op[]): { inverse: Op[]; changes: Change[] } {
    const inverses: Op[][] = []
    const changes: Change[] = []
    for (const op of ops) {
      const applied = applyOp(this.sheets, op)
      inverses.push(applied.inverse)
      changes.push(applied.change)
    }
    this.version++
    return { inverse: inverses.reverse().flat(), changes }
  }

  private emit(changes: Change[]) {
    for (const listener of this.listeners) {
      listener(changes)
    }
  }
}
