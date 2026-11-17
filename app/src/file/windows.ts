import { SheetController } from "@/sheet/controller"

/**
 * View › Window: one workbook shown in more than one window, as Excel's New Window does. In Cells the windows are
 * panels inside the app's own window. Each is a SheetController on the same workbook, with its own sheet, selection,
 * scroll, zoom and view; edits in one show in all. The ribbon, formula bar, sheet tabs and status bar follow the
 * active window.
 *
 * Shown either maximized (only the active window, as Excel shows one window at a time by default) or arranged as
 * Arrange All's Tiled, Horizontal, Vertical or Cascade. View Side by Side puts two windows one above the other and
 * can scroll them together.
 */

export type Arrangement = "tiled" | "horizontal" | "vertical" | "cascade"

type Window = { c: SheetController; n: number; id: number }

let ids = 0

export class DocumentWindows {
  private list: Window[]
  private hiddenSet = new Set<SheetController>()
  private activeWindow: Window
  /** Front to back, for Cascade (the active window is first). */
  private stack: SheetController[]
  /** Null: maximized, only the active window shows. */
  arrangement: Arrangement | null = null
  /** View Side by Side: the two windows compared, and whether they scroll together. */
  sideBySide: { a: SheetController; b: SheetController; before: Arrangement | null } | null = null
  syncScroll = true
  private stopSync: (() => void) | null = null
  private listeners = new Set<() => void>()
  private version = 0

  constructor(
    main: SheetController,
    /** Gives a new window what the document gives its controller (unsaved-change counting, the file name). */
    private readonly adopt: (c: SheetController) => void = () => undefined,
  ) {
    this.activeWindow = { c: main, n: 1, id: ++ids }
    this.list = [this.activeWindow]
    this.stack = [main]
  }

  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  getVersion = () => this.version

  private changed() {
    this.version++
    for (const listener of [...this.listeners]) listener()
  }

  get active(): SheetController {
    return this.activeWindow.c
  }

  /** Every window, in the order they were opened. */
  get all(): SheetController[] {
    return this.list.map((w) => w.c)
  }

  /** The windows that aren't hidden, in the order they were opened. */
  get visible(): SheetController[] {
    return this.all.filter((c) => !this.hiddenSet.has(c))
  }

  get hidden(): SheetController[] {
    return this.all.filter((c) => this.hiddenSet.has(c))
  }

  isHidden(c: SheetController) {
    return this.hiddenSet.has(c)
  }

  /** Back to front for Cascade: the active window last, so it draws on top. */
  get cascadeOrder(): SheetController[] {
    return [...this.stack].reverse().filter((c) => !this.hiddenSet.has(c))
  }

  /** Whether the windows show as panels (more than one visible, arranged). */
  get arranged(): boolean {
    return this.arrangement !== null && this.visible.length > 1
  }

  /** A key that stays with the window while others close and the numbers change. */
  key(c: SheetController): number {
    return this.list.find((w) => w.c === c)?.id ?? 0
  }

  /** A window's number ("Book1:2"); null while the workbook has a single window (its title has none). */
  number(c: SheetController): number | null {
    if (this.list.length < 2) return null
    return this.list.find((w) => w.c === c)?.n ?? null
  }

  /** The title Excel gives a window: "Budget:2" when the workbook has more than one. */
  title(base: string, c: SheetController = this.active): string {
    const n = this.number(c)
    return n === null ? base : `${base}:${n}`
  }

  /** New Window: another window on the workbook, showing what the active one shows; it becomes the active window. */
  newWindow(): SheetController {
    const from = this.active
    if (from.edit) from.commitEdit()
    const c = new SheetController(from.wb)
    c.loadViews(from.saveViews(), from.sheetIndex)
    c.showFormulaBar = from.showFormulaBar
    this.adopt(c)
    const n = Math.max(...this.list.map((w) => w.n)) + 1
    const w = { c, n, id: ++ids }
    this.list.push(w)
    this.focus(w)
    this.changed()
    return c
  }

  activate(c: SheetController) {
    const w = this.list.find((x) => x.c === c)
    if (!w || w === this.activeWindow) return
    if (this.activeWindow.c.edit) this.activeWindow.c.commitEdit()
    this.hiddenSet.delete(c)
    this.focus(w)
    this.changed()
  }

  /** Ctrl+F6 and Ctrl+Shift+F6: the next or previous visible window. */
  cycle(step: 1 | -1) {
    const shown = this.visible
    if (shown.length < 2) return
    const i = shown.indexOf(this.active)
    this.activate(shown[(i + step + shown.length) % shown.length]!)
  }

  private focus(w: Window) {
    this.activeWindow = w
    this.stack = [w.c, ...this.stack.filter((c) => c !== w.c)]
  }

  /** Arrange All. */
  arrange(kind: Arrangement) {
    this.endSideBySide(false)
    this.arrangement = kind
    this.changed()
  }

  /** The window fills the area again; the others stay open behind it. */
  maximize() {
    this.endSideBySide(false)
    this.arrangement = null
    this.changed()
  }

  /** Closes a window (not the workbook); the last window can't be closed this way. */
  close(c: SheetController) {
    if (this.list.length < 2) return
    if (c.edit) c.commitEdit()
    if (this.sideBySide && (this.sideBySide.a === c || this.sideBySide.b === c)) this.endSideBySide(false)
    this.list = this.list.filter((w) => w.c !== c)
    this.stack = this.stack.filter((x) => x !== c)
    this.hiddenSet.delete(c)
    // Excel numbers what is left 1, 2, … again.
    this.list.forEach((w, i) => (w.n = i + 1))
    if (this.activeWindow.c === c) {
      const next = this.stack.find((x) => !this.hiddenSet.has(x)) ?? this.list[0]!.c
      this.hiddenSet.delete(next)
      this.focus(this.list.find((w) => w.c === next)!)
    }
    if (this.list.length < 2) this.arrangement = null
    c.dispose()
    this.changed()
  }

  /** Hide: the active window goes out of sight; another one becomes active. Not for the only visible window. */
  hide() {
    const shown = this.visible
    if (shown.length < 2) return
    const c = this.active
    if (c.edit) c.commitEdit()
    if (this.sideBySide && (this.sideBySide.a === c || this.sideBySide.b === c)) this.endSideBySide(false)
    this.hiddenSet.add(c)
    const next = this.stack.find((x) => !this.hiddenSet.has(x))!
    this.focus(this.list.find((w) => w.c === next)!)
    this.changed()
  }

  unhide(c: SheetController) {
    if (!this.hiddenSet.has(c)) return
    this.hiddenSet.delete(c)
    this.focus(this.list.find((w) => w.c === c)!)
    this.changed()
  }

  /** View Side by Side: the active window above `other`, scrolling together until it's turned off. */
  startSideBySide(other: SheetController) {
    const a = this.active
    if (other === a || !this.list.some((w) => w.c === other)) return
    this.hiddenSet.delete(other)
    this.sideBySide = { a, b: other, before: this.arrangement }
    this.arrangement = "horizontal"
    this.syncScroll = true
    this.listenForSync()
    this.changed()
  }

  /** Turns View Side by Side off: the windows go back to how they were shown before it. */
  endSideBySide(notify = true) {
    if (!this.sideBySide) return
    this.arrangement = this.sideBySide.before
    this.sideBySide = null
    this.stopSync?.()
    this.stopSync = null
    if (notify) this.changed()
  }

  toggleSyncScroll() {
    if (!this.sideBySide) return
    this.syncScroll = !this.syncScroll
    this.changed()
  }

  /** Reset Window Position: the two windows compared share the area equally again. */
  resetPosition() {
    if (!this.sideBySide) return
    this.arrangement = "horizontal"
    this.changed()
  }

  /** Whether a window is shown (maximized: only the active one). */
  shows(c: SheetController) {
    if (this.hiddenSet.has(c)) return false
    if (this.sideBySide) return c === this.sideBySide.a || c === this.sideBySide.b
    return this.arranged || c === this.active
  }

  /** The windows drawn in the area, in the order the arrangement lays them out. */
  get shown(): SheetController[] {
    return this.visible.filter((c) => this.shows(c))
  }

  /** Synchronous Scrolling: what one window scrolls, the other scrolls by the same rows and columns. */
  private listenForSync() {
    this.stopSync?.()
    const pair = this.sideBySide
    if (!pair) return
    const at = (c: SheetController) => ({ row: c.sheet.rows.indexAt(c.scrollY), col: c.sheet.cols.indexAt(c.scrollX) })
    let last = new Map([
      [pair.a, at(pair.a)],
      [pair.b, at(pair.b)],
    ])
    let moving = false
    const follow = (from: SheetController, to: SheetController) => () => {
      if (moving) return
      const was = last.get(from)!
      const now = at(from)
      if (this.syncScroll && (now.row !== was.row || now.col !== was.col)) {
        const there = at(to)
        const row = Math.max(0, there.row + now.row - was.row)
        const col = Math.max(0, there.col + now.col - was.col)
        moving = true
        to.scrollTo(to.sheet.cols.offset(col), to.sheet.rows.offset(row))
        moving = false
      }
      last = new Map([
        [pair.a, at(pair.a)],
        [pair.b, at(pair.b)],
      ])
    }
    const offA = pair.a.subscribeScroll(follow(pair.a, pair.b))
    const offB = pair.b.subscribeScroll(follow(pair.b, pair.a))
    this.stopSync = () => {
      offA()
      offB()
    }
  }

  /** The document closed: every window but the first stops following the workbook. */
  dispose() {
    this.stopSync?.()
    for (const w of this.list.slice(1)) w.c.dispose()
  }
}
