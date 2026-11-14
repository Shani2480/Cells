import {
  MAX_COLS,
  MAX_ROWS,
  addRange,
  blockFromText,
  cellName,
  clampCell,
  clearContentsOp,
  currentRegion,
  dataEdge,
  decodeTSV,
  editText,
  encodeHTML,
  encodeTSV,
  expandToColumns,
  expandToRows,
  extendSelection,
  fillInputOp,
  fillOp,
  formatValue,
  isSingleCell,
  isWholeColumns,
  isWholeRows,
  lastCell,
  moveBlockOps,
  moveWithinSelection,
  pasteArea,
  pasteBlockOp,
  rangeCols,
  rangeName,
  rangeOf,
  rangeRows,
  readBlock,
  sameRange,
  selectAll,
  selectCell,
  selectCols,
  selectRange,
  selectRows,
  selectionBounds,
  step,
  writeInputOp,
  type CellAddress,
  type CellBlock,
  type Direction,
  type Op,
  type Range,
  type Selection,
  type Sheet,
  Workbook,
} from "@cells/core"

/** Ready: the grid has the keys. Enter: typing a fresh entry (arrows commit). Edit: F2 editing (arrows move the caret). */
export type CellMode = "ready" | "enter" | "edit"

export type EditState = {
  row: number
  col: number
  text: string
  mode: "enter" | "edit"
  /** Where typing happens: in the cell or in the formula bar (Excel keeps both in step). */
  surface: "cell" | "bar"
}

/** The marching-ants range after Copy or Cut, until Esc, a new entry or a paste of a cut. */
export type ClipState = { range: Range; cut: boolean; text: string; block: CellBlock }

export type KeyInput = {
  key: string
  /** Ctrl on Linux/Windows, ⌘ on macOS. */
  primary: boolean
  shift: boolean
  alt: boolean
}

export type InsertDeleteRequest = { kind: "insert" | "delete" }

/** Things the controller asks the shell to do (dialogs and the like). */
export type ShellRequests = {
  goTo(): void
  insertDelete(request: InsertDeleteRequest): void
}

export const HEADER_HEIGHT = 20

/**
 * The state and behaviour of one sheet view: selection, in-cell editing, scrolling, clipboard and every keyboard
 * command, on top of the core workbook. No React and no DOM: views subscribe and draw.
 */
export class SheetController {
  readonly wb: Workbook<Selection>
  sheetIndex = 0
  selection: Selection = selectCell({ row: 0, col: 0 })
  edit: EditState | null = null
  clip: ClipState | null = null
  endMode = false
  /** Preview of a fill-handle drag (the area it will fill). */
  fillPreview: Range | null = null
  /** Scroll position in sheet pixels at 100% (top-left of the cell area). */
  scrollX = 0
  scrollY = 0
  /** Size of the cell area (without headers), set by the view. */
  viewWidth = 800
  viewHeight = 600
  rowHeaderWidth = 40
  shell: ShellRequests = { goTo() {}, insertDelete() {} }
  /** Bumped on every change the view may need to draw; the React hook compares it. */
  version = 0

  private readonly listeners = new Set<() => void>()
  /** Tab moves remember where they started, so Enter returns to that column on the next row (Excel). */
  private tabStartCol: number | null = null
  /** Ctrl+A in data selects the region first, the whole sheet on the next press. */
  private lastCtrlARegion: Range | null = null

  constructor(wb: Workbook<Selection> = new Workbook<Selection>()) {
    this.wb = wb
    wb.subscribe(() => this.changed())
  }

  get sheet(): Sheet {
    return this.wb.sheets[this.sheetIndex]
  }

  get mode(): CellMode {
    return this.edit ? this.edit.mode : "ready"
  }

  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  getVersion = () => this.version

  changed() {
    this.version++
    for (const listener of this.listeners) listener()
  }

  // ── Selection ──────────────────────────────────────────────────────────────────────────────────────────────

  setSelection(selection: Selection, reveal: CellAddress | null = selection.focus) {
    this.selection = selection
    this.lastCtrlARegion = null
    if (reveal) this.scrollIntoView(reveal)
    this.changed()
  }

  selectCellAt(address: CellAddress) {
    this.tabStartCol = null
    this.setSelection(selectCell(clampCell(address)))
  }

  /** The Name Box text: the active cell, or "3R x 2C" while a drag is in progress. */
  nameBoxText(dragging = false): string {
    if (dragging) {
      const range = this.selection.ranges[this.selection.activeRange]
      if (!sameRange(range, rangeOf(this.selection.active))) {
        if (isWholeColumns(range) && isWholeRows(range)) return `${MAX_ROWS}R x ${MAX_COLS}C`
        if (isWholeColumns(range)) return `${rangeCols(range)}C`
        if (isWholeRows(range)) return `${rangeRows(range)}R`
        return `${rangeRows(range)}R x ${rangeCols(range)}C`
      }
    }
    return cellName(this.selection.active)
  }

  /** The formula bar's text: what is being typed, or the active cell's input. */
  formulaBarText(): string {
    if (this.edit) return this.edit.text
    const { row, col } = this.selection.active
    return this.sheet.cells.input(row, col) ?? editText(this.sheet.cells.value(row, col))
  }

  // ── Editing ────────────────────────────────────────────────────────────────────────────────────────────────

  /** Starts an entry: typing replaces (enter mode) or F2/double-click edits the content (edit mode). */
  beginEdit(mode: "enter" | "edit", text?: string, surface: "cell" | "bar" = "cell") {
    const { row, col } = this.selection.active
    this.edit = {
      row,
      col,
      mode,
      surface,
      text: text ?? (mode === "edit" ? this.formulaBarText() : ""),
    }
    this.endMode = false
    this.scrollIntoView(this.selection.active)
    this.changed()
  }

  setEditText(text: string) {
    if (!this.edit) return
    this.edit = { ...this.edit, text }
    this.changed()
  }

  setEditSurface(surface: "cell" | "bar") {
    if (this.edit && this.edit.surface !== surface) {
      this.edit = { ...this.edit, surface }
      this.changed()
    }
  }

  toggleEditMode() {
    if (!this.edit) return
    this.edit = { ...this.edit, mode: this.edit.mode === "enter" ? "edit" : "enter" }
    this.changed()
  }

  cancelEdit() {
    if (!this.edit) return
    this.edit = null
    this.changed()
  }

  /**
   * Commits the entry. `fill` (Ctrl+Enter) writes it to every selected cell. Returns false when nothing was being
   * edited. A new entry clears the marching ants, as in Excel.
   */
  commitEdit(fill = false): boolean {
    const edit = this.edit
    if (!edit) return false
    this.edit = null
    const before = this.selection
    let op: Op
    if (fill && !isSingleCell(this.selection)) {
      op = fillInputOp(this.sheetIndex, this.sheet, this.wb.styles, this.selection.ranges, edit.text)
    } else {
      op = writeInputOp(this.sheetIndex, this.sheet, this.wb.styles, edit.row, edit.col, edit.text)
    }
    const label = edit.text ? `Typing '${truncate(edit.text)}' in ${cellName(edit)}` : `Clear ${cellName(edit)}`
    this.clip = null
    this.wb.transact([op], { label, before, after: before })
    return true
  }

  // ── Movement ───────────────────────────────────────────────────────────────────────────────────────────────

  /** Arrow keys (with Ctrl for data edges, Shift to extend). */
  arrow(dir: Direction, ctrl: boolean, shift: boolean) {
    this.tabStartCol = null
    const jump = ctrl || this.endMode
    this.endMode = false
    if (shift) {
      const focus = jump ? dataEdge(this.sheet, this.selection.focus, dir) : step(this.sheet, this.selection.focus, dir)
      this.setSelection(extendSelection(this.selection, keepWholeLines(this.selection, focus)), focus)
    } else {
      const target = jump
        ? dataEdge(this.sheet, this.selection.active, dir)
        : step(this.sheet, this.selection.active, dir)
      this.setSelection(selectCell(target))
    }
  }

  /** Enter / Shift+Enter / Tab / Shift+Tab: move within a multi-cell selection, or step (with Excel's Tab return). */
  enterMove(dir: Direction) {
    const single = isSingleCell(this.selection)
    if (single && dir === "down" && this.tabStartCol !== null) {
      const row = step(this.sheet, this.selection.active, "down").row
      const target = { row, col: this.tabStartCol }
      this.tabStartCol = null
      this.setSelection(selectCell(target))
      return
    }
    if (single && (dir === "right" || dir === "left")) {
      this.tabStartCol ??= this.selection.active.col
    } else if (single) {
      this.tabStartCol = null
    }
    const next = moveWithinSelection(this.sheet, this.selection, dir)
    this.selection = next
    this.lastCtrlARegion = null
    this.scrollIntoView(next.active)
    this.changed()
  }

  home(ctrl: boolean, shift: boolean) {
    this.tabStartCol = null
    this.endMode = false
    const from = shift ? this.selection.focus : this.selection.active
    const target = ctrl ? { row: 0, col: 0 } : { row: from.row, col: 0 }
    if (shift) this.setSelection(extendSelection(this.selection, target), target)
    else this.setSelection(selectCell(target))
  }

  ctrlEnd(shift: boolean) {
    this.tabStartCol = null
    this.endMode = false
    const target = lastCell(this.sheet)
    if (shift) this.setSelection(extendSelection(this.selection, target), target)
    else this.setSelection(selectCell(target))
  }

  /** Page Down/Up (Alt: a screen right/left): moves the active cell and the view by one screen. */
  page(dir: 1 | -1, horizontal: boolean, shift: boolean) {
    this.tabStartCol = null
    this.endMode = false
    const axis = horizontal ? this.sheet.cols : this.sheet.rows
    const span = horizontal ? this.viewWidth : this.viewHeight
    const from = shift ? this.selection.focus : this.selection.active
    const index = horizontal ? from.col : from.row
    const targetPx = axis.offset(index) + dir * span
    const targetIndex = Math.max(0, Math.min(axis.count - 1, axis.indexAt(Math.max(0, targetPx))))
    const target = horizontal ? { row: from.row, col: targetIndex } : { row: targetIndex, col: from.col }
    if (horizontal) this.scrollX = Math.max(0, this.scrollX + dir * span)
    else this.scrollY = Math.max(0, this.scrollY + dir * span)
    if (shift) this.setSelection(extendSelection(this.selection, target), target)
    else this.setSelection(selectCell(target), target)
  }

  toggleEndMode() {
    this.endMode = !this.endMode
    this.changed()
  }

  /** Ctrl+A: the current region first (when in data), then the whole sheet. */
  selectAllOrRegion() {
    const region = currentRegion(this.sheet, this.selection.active)
    const isData = region.r0 !== region.r1 || region.c0 !== region.c1 || this.sheet.cells.hasValue(region.r0, region.c0)
    const current = this.selection.ranges[this.selection.ranges.length - 1]
    if (isData && !(this.lastCtrlARegion && sameRange(current, this.lastCtrlARegion))) {
      const selection = selectRange(region, this.selection.active)
      this.setSelection(selection, null)
      this.lastCtrlARegion = region
      return
    }
    this.setSelection(selectAll(this.selection.active), null)
  }

  /** Ctrl+Shift+* : the current region. */
  selectCurrentRegion() {
    this.setSelection(selectRange(currentRegion(this.sheet, this.selection.active), this.selection.active), null)
  }

  selectColumns() {
    this.setSelection(expandToColumns(this.selection), null)
  }

  selectRowsOfSelection() {
    this.setSelection(expandToRows(this.selection), null)
  }

  // ── Mouse selection (called by the view) ──────────────────────────────────────────────────────────────────

  pointerCell(address: CellAddress, options: { shift: boolean; ctrl: boolean }) {
    this.tabStartCol = null
    this.endMode = false
    if (options.shift) {
      this.setSelection(extendSelection(this.selection, address), address)
    } else if (options.ctrl) {
      this.setSelection(addRange(this.selection, address), address)
    } else {
      this.setSelection(selectCell(address), address)
    }
  }

  /** Dragging with the button down: the last range follows the pointer. */
  dragTo(address: CellAddress) {
    const focus = keepWholeLines(this.selection, clampCell(address))
    if (focus.row === this.selection.focus.row && focus.col === this.selection.focus.col) return
    this.setSelection(extendSelection(this.selection, focus), address)
  }

  pointerRowHeader(row: number, options: { shift: boolean; ctrl: boolean }) {
    const leftCol = this.sheet.cols.indexAt(this.scrollX)
    if (options.shift) {
      const anchor = this.selection.anchor.row
      this.setSelection(
        extendSelection({ ...this.selection, anchor: { row: anchor, col: 0 } }, { row, col: MAX_COLS - 1 }),
        null,
      )
    } else if (options.ctrl) {
      this.setSelection(
        addRange(this.selection, { row, col: leftCol }, { r0: row, r1: row, c0: 0, c1: MAX_COLS - 1 }),
        null,
      )
    } else {
      this.setSelection(selectRows(row, row, leftCol), null)
    }
  }

  pointerColHeader(col: number, options: { shift: boolean; ctrl: boolean }) {
    const topRow = this.sheet.rows.indexAt(this.scrollY)
    if (options.shift) {
      const anchor = this.selection.anchor.col
      this.setSelection(
        extendSelection({ ...this.selection, anchor: { row: 0, col: anchor } }, { row: MAX_ROWS - 1, col }),
        null,
      )
    } else if (options.ctrl) {
      this.setSelection(
        addRange(this.selection, { row: topRow, col }, { r0: 0, r1: MAX_ROWS - 1, c0: col, c1: col }),
        null,
      )
    } else {
      this.setSelection(selectCols(col, col, topRow), null)
    }
  }

  selectEverything() {
    this.setSelection(
      selectAll({ row: this.sheet.rows.indexAt(this.scrollY), col: this.sheet.cols.indexAt(this.scrollX) }),
      null,
    )
  }

  // ── Content commands ───────────────────────────────────────────────────────────────────────────────────────

  /** Delete: clears contents, keeps formats. */
  clearContents() {
    const op = clearContentsOp(this.sheetIndex, this.sheet, this.selection.ranges)
    if (op) {
      this.wb.transact([op], { label: "Clear", before: this.selection, after: this.selection })
    }
  }

  undo() {
    if (this.edit) return
    const entry = this.wb.undo()
    if (entry?.before) this.restoreSelection(entry.before)
  }

  redo() {
    if (this.edit) return
    const entry = this.wb.redo()
    if (entry?.after) this.restoreSelection(entry.after)
  }

  private restoreSelection(selection: Selection) {
    this.selection = selection
    this.clip = null
    this.scrollIntoView(selection.active)
    this.changed()
  }

  /** Ctrl+D / Ctrl+R: fill down/right from the first row/column of the selection. */
  fillFromEdge(dir: "down" | "right") {
    const ops: Op[] = []
    for (const range of this.selection.ranges) {
      const source = dir === "down" ? { ...range, r1: range.r0 } : { ...range, c1: range.c0 }
      // With one row (or column) selected, Excel fills from the one above (or to the left).
      if ((dir === "down" && range.r0 === range.r1) || (dir === "right" && range.c0 === range.c1)) {
        const from =
          dir === "down"
            ? { ...range, r0: range.r0 - 1, r1: range.r0 - 1 }
            : { ...range, c0: range.c0 - 1, c1: range.c0 - 1 }
        if (from.r0 < 0 || from.c0 < 0) continue
        ops.push(pasteBlockOp(this.sheetIndex, range, readBlock(this.sheet, from)))
        continue
      }
      ops.push(pasteBlockOp(this.sheetIndex, range, readBlock(this.sheet, source)))
    }
    if (ops.length) {
      this.wb.transact(ops, {
        label: dir === "down" ? "Fill Down" : "Fill Right",
        before: this.selection,
        after: this.selection,
      })
    }
  }

  /** The fill handle was dropped: `target` is the grown selection. */
  fillTo(target: Range, invert: boolean) {
    const source = this.selection.ranges[this.selection.ranges.length - 1]
    this.fillPreview = null
    if (sameRange(source, target)) {
      this.changed()
      return
    }
    const shrinking = target.r1 < source.r1 || target.c1 < source.c1
    if (shrinking && target.r0 === source.r0 && target.c0 === source.c0) {
      // Dragging the handle back inside the selection clears what it leaves (Excel's gray-out erase).
      const cleared: Range = target.r1 < source.r1 ? { ...source, r0: target.r1 + 1 } : { ...source, c0: target.c1 + 1 }
      const op = clearContentsOp(this.sheetIndex, this.sheet, [cleared])
      const after = selectRange(target, this.selection.active)
      if (op) this.wb.transact([op], { label: "Clear", before: this.selection, after })
      this.setSelection(after, null)
      return
    }
    const op = fillOp(this.sheetIndex, this.sheet, source, target, invert)
    const after = selectRange(target, this.selection.active)
    if (op) this.wb.transact([op], { label: "AutoFill", before: this.selection, after })
    this.setSelection(after, null)
  }

  // ── Clipboard ──────────────────────────────────────────────────────────────────────────────────────────────

  /** Copy/Cut: what goes on the system clipboard, and the marching ants. Multiple ranges cannot be copied. */
  copy(cut: boolean): { text: string; html: string } | null {
    if (this.selection.ranges.length !== 1) return null
    const range = clampToUsed(this.sheet, this.selection.ranges[0])
    const block = readBlock(this.sheet, range)
    const rows = block.cells.map((row) =>
      Array.from({ length: block.cols }, (_, c) => {
        const cell = row[c]
        return formatValue(cell?.value, this.wb.styles.get(cell?.style ?? 0).numberFormat) ?? ""
      }),
    )
    const text = encodeTSV(rows)
    this.clip = { range, cut, text, block }
    this.changed()
    return { text, html: encodeHTML(rows) }
  }

  clearClip() {
    if (this.clip) {
      this.clip = null
      this.changed()
    }
  }

  /**
   * Paste: our own copy (full content and formats) when the clipboard still holds the text we put there, else the
   * clipboard's text parsed as typed input. A cut moves its cells and ends the marching ants.
   */
  paste(text: string | null, html?: string | null): boolean {
    if (this.selection.ranges.length !== 1) return false
    const target = this.selection.ranges[0]
    const own = this.clip && text !== null && normalizeNewlines(text) === normalizeNewlines(this.clip.text)
    let block: CellBlock
    if (own && this.clip) {
      block = this.clip.block
    } else if (text !== null && text !== "") {
      block = blockFromText(this.wb.styles, decodeTSV(text))
    } else if (html) {
      block = blockFromText(this.wb.styles, htmlTableRows(html))
    } else {
      return false
    }
    const before = this.selection
    if (own && this.clip?.cut) {
      const area = { r0: target.r0, c0: target.c0, r1: target.r0 + block.rows - 1, c1: target.c0 + block.cols - 1 }
      const ops = moveBlockOps(this.sheetIndex, this.sheet, this.clip.range, area)
      const after = selectRange(area, { row: area.r0, col: area.c0 })
      this.clip = null
      this.wb.transact(ops, { label: "Move", before, after })
      this.setSelection(after, null)
      return true
    }
    const area = pasteArea(target, block)
    const after = selectRange(area, this.selection.active)
    this.wb.transact([pasteBlockOp(this.sheetIndex, area, block)], { label: "Paste", before, after })
    this.setSelection(after, null)
    return true
  }

  // ── Rows and columns ───────────────────────────────────────────────────────────────────────────────────────

  /** The rows (or columns) a command acts on: every row touched by the selection, merged. */
  selectedLines(axis: "rows" | "cols"): [number, number][] {
    const spans = this.selection.ranges.map((range) =>
      axis === "rows" ? ([range.r0, range.r1] as [number, number]) : ([range.c0, range.c1] as [number, number]),
    )
    spans.sort((a, b) => a[0] - b[0])
    const merged: [number, number][] = []
    for (const span of spans) {
      const last = merged.at(-1)
      if (last && span[0] <= last[1] + 1) last[1] = Math.max(last[1], span[1])
      else merged.push([...span])
    }
    return merged
  }

  insertLines(axis: "rows" | "cols") {
    // Insert from the bottom up so earlier spans keep their indexes.
    const spans = this.selectedLines(axis).reverse()
    const ops: Op[] = spans.map(([from, to]) => ({
      kind: axis === "rows" ? "insertRows" : "insertCols",
      sheet: this.sheetIndex,
      at: from,
      count: to - from + 1,
    }))
    this.wb.transact(ops, {
      label: axis === "rows" ? "Insert Rows" : "Insert Columns",
      before: this.selection,
      after: this.selection,
    })
  }

  deleteLines(axis: "rows" | "cols") {
    const spans = this.selectedLines(axis).reverse()
    const ops: Op[] = spans.map(([from, to]) => ({
      kind: axis === "rows" ? "deleteRows" : "deleteCols",
      sheet: this.sheetIndex,
      at: from,
      count: to - from + 1,
    }))
    this.wb.transact(ops, {
      label: axis === "rows" ? "Delete Rows" : "Delete Columns",
      before: this.selection,
      after: this.selection,
    })
  }

  /** Insert/Delete cells with a shift (the Insert/Delete dialogs' first two choices). */
  shiftCells(kind: "insert" | "delete", shift: "down" | "right" | "up" | "left") {
    if (this.selection.ranges.length !== 1) return
    const range = this.selection.ranges[0]
    const op: Op =
      kind === "insert"
        ? { kind: "insertCells", sheet: this.sheetIndex, range, shift: shift as "down" | "right" }
        : { kind: "deleteCells", sheet: this.sheetIndex, range, shift: shift as "up" | "left" }
    this.wb.transact([op], {
      label: kind === "insert" ? "Insert Cells" : "Delete Cells",
      before: this.selection,
      after: this.selection,
    })
  }

  /** Ctrl+- / Ctrl++ : whole rows or columns act at once; a cell range asks how to shift (the shell's dialog). */
  insertOrDelete(kind: "insert" | "delete") {
    const range = this.selection.ranges[this.selection.ranges.length - 1]
    if (isWholeRows(range) && !isWholeColumns(range)) {
      kind === "insert" ? this.insertLines("rows") : this.deleteLines("rows")
    } else if (isWholeColumns(range) && !isWholeRows(range)) {
      kind === "insert" ? this.insertLines("cols") : this.deleteLines("cols")
    } else if (isWholeColumns(range) && isWholeRows(range)) {
      if (kind === "delete") this.clearContents()
    } else {
      this.shell.insertDelete({ kind })
    }
  }

  setLineSize(axis: "rows" | "cols", indexes: number[], size: number | null) {
    const target = axis === "rows" ? this.sheet.rows : this.sheet.cols
    const entries = indexes.map((index) => ({ index, size, hidden: size === 0 ? false : target.isHidden(index) }))
    if (size === 0) {
      // Excel: a width/height of 0 hides the line.
      this.setHidden(axis, indexes, true)
      return
    }
    this.wb.transact([{ kind: "axis", sheet: this.sheetIndex, axis, entries }], {
      label: axis === "rows" ? "Row Height" : "Column Width",
      before: this.selection,
      after: this.selection,
    })
  }

  setHidden(axis: "rows" | "cols", indexes: number[], hidden: boolean) {
    const target = axis === "rows" ? this.sheet.rows : this.sheet.cols
    const entries = indexes.map((index) => ({
      index,
      size: target.hasCustomSize(index) ? target.storedSize(index) : null,
      hidden,
    }))
    this.wb.transact([{ kind: "axis", sheet: this.sheetIndex, axis, entries }], {
      label: hidden ? "Hide" : "Unhide",
      before: this.selection,
      after: this.selection,
    })
  }

  /** Every index in the selection along `axis` (capped: whole-sheet selections act on used lines only). */
  selectedIndexes(axis: "rows" | "cols"): number[] {
    const out: number[] = []
    for (const [from, to] of this.selectedLines(axis)) {
      const span = to - from + 1
      if (span > 20_000) {
        // Whole sheet: Excel acts on all; we act on the lines that differ from the default plus the used range.
        const used = this.sheet.cells.usedBounds()
        const end = used ? Math.min(to, axis === "rows" ? used.r1 : used.c1) : from
        for (let i = from; i <= end; i++) out.push(i)
        continue
      }
      for (let i = from; i <= to; i++) out.push(i)
    }
    return out
  }

  /** Ctrl+9 / Ctrl+0 hide, Ctrl+Shift+9 / Ctrl+Shift+0 unhide (the selection's rows or columns). */
  hideSelected(axis: "rows" | "cols", hidden: boolean) {
    this.setHidden(axis, this.selectedIndexes(axis), hidden)
  }

  // ── Scrolling ──────────────────────────────────────────────────────────────────────────────────────────────

  scrollTo(x: number, y: number) {
    const nx = Math.max(0, Math.min(x, this.sheet.cols.total - this.viewWidth))
    const ny = Math.max(0, Math.min(y, this.sheet.rows.total - this.viewHeight))
    if (nx !== this.scrollX || ny !== this.scrollY) {
      this.scrollX = nx
      this.scrollY = ny
      this.changed()
    }
  }

  /** Scrolls the least amount that shows the whole cell (or as much as fits). */
  scrollIntoView(address: CellAddress) {
    const rows = this.sheet.rows
    const cols = this.sheet.cols
    const top = rows.offset(address.row)
    const bottom = top + rows.size(address.row)
    const left = cols.offset(address.col)
    const right = left + cols.size(address.col)
    let { scrollX, scrollY } = this
    if (top < scrollY) scrollY = top
    else if (bottom > scrollY + this.viewHeight) scrollY = Math.min(top, bottom - this.viewHeight)
    if (left < scrollX) scrollX = left
    else if (right > scrollX + this.viewWidth) scrollX = Math.min(left, right - this.viewWidth)
    this.scrollX = Math.max(0, scrollX)
    this.scrollY = Math.max(0, scrollY)
  }

  /** Ctrl+Backspace: bring the active cell into view. */
  revealActive() {
    this.scrollIntoView(this.selection.active)
    this.changed()
  }

  /** Go To (F5 / Ctrl+G) and the Name Box: select a reference; false when it does not parse. */
  goTo(range: Range): void {
    const active = { row: range.r0, col: range.c0 }
    const selection = sameRange(range, rangeOf(active)) ? selectCell(active) : selectRange(range, active)
    this.tabStartCol = null
    this.setSelection(selection, active)
  }

  // ── Keyboard ───────────────────────────────────────────────────────────────────────────────────────────────

  /**
   * Keys pressed while the grid has focus and no entry is open. Returns true when handled (the view then
   * prevents the default). Printable characters are not handled here: the hidden editor receives them and starts
   * an entry.
   */
  keyReady(input: KeyInput): boolean {
    const { key, primary, shift, alt } = input
    const arrows: Record<string, Direction> = {
      ArrowUp: "up",
      ArrowDown: "down",
      ArrowLeft: "left",
      ArrowRight: "right",
    }
    if (arrows[key] && !alt) {
      this.arrow(arrows[key], primary, shift)
      return true
    }
    switch (key) {
      case "Enter":
        if (alt || primary) return false
        this.enterMove(shift ? "up" : "down")
        return true
      case "Tab":
        if (primary || alt) return false
        this.enterMove(shift ? "left" : "right")
        return true
      case "Home":
        this.home(primary, shift)
        return true
      case "End":
        if (primary) this.ctrlEnd(shift)
        else this.toggleEndMode()
        return true
      case "PageDown":
      case "PageUp":
        if (primary) return false // Ctrl+PgUp/PgDn switch sheets (Phase 3).
        this.page(key === "PageDown" ? 1 : -1, alt, shift)
        return true
      case "F2":
        this.beginEdit("edit")
        return true
      case "F5":
        this.shell.goTo()
        return true
      case "Delete":
        this.clearContents()
        return true
      case "Backspace":
        if (primary) {
          this.revealActive()
          return true
        }
        this.beginEdit("enter", "")
        return true
      case "Escape":
        this.endMode = false
        this.clearClip()
        this.changed()
        return true
    }
    if (primary && !alt) {
      const lower = key.toLowerCase()
      if (lower === "z" && !shift) return (this.undo(), true)
      if (lower === "y" || (lower === "z" && shift)) return (this.redo(), true)
      if (lower === "a" && !shift) return (this.selectAllOrRegion(), true)
      if (key === " " && shift) return (this.selectEverything(), true)
      if (key === " ") return (this.selectColumns(), true)
      if (lower === "g") return (this.shell.goTo(), true)
      if (lower === "d") return (this.fillFromEdge("down"), true)
      if (lower === "r") return (this.fillFromEdge("right"), true)
      if (key === "-" || key === "_") return (this.insertOrDelete("delete"), true)
      if (key === "+" || (key === "=" && shift)) return (this.insertOrDelete("insert"), true)
      if (key === "*" || (key === "8" && shift)) return (this.selectCurrentRegion(), true)
      if (key === "9" || key === "(") return (this.hideSelected("rows", !shift && key === "9"), true)
      if (key === "0" || key === ")") return (this.hideSelected("cols", !shift && key === "0"), true)
    }
    if (!primary && !alt && shift && key === " ") {
      this.selectRowsOfSelection()
      return true
    }
    return false
  }

  /**
   * Keys while an entry is open. Returns "commit-move" style outcomes for the view: true when handled. Arrows
   * commit in enter mode; in edit mode they move the caret (not handled here).
   */
  keyEditing(input: KeyInput): boolean {
    const { key, primary, shift, alt } = input
    if (!this.edit) return false
    if (key === "Escape") {
      this.cancelEdit()
      return true
    }
    if (key === "Enter") {
      if (alt) return false // Alt+Enter: a line break inside the cell; the editor inserts it.
      this.commitEdit(primary)
      if (!primary) this.enterMove(shift ? "up" : "down")
      return true
    }
    if (key === "Tab") {
      this.commitEdit()
      this.enterMove(shift ? "left" : "right")
      return true
    }
    if (key === "F2") {
      this.toggleEditMode()
      return true
    }
    const arrows: Record<string, Direction> = {
      ArrowUp: "up",
      ArrowDown: "down",
      ArrowLeft: "left",
      ArrowRight: "right",
    }
    if (arrows[key] && this.edit.mode === "enter" && this.edit.surface === "cell" && !alt) {
      this.commitEdit()
      this.arrow(arrows[key], primary, shift)
      return true
    }
    return false
  }

  // ── Status bar ─────────────────────────────────────────────────────────────────────────────────────────────

  get referenceText() {
    return rangeName(selectionBounds(this.selection))
  }
}

function truncate(text: string) {
  return text.length > 24 ? `${text.slice(0, 24)}…` : text
}

/** Shift+arrows on a whole-row (or column) selection keep it whole. */
function keepWholeLines(selection: Selection, focus: CellAddress): CellAddress {
  const range = selection.ranges[selection.ranges.length - 1]
  if (isWholeRows(range) && !isWholeColumns(range)) return { row: focus.row, col: MAX_COLS - 1 }
  if (isWholeColumns(range) && !isWholeRows(range)) return { row: MAX_ROWS - 1, col: focus.col }
  return focus
}

/** Copying whole columns or the whole sheet copies the used part (as the clipboard formats do in Excel). */
function clampToUsed(sheet: Sheet, range: Range): Range {
  const used = sheet.cells.usedBounds()
  if (!isWholeColumns(range) && !isWholeRows(range)) return range
  if (!used) return { ...range, r1: range.r0, c1: range.c0 }
  return {
    r0: range.r0,
    c0: range.c0,
    r1: isWholeColumns(range) ? Math.max(range.r0, used.r1) : range.r1,
    c1: isWholeRows(range) ? Math.max(range.c0, used.c1) : range.c1,
  }
}

function normalizeNewlines(text: string) {
  return text.replace(/\r\n/g, "\n").replace(/\n$/, "")
}

/** Rows of a pasted HTML table (other apps' copies), read with the browser's parser. */
export function htmlTableRows(html: string): string[][] {
  if (typeof DOMParser === "undefined") return []
  const doc = new DOMParser().parseFromString(html, "text/html")
  const rows = [...doc.querySelectorAll("tr")]
  return rows.map((row) =>
    [...row.querySelectorAll("td, th")].map((cell) => (cell as HTMLElement).innerText ?? cell.textContent ?? ""),
  )
}
