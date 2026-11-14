import {
  CLEAR_ALL,
  EXPLICIT_DEFAULT_ID,
  FONT_SIZES,
  MAX_COLS,
  MAX_ROWS,
  borderOps,
  changeDecimals,
  clearContentsOp,
  formatValue,
  rangeCols,
  rangeRows,
  stepFontSize,
  type Border,
  type BorderPreset,
  type CellStyle,
  type HAlign,
  type Op,
  type Range,
  type StylePatch,
  type VAlign,
} from "@cells/core"
import type { SheetController } from "./controller"

/** Formats that Home-tab buttons and Ctrl+Shift shortcuts apply (as Excel's en-US built-ins, read from Excel). */
export const NUMBER_FORMATS = {
  general: "General",
  number: "#,##0.00",
  currency: "$#,##0.00_);($#,##0.00)",
  accounting: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  comma: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  percent: "0%",
  scientific: "0.00E+00",
  shortDate: "m/d/yyyy",
  longDate: "[$-x-sysdate]dddd, mmmm dd, yyyy",
  dayMonthYear: "d-mmm-yy",
  time: "h:mm AM/PM",
  longTime: "[$-x-systime]h:mm:ss AM/PM",
  fraction: "# ?/?",
  text: "@",
} as const

/** Indent levels stop here (Excel's limit). */
export const MAX_INDENT = 250

/**
 * Formatting commands on the selection, as Excel's Home tab and shortcuts do them. Toggles (Bold, Italic, Wrap,
 * alignment) and steps (Grow Font, Indent) read the active cell and apply the result to the whole selection, which is
 * what Excel does (checked with its own ribbon commands).
 */
export class FormatCommands {
  /** Format Painter: the styles to paint and whether it stays on (double-click) after painting. */
  painter: { range: Range; sheet: number; sticky: boolean } | null = null

  constructor(private readonly c: SheetController) {}

  /** The active cell's format: what the ribbon shows as pressed. */
  get active(): CellStyle {
    const { row, col } = this.c.selection.active
    return this.c.wb.styles.get(this.c.sheet.styleAt(row, col))
  }

  apply(patch: StylePatch, label = "Format") {
    this.run([{ kind: "format", sheet: this.c.sheetIndex, ranges: [...this.c.selection.ranges], patch }], label)
  }

  private run(ops: Op[], label: string) {
    if (!ops.length || this.c.edit) return
    const sel = this.c.selection
    this.c.transact(ops, { label, before: sel, after: sel })
  }

  // ── Font ───────────────────────────────────────────────────────────────────────────────────────────────────

  toggle(field: "bold" | "italic" | "strike") {
    this.apply({ [field]: this.active[field] ? null : true }, field === "strike" ? "Strikethrough" : cap(field))
  }

  /** Underline and Double Underline: pressing the one that is on turns it off. */
  underline(kind: "single" | "double" = "single") {
    this.apply({ underline: this.active.underline === kind ? null : kind }, "Underline")
  }

  fontName(name: string) {
    this.apply({ fontName: name === DEFAULT_FONT ? null : name }, "Font")
  }

  /** A size in points; Excel accepts 1 to 409 in half points. */
  fontSize(size: number) {
    if (!Number.isFinite(size) || size < 1 || size > 409) return false
    const pt = Math.round(size * 2) / 2
    this.apply({ fontSize: pt === 11 ? null : pt }, "Font Size")
    return true
  }

  /** Grow Font / Shrink Font: the active cell's size one step along Excel's list, for every selected cell. */
  stepFont(dir: 1 | -1) {
    this.fontSize(stepFontSize(this.active.fontSize ?? 11, dir))
  }

  fontColor(color: string | null) {
    this.apply({ fontColor: color }, "Font Color")
  }

  fill(color: string | null) {
    this.apply({ fill: color }, "Fill Color")
  }

  // ── Alignment ──────────────────────────────────────────────────────────────────────────────────────────────

  /** Left, Center, Right: pressing the one that is on goes back to General alignment. */
  hAlign(align: HAlign) {
    this.apply({ hAlign: this.active.hAlign === align ? null : align, indent: align === "center" ? null : undefined })
  }

  /** Top, Middle, Bottom: Bottom is the default, so pressing it (or the one that is on) clears it. */
  vAlign(align: VAlign) {
    this.apply({ vAlign: align === "bottom" || this.active.vAlign === align ? null : align }, "Alignment")
  }

  wrap() {
    this.apply({ wrap: this.active.wrap ? null : true }, "Wrap Text")
  }

  /** Orientation menu: angles in Excel's encoding (45, 135 clockwise, 255 stacked, 90 up, 180 down). */
  rotation(value: number) {
    this.apply({ rotation: this.active.rotation === value ? null : value }, "Orientation")
  }

  /** Increase / Decrease Indent. Indenting turns the alignment to Left (Excel does, even from Center). */
  indent(dir: 1 | -1) {
    const level = Math.max(0, Math.min(MAX_INDENT, (this.active.indent ?? 0) + dir))
    const align = this.active.hAlign
    const keeps = align === "left" || align === "right" || align === "distributed"
    if (dir === 1) this.apply({ indent: level, hAlign: keeps ? undefined : "left" }, "Indent")
    else this.apply({ indent: level || null }, "Indent")
  }

  // ── Merge ──────────────────────────────────────────────────────────────────────────────────────────────────

  /** The merge areas a command makes from the selection (Merge Across: one per row). */
  private mergeAreas(across: boolean): Range[] {
    const out: Range[] = []
    for (const r of this.c.selection.ranges) {
      if (rangeRows(r) * rangeCols(r) < 2 || (across && r.c0 === r.c1)) continue
      if (across) for (let row = r.r0; row <= Math.min(r.r1, r.r0 + 9_999); row++) out.push({ ...r, r0: row, r1: row })
      else out.push(r)
    }
    return out
  }

  /** Whether merging would throw values away (more than the top-left value in an area). */
  mergeLosesData(across = false): boolean {
    return this.mergeAreas(across).some((area) => {
      let seen = 0
      this.c.sheet.cells.forEach(area, (_r, _c, data) => {
        if (data.value !== undefined && data.value !== "") seen++
      })
      return seen > 1
    })
  }

  /**
   * Merge & Center, Merge Across, Merge Cells. Only the top-left value stays (`confirm` asks first, as Excel's
   * "Merging cells only keeps the upper-left value" warning does).
   */
  merge(kind: "center" | "across" | "cells") {
    const areas = this.mergeAreas(kind === "across")
    if (!areas.length) return
    // Merge & Center on a merged selection unmerges (the button is a toggle).
    if (kind === "center" && this.isMerged()) return this.unmerge()
    const go = () => {
      const sheet = this.c.sheetIndex
      const ops: Op[] = []
      for (const area of areas) {
        const rest = [
          { ...area, r0: area.r0, r1: area.r0, c0: area.c0 + 1 },
          { ...area, r0: area.r0 + 1 },
        ].filter((r) => r.r0 <= r.r1 && r.c0 <= r.c1)
        const clear = clearContentsOp(sheet, this.c.sheet, rest)
        if (clear) ops.push(clear)
      }
      ops.push({ kind: "merge", sheet, ranges: areas })
      if (kind === "center") ops.push({ kind: "format", sheet, ranges: areas, patch: { hAlign: "center" } })
      this.run(ops, kind === "center" ? "Merge & Center" : "Merge Cells")
      this.c.setSelection(this.c.selection, null)
    }
    if (this.mergeLosesData(kind === "across"))
      this.c.shell.confirm("Merging cells only keeps the upper-left value and discards other values.", go)
    else go()
  }

  unmerge() {
    this.run([{ kind: "unmerge", sheet: this.c.sheetIndex, ranges: [...this.c.selection.ranges] }], "Unmerge Cells")
  }

  /** Whether the selection holds a merge (Merge & Center shows as pressed). */
  isMerged(): boolean {
    return this.c.selection.ranges.some((r) => this.c.sheet.merges.overlapping(r).length > 0)
  }

  // ── Borders ────────────────────────────────────────────────────────────────────────────────────────────────

  /** The pen chosen under Line Color and Line Style; the Borders button remembers the last preset. */
  pen: Border = { style: "thin" }
  lastBorder: BorderPreset = "bottom"

  border(preset: BorderPreset) {
    this.lastBorder = preset
    const ops = this.c.selection.ranges.flatMap((r) => borderOps(this.c.sheetIndex, r, preset, this.pen))
    this.run(ops, "Borders")
  }

  // ── Number ─────────────────────────────────────────────────────────────────────────────────────────────────

  numberFormat(code: string, cellStyle?: string) {
    this.apply({ numberFormat: code === "General" ? null : code, cellStyle: cellStyle ?? undefined }, "Number Format")
  }

  /** Increase / Decrease Decimal, from the active cell's format (or, on General, from what it shows). */
  decimals(delta: 1 | -1) {
    const { row, col } = this.c.selection.active
    const value = this.c.sheet.cells.value(row, col)
    const shown = typeof value === "number" ? (this.c.shownText?.(row, col) ?? formatValue(value) ?? "") : ""
    const next = changeDecimals(this.active.numberFormat, delta, shown)
    if (next !== null) this.apply({ numberFormat: next }, delta > 0 ? "Increase Decimal" : "Decrease Decimal")
  }

  // ── Clear ──────────────────────────────────────────────────────────────────────────────────────────────────

  clearFormats() {
    const sheet = this.c.sheetIndex
    const ops: Op[] = [{ kind: "format", sheet, ranges: [...this.c.selection.ranges], patch: CLEAR_ALL }]
    if (this.isMerged()) ops.push({ kind: "unmerge", sheet, ranges: [...this.c.selection.ranges] })
    this.run(ops, "Clear Formats")
  }

  clearAll() {
    const sheet = this.c.sheetIndex
    const ops: Op[] = []
    const clear = clearContentsOp(sheet, this.c.sheet, this.c.selection.ranges)
    if (clear) ops.push(clear)
    ops.push({ kind: "format", sheet, ranges: [...this.c.selection.ranges], patch: CLEAR_ALL })
    if (this.isMerged()) ops.push({ kind: "unmerge", sheet, ranges: [...this.c.selection.ranges] })
    const notes = this.c.annotations.removalOp()
    if (notes) ops.push(notes)
    this.run(ops, "Clear")
  }

  // ── Format Painter ─────────────────────────────────────────────────────────────────────────────────────────

  /** Picks up the selection's formats; the next selection gets them. Pressing it again turns it off. */
  armPainter(sticky = false) {
    if (this.painter && !sticky) {
      this.painter = null
    } else {
      this.painter = { range: this.c.selection.ranges[0], sheet: this.c.sheetIndex, sticky }
    }
    this.c.changed()
  }

  cancelPainter() {
    if (!this.painter) return
    this.painter = null
    this.c.changed()
  }

  /**
   * Paints the picked-up formats onto `target`, repeating the source's pattern when the target is bigger (Excel
   * tiles it the same way). Whole rows and columns are capped to the used area.
   */
  paint(target: Range) {
    const p = this.painter
    if (!p) return
    if (!p.sticky) this.painter = null
    const src = this.c.wb.sheets[p.sheet]
    const sheet = this.c.sheet
    const used = sheet.cells.usedBounds()
    const t = {
      ...target,
      r1: target.r1 === MAX_ROWS - 1 && used ? Math.max(target.r0, used.r1) : target.r1,
      c1: target.c1 === MAX_COLS - 1 && used ? Math.max(target.c0, used.c1) : target.c1,
    }
    const h = rangeRows(p.range)
    const w = rangeCols(p.range)
    const rows = Math.max(rangeRows(t), h)
    const cols = Math.max(rangeCols(t), w)
    const cells: number[] = []
    for (let dr = 0; dr < rows; dr++) {
      for (let dc = 0; dc < cols; dc++) {
        const row = t.r0 + dr
        const col = t.c0 + dc
        if (row >= MAX_ROWS || col >= MAX_COLS) continue
        const id = src.styleAt(p.range.r0 + (dr % h), p.range.c0 + (dc % w))
        const line = sheet.rowStyles.get(row) || sheet.colStyles.get(col)
        cells.push(row, col, id === 0 && line ? EXPLICIT_DEFAULT_ID : id)
      }
    }
    // The undo record is built by the op itself; "styles" puts exact ids.
    this.run([{ kind: "styles", sheet: this.c.sheetIndex, record: { cells, rows: [], cols: [] } }], "Format Painter")
    this.c.changed()
  }
}

/** The body font a new workbook uses (Excel's default). */
export const DEFAULT_FONT = "Aptos Narrow"

export { FONT_SIZES }

const cap = (s: string) => s[0].toUpperCase() + s.slice(1)
