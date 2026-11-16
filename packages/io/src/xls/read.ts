import {
  errorValue,
  Workbook,
  ERROR_CODES,
  type Border,
  type BorderStyle,
  type CellData,
  type CellStyle,
  type CellValue,
  type DefinedName,
  type ErrorCode,
  type HAlign,
  type Sheet,
  type Underline,
  type VAlign,
} from "@cells/core"
import { BUILTIN_CELL_STYLES, BUILTIN_NUMBER_FORMATS, INDEXED_COLORS } from "../xlsx/constants"
import { formulaFromFile } from "../xlsx/formula"
import { readPrintArea, readPrintTitles } from "../xlsx/page-setup"
import { DEFAULT_VIEW, needsApostrophe, type SheetView } from "../xlsx/sheet-read"
import { colWidthToPx, pointsToPx } from "../xlsx/units"
import { Cursor, R, readRecords, type BiffRecord } from "./biff"
import { CfbError, CompoundFile } from "./cfb"
import { decodeFormula, errorText, FormulaSkip, type FormulaContext, type Xti } from "./formula"

/**
 * Opens an Excel 97-2003 workbook (.xls, BIFF8 in a compound file): its sheets, values, formulas, number formats,
 * fonts, fills, borders, alignment, column widths, row heights, hidden rows and columns, merged cells, frozen panes,
 * defined names and print areas. It opens as a new workbook, as Excel's "convert" does: Cells saves .xlsx, never .xls.
 * Notes too. Left out for now, each with a warning: pictures and charts, conditional formats, data validation, AutoFilter,
 * hyperlinks and sheet protection.
 */

export class XlsError extends Error {
  constructor(
    message: string,
    /** "encrypted": protected with a password to open; "old": Excel 5.0/95 or older; "damaged": unreadable. */
    readonly kind: "encrypted" | "old" | "damaged",
  ) {
    super(message)
    this.name = "XlsError"
  }
}

export type XlsRead = {
  workbook: Workbook
  views: SheetView[]
  activeSheet: number
  warnings: string[]
}

type Font = {
  name: string
  size: number
  bold: boolean
  italic: boolean
  strike: boolean
  underline: Underline | null
  vertAlign: "superscript" | "subscript" | null
  color: number
}

const H_ALIGN: (HAlign | undefined)[] = [
  undefined,
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerContinuous",
  "distributed",
]
const V_ALIGN: (VAlign | undefined)[] = ["top", "center", undefined, "justify", "distributed"]
const BORDER_STYLES: (BorderStyle | null)[] = [
  null,
  "thin",
  "medium",
  "dashed",
  "dotted",
  "thick",
  "double",
  "hair",
  "mediumDashed",
  "dashDot",
  "mediumDashDot",
  "dashDotDot",
  "mediumDashDotDot",
  "slantDashDot",
]
const UNDERLINES: Record<number, Underline> = {
  0x01: "single",
  0x02: "double",
  0x21: "singleAccounting",
  0x22: "doubleAccounting",
}

/** An RK number ([MS-XLS] 2.5.217): 30 bits of a double's top or an integer, maybe divided by 100. */
export function rkValue(rk: number): number {
  let n: number
  if (rk & 0x02) n = rk >> 2
  else {
    const buf = new DataView(new ArrayBuffer(8))
    buf.setUint32(4, rk & 0xfffffffc, true)
    n = buf.getFloat64(0, true)
  }
  return rk & 0x01 ? n / 100 : n
}

function cellError(code: number): CellValue {
  const text = errorText(code)
  return (ERROR_CODES as readonly string[]).includes(text) ? errorValue(text as ErrorCode) : errorValue("#N/A")
}

/** The workbook-wide part of the stream: what the sheets refer to. */
class Globals {
  sheets: { name: string; offset: number; state: number; type: number }[] = []
  strings: string[] = []
  fonts: Font[] = []
  formats = new Map<number, string>()
  xfs: Cursor[] = []
  palette = INDEXED_COLORS.slice()
  date1904 = false
  activeTab = 0
  xti: Xti[] = []
  books: { self: boolean; addIn: boolean; names: string[] }[] = []
  /** Style XF index → its cell style's name (STYLE records), "Normal" left out. */
  styleNames = new Map<number, string>()
  names: {
    name: string
    sheet: number
    hidden: boolean
    builtin: string | null
    record: BiffRecord
    at: number
    cce: number
  }[] = []
}

const BUILTIN_NAMES: Record<number, string> = {
  0x00: "Consolidate_Area",
  0x01: "Auto_Open",
  0x02: "Auto_Close",
  0x03: "Extract",
  0x04: "Database",
  0x05: "Criteria",
  0x06: "Print_Area",
  0x07: "Print_Titles",
  0x08: "Recorder",
  0x09: "Data_Form",
  0x0a: "Auto_Activate",
  0x0b: "Auto_Deactivate",
  0x0c: "Sheet_Title",
  0x0d: "_FilterDatabase",
}

const BUILTIN_STYLE_NAMES = new Map(Object.entries(BUILTIN_CELL_STYLES).map(([name, id]) => [id, name]))

function readGlobals(records: BiffRecord[], end: number, g: Globals, warnings: string[]) {
  for (let i = 1; i < end; i++) {
    const rec = records[i]
    const c = new Cursor(rec.data, 0, rec.breaks)
    switch (rec.type) {
      case R.FILEPASS:
        throw new XlsError("The workbook is protected with a password to open.", "encrypted")
      case R.BOUNDSHEET: {
        const offset = c.u32()
        const state = c.u8() & 0x03
        const type = c.u8()
        g.sheets.push({ name: c.string(1), offset, state, type })
        break
      }
      case R.SST: {
        c.skip(8)
        while (c.left > 0) g.strings.push(c.string())
        break
      }
      case R.FONT: {
        const height = c.u16()
        const flags = c.u16()
        const color = c.u16()
        const weight = c.u16()
        const script = c.u16()
        const underline = c.u8()
        c.skip(3)
        g.fonts.push({
          name: c.string(1),
          size: height / 20,
          bold: weight >= 700,
          italic: (flags & 0x02) !== 0,
          strike: (flags & 0x08) !== 0,
          underline: UNDERLINES[underline] ?? null,
          vertAlign: script === 1 ? "superscript" : script === 2 ? "subscript" : null,
          color,
        })
        break
      }
      case R.FORMAT: {
        const id = c.u16()
        g.formats.set(id, c.string())
        break
      }
      case R.XF:
        g.xfs.push(c)
        break
      case R.PALETTE: {
        const count = c.u16()
        for (let k = 0; k < count; k++) {
          const [r, gr, b] = [c.u8(), c.u8(), c.u8()]
          c.skip(1)
          g.palette[8 + k] = [r, gr, b]
            .map((v) => v.toString(16).padStart(2, "0"))
            .join("")
            .toUpperCase()
        }
        break
      }
      case 0x0293: {
        // STYLE: a built-in style by number, or a style of the user's own by name.
        const ixfe = c.u16()
        // LibreOffice writes Excel's built-in styles as its own, named "Excel Built-in Good".
        const name = ixfe & 0x8000 ? BUILTIN_STYLE_NAMES.get(c.u8()) : c.string().replace(/^Excel Built-in /, "")
        if (name && name !== "Normal") g.styleNames.set(ixfe & 0x0fff, name)
        break
      }
      case R.DATEMODE:
        g.date1904 = c.u16() === 1
        break
      case R.WINDOW1:
        c.skip(10)
        g.activeTab = c.u16()
        break
      case R.SUPBOOK: {
        const count = c.u16()
        const marker = c.u16()
        // 0x0401: this workbook; 0x3A01: add-in functions; else another workbook (its path follows).
        g.books.push({ self: marker === 0x0401, addIn: marker === 0x3a01 && count === 1, names: [] })
        break
      }
      case R.EXTERNNAME: {
        const book = g.books[g.books.length - 1]
        if (!book) break
        c.skip(6)
        book.names.push(c.string(1))
        break
      }
      case R.EXTERNSHEET: {
        const count = c.u16()
        for (let k = 0; k < count; k++) {
          const supbook = c.u16()
          const book = g.books[supbook]
          const first = c.i16()
          const last = c.i16()
          g.xti.push({
            sheets: first < 0 || last < 0 ? null : [first, last],
            book: book?.self || book?.addIn ? null : "other",
            supbook,
          })
        }
        break
      }
      case R.NAME: {
        const flags = c.u16()
        c.skip(1)
        const length = c.u8()
        const cce = c.u16()
        c.skip(2)
        const sheet = c.u16()
        c.skip(4)
        const high = (c.u8() & 1) === 1
        const builtinCode = flags & 0x20 ? (high ? c.u16() : c.u8()) : null
        const name =
          builtinCode === null ? c.chars(length, high) : (BUILTIN_NAMES[builtinCode] ?? `Builtin${builtinCode}`)
        const builtin = builtinCode === null ? null : name
        g.names.push({ name, sheet, hidden: (flags & 0x01) !== 0, builtin, record: rec, at: c.pos, cce })
        break
      }
    }
  }
  if (g.date1904)
    warnings.push("The workbook uses the 1904 date system; Cells shows its dates from 1900, 4 years and 1 day earlier.")
}

/** An XF record's look as a Cells style (fonts relative to the workbook's first font, as the .xlsx reader does). */
function xfStyle(g: Globals, xf: Cursor): CellStyle {
  const c = new Cursor(xf.data)
  const fontIndex = c.u16()
  const formatId = c.u16()
  const protect = c.u16()
  const align = c.u8()
  const rotation = c.u8()
  const indentShrink = c.u8()
  c.skip(1)
  const border1 = c.u32()
  const border2 = c.u32()
  const fill = c.u16()
  const style: CellStyle = {}
  const color = (index: number) => {
    const hex = g.palette[index]
    return index < 64 && hex ? "#" + hex.toUpperCase() : undefined
  }

  const format = g.formats.get(formatId) ?? BUILTIN_NUMBER_FORMATS[formatId]
  if (format && format !== "General") style.numberFormat = format

  // Font 4 doesn't exist: indexes past 3 are one higher than the font's place in the list.
  const font = g.fonts[fontIndex >= 4 ? fontIndex - 1 : fontIndex]
  const base = g.fonts[0]
  if (font) {
    if (base && font.name !== base.name) style.fontName = font.name
    if (base && font.size !== base.size) style.fontSize = font.size
    if (font.bold) style.bold = true
    if (font.italic) style.italic = true
    if (font.strike) style.strike = true
    if (font.underline) style.underline = font.underline
    if (font.vertAlign) style.vertAlign = font.vertAlign
    const fc = color(font.color)
    if (fc && fc !== "#000000") style.fontColor = fc
  }

  const h = H_ALIGN[align & 0x07]
  if (h) style.hAlign = h
  if (align & 0x08) style.wrap = true
  const v = V_ALIGN[(align >> 4) & 0x07]
  if (v) style.vAlign = v
  if (rotation) style.rotation = rotation
  if (indentShrink & 0x0f) style.indent = indentShrink & 0x0f
  if (indentShrink & 0x10) style.shrink = true

  const side = (dg: number, icv: number): Border | undefined => {
    const s = BORDER_STYLES[dg]
    if (!s) return undefined
    const bc = color(icv)
    return bc && bc !== "#000000" ? { style: s, color: bc } : { style: s }
  }
  const left = side(border1 & 0x0f, (border1 >>> 16) & 0x7f)
  const right = side((border1 >>> 4) & 0x0f, (border1 >>> 23) & 0x7f)
  const top = side((border1 >>> 8) & 0x0f, border2 & 0x7f)
  const bottom = side((border1 >>> 12) & 0x0f, (border2 >>> 7) & 0x7f)
  if (left) style.borderLeft = left
  if (right) style.borderRight = right
  if (top) style.borderTop = top
  if (bottom) style.borderBottom = bottom
  const diag = (border1 >>> 30) & 0x03
  const diagonal = side((border2 >>> 21) & 0x0f, (border2 >>> 14) & 0x7f)
  if (diagonal && diag) {
    style.borderDiagonal = diagonal
    if (diag & 0x01) style.diagonalDown = true
    if (diag & 0x02) style.diagonalUp = true
  }

  // A fill pattern: Cells draws solid fills; other patterns show as their foreground colour.
  const pattern = (border2 >>> 26) & 0x3f
  if (pattern) {
    const fg = color(fill & 0x7f) ?? color((fill >> 7) & 0x7f)
    if (fg) style.fill = fg
  }

  const named = g.styleNames.get(protect >> 4)
  if (named && !(protect & 0x04)) style.cellStyle = named
  if (!(protect & 0x01)) style.locked = false
  if (protect & 0x02) style.hiddenFormula = true
  return style
}

/** The cells, sizes, merges and view of one worksheet's records. */
function readSheet(
  records: BiffRecord[],
  from: number,
  to: number,
  sheet: Sheet,
  ctx: Omit<FormulaContext, "row" | "col">,
  styleId: (xf: number) => number,
  g: Globals,
  warnings: string[],
): SheetView {
  const view: SheetView = { ...DEFAULT_VIEW }
  const shared = new Map<string, { rec: BiffRecord; r0: number; r1: number; c0: number; c1: number }>()
  const arrays = new Set<string>()
  let skipped = 0
  const left = new Set<string>()
  let defaultHeight = 255
  for (let i = from; i < to; i++)
    if (records[i].type === R.DEFAULTROWHEIGHT) defaultHeight = new Cursor(records[i].data, 2).u16()
  // Notes: each drawing object (OBJ) is followed by its text box (TXO); NOTE records name the cell, at the end.
  const texts = new Map<number, string>()
  let lastObj = -1

  // Shared formulas' and array formulas' records follow their first cell: collect them first.
  for (let i = from; i < to; i++) {
    const rec = records[i]
    if (rec.type === R.SHRFMLA || rec.type === R.ARRAY) {
      const c = new Cursor(rec.data)
      const r0 = c.u16()
      const r1 = c.u16()
      const c0 = c.u8()
      const c1 = c.u8()
      if (rec.type === R.SHRFMLA) shared.set(`${r0},${c0}`, { rec, r0, r1, c0, c1 })
      else arrays.add(`${r0},${c0}`)
    }
  }
  if (arrays.size) warnings.push(`${sheet.name} has array formulas; Cells keeps their values.`)

  const put = (row: number, col: number, xf: number, value: CellValue | undefined, input?: string) => {
    const data: CellData = {}
    const style = styleId(xf)
    if (style) data.style = style
    if (value !== undefined) data.value = value
    if (input) data.input = input
    else if (typeof value === "string" && value !== "" && needsApostrophe(value)) data.input = "'" + value
    if (data.value !== undefined || data.style) sheet.cells.set(row, col, data)
  }

  const formulaText = (row: number, col: number, c: Cursor, cce: number): string | undefined => {
    try {
      return formulaFromFile(decodeFormula(c, cce, { ...ctx, row, col }))
    } catch (error) {
      if (!(error instanceof FormulaSkip)) throw error
      skipped++
      return undefined
    }
  }

  for (let i = from; i < to; i++) {
    const rec = records[i]
    const c = new Cursor(rec.data, 0, rec.breaks)
    switch (rec.type) {
      case R.BOF: {
        // An embedded chart's substream: skip to its EOF.
        let depth = 1
        while (depth && ++i < to) {
          if (records[i].type === R.BOF) depth++
          else if (records[i].type === R.EOF) depth--
        }
        left.add("charts")
        break
      }
      case R.NUMBER: {
        const row = c.u16()
        const col = c.u16()
        const xf = c.u16()
        put(row, col, xf, c.f64())
        break
      }
      case R.RK: {
        const row = c.u16()
        const col = c.u16()
        const xf = c.u16()
        put(row, col, xf, rkValue(c.u32()))
        break
      }
      case R.MULRK: {
        const row = c.u16()
        let col = c.u16()
        while (c.left > 2) {
          const xf = c.u16()
          put(row, col++, xf, rkValue(c.u32()))
        }
        break
      }
      case R.LABELSST: {
        const row = c.u16()
        const col = c.u16()
        const xf = c.u16()
        put(row, col, xf, g.strings[c.u32()] ?? "")
        break
      }
      case R.LABEL: {
        const row = c.u16()
        const col = c.u16()
        const xf = c.u16()
        put(row, col, xf, c.string())
        break
      }
      case R.BOOLERR: {
        const row = c.u16()
        const col = c.u16()
        const xf = c.u16()
        const v = c.u8()
        put(row, col, xf, c.u8() ? cellError(v) : v !== 0)
        break
      }
      case R.BLANK: {
        const row = c.u16()
        const col = c.u16()
        put(row, col, c.u16(), undefined)
        break
      }
      case R.MULBLANK: {
        const row = c.u16()
        let col = c.u16()
        while (c.left > 2) put(row, col++, c.u16(), undefined)
        break
      }
      case R.FORMULA: {
        const row = c.u16()
        const col = c.u16()
        const xf = c.u16()
        const raw = c.bytes(8)
        c.skip(6)
        const cce = c.u16()
        let value: CellValue | undefined
        if (raw[6] === 0xff && raw[7] === 0xff) {
          if (raw[0] === 0x00) {
            // A text result: the STRING record after it.
            const next = records[i + 1]?.type === R.SHRFMLA || records[i + 1]?.type === R.ARRAY ? i + 2 : i + 1
            const str = records[next]
            value = str?.type === R.STRING ? new Cursor(str.data, 0, str.breaks).string() : ""
          } else if (raw[0] === 0x01) value = raw[2] !== 0
          else if (raw[0] === 0x02) value = cellError(raw[2])
          else value = ""
        } else value = new DataView(raw.buffer, raw.byteOffset, 8).getFloat64(0, true)
        let input: string | undefined
        if (cce >= 5 && c.data[c.pos] === 0x01) {
          // PtgExp: the shared formula whose first cell this names, decoded for this cell.
          const p = new Cursor(c.data, c.pos + 1)
          const key = `${p.u16()},${p.u16()}`
          const sh = shared.get(key)
          if (sh) {
            const s = new Cursor(sh.rec.data, 8, sh.rec.breaks)
            input = formulaText(row, col, s, s.u16())
          } else if (!arrays.has(key)) skipped++
        } else input = formulaText(row, col, c, cce)
        put(row, col, xf, value ?? 0, input ? input : undefined)
        break
      }
      case R.ROW: {
        const row = c.u16()
        c.skip(4)
        const height = c.u16() & 0x7fff
        c.skip(4)
        const flags = c.u32()
        // A height set by hand, or one Excel fitted to the row's text (kept as fitted, as the .xlsx reader does).
        const px = pointsToPx(height / 20)
        if (flags & 0x40) sheet.rows.setSize(row, px)
        else if (height !== defaultHeight && px !== sheet.rows.size(row)) sheet.rows.setSize(row, px, true)
        if (flags & 0x20) sheet.rows.setHidden(row, true)
        // Outline level (bits 0–2) and the collapsed mark (bit 4).
        if (flags & 0x07) sheet.rows.setLevel(row, flags & 0x07)
        if (flags & 0x10) sheet.rows.setCollapsed(row, true)
        break
      }
      case R.COLINFO: {
        const first = c.u16()
        const last = Math.min(c.u16(), 255)
        const coldx = c.u16()
        c.skip(2)
        const flags = c.u16()
        const px = colWidthToPx(coldx / 256)
        // About the default width is the default: LibreOffice writes its own default (8.6 characters) for columns
        // nobody widened, and its flags don't tell them apart.
        const width = Math.abs(px - sheet.cols.defaultSize) <= 4 ? null : px
        for (let col = first; col <= last; col++) {
          if (width !== null && width !== sheet.cols.size(col)) sheet.cols.setSize(col, width)
          if (flags & 0x01) sheet.cols.setHidden(col, true)
          if ((flags >> 8) & 0x07) sheet.cols.setLevel(col, (flags >> 8) & 0x07)
          if (flags & 0x1000) sheet.cols.setCollapsed(col, true)
        }
        break
      }
      case R.MERGECELLS: {
        const count = c.u16()
        for (let k = 0; k < count; k++) {
          const r0 = c.u16()
          const r1 = c.u16()
          const c0 = c.u16()
          const c1 = c.u16()
          if (r0 !== r1 || c0 !== c1) sheet.merges.add({ r0, r1, c0, c1 })
        }
        break
      }
      case R.WSBOOL: {
        // Where outline summaries sit (Data › Outline › Settings).
        const flags = c.u16()
        sheet.outline.summaryBelow = (flags & 0x40) !== 0
        sheet.outline.summaryRight = (flags & 0x80) !== 0
        break
      }
      case R.WINDOW2: {
        const flags = c.u16()
        if (!(flags & 0x80)) sheet.outline.showSymbols = false
        const top = c.u16()
        const leftCol = c.u16()
        view.showGridlines = (flags & 0x02) !== 0
        view.showHeadings = (flags & 0x04) !== 0
        view.selected = (flags & 0x200) !== 0
        view.topLeft = { row: top, col: leftCol }
        if (flags & 0x08) view.freeze = { row: 0, col: 0, top, left: leftCol }
        break
      }
      case R.PANE: {
        // Frozen panes: how many columns and rows are frozen, from the window's top-left cell.
        const x = c.u16()
        const y = c.u16()
        const { row, col } = view.topLeft
        if (view.freeze) view.freeze = x || y ? { row: row + y, col: col + x, top: row, left: col } : null
        break
      }
      case R.SCL: {
        const num = c.u16()
        const den = c.u16()
        if (num && den) view.zoom = num / den
        break
      }
      case 0x005d: {
        // OBJ: its first part (ftCmo) gives the kind of object and its id.
        c.skip(4)
        const kind = c.u16()
        lastObj = c.u16()
        // 0x19 a note's box, 0x14 a filter or list's drop-down button; the rest are pictures, shapes and charts.
        if (kind === 0x05) left.add("charts")
        else if (kind !== 0x19 && kind !== 0x14) left.add("pictures and shapes")
        break
      }
      case 0x01b6: {
        // TXO: the text is in the first CONTINUE piece, after a flags byte.
        c.skip(10)
        const length = c.u16()
        const start = rec.breaks[0]
        if (lastObj >= 0 && start !== undefined) {
          const t = new Cursor(rec.data, start, rec.breaks)
          const high = (t.u8() & 1) === 1
          texts.set(lastObj, t.chars(length, high))
        }
        break
      }
      case 0x001c: {
        const row = c.u16()
        const col = c.u16()
        const flags = c.u16()
        const id = c.u16()
        const author = c.string()
        // A threaded comment's placeholder for older Excel: the thread itself isn't in an .xls.
        if (author.startsWith("tc=")) break
        let text = (texts.get(id) ?? "").replace(/\r\n?/g, "\n")
        // Excel starts a note with its author's name in bold: "Name:\n".
        if (author && text.startsWith(author + ":")) text = text.slice(author.length + 1).replace(/^\n/, "")
        sheet.annotations.set(row, col, { kind: "note", author, text, shown: (flags & 0x02) !== 0 })
        break
      }
      case 0x01b1:
        left.add("conditional formats")
        break
      case 0x01be:
        left.add("data validation")
        break
      case 0x01b8:
        left.add("hyperlinks")
        break
      case 0x009d:
        left.add("AutoFilter")
        break
      case R.PROTECT:
        if (c.u16()) left.add("sheet protection")
        break
    }
  }
  if (skipped) warnings.push(`${sheet.name}: ${skipped} formulas Cells couldn't read keep their values.`)
  if (left.size) warnings.push(`${sheet.name}: Cells doesn't read ${[...left].join(", ")} from .xls files yet.`)
  return view
}

/** Opens an .xls file. */
export async function openXls(
  bytes: Uint8Array,
  _fileName = "Book1.xls",
  options: { calculate?: boolean; progress?: (done: number) => void | Promise<void> } = {},
): Promise<XlsRead> {
  let file: CompoundFile
  try {
    file = new CompoundFile(bytes)
  } catch (error) {
    throw new XlsError(error instanceof CfbError ? error.message : String(error), "damaged")
  }
  if (file.stream("EncryptionInfo") || file.stream("EncryptedPackage"))
    throw new XlsError("The workbook is protected with a password to open.", "encrypted")
  const stream = file.stream("Workbook")
  if (!stream) {
    if (file.stream("Book")) throw new XlsError("The file is in the Excel 5.0/95 format.", "old")
    throw new XlsError("The file has no workbook stream.", "damaged")
  }
  await options.progress?.(0.05)
  const records = readRecords(stream)
  if (records[0]?.type !== R.BOF) throw new XlsError("The workbook stream doesn't start with BOF.", "damaged")
  const version = new Cursor(records[0].data).u16()
  if (version !== 0x0600) throw new XlsError("The file is in the Excel 5.0/95 format.", "old")
  const globalsEnd = records.findIndex((r) => r.type === R.EOF)
  if (globalsEnd < 0) throw new XlsError("The workbook stream ends early.", "damaged")

  const warnings: string[] = []
  const g = new Globals()
  readGlobals(records, globalsEnd, g, warnings)
  const worksheets = g.sheets
  if (!worksheets.length) throw new XlsError("The workbook has no sheets.", "damaged")

  const wb = new Workbook(worksheets.map((s) => s.name))
  const styleIds = new Map<number, number>()
  const styleId = (xf: number) => {
    let id = styleIds.get(xf)
    if (id !== undefined) return id
    const rec = g.xfs[xf]
    id = rec && xf > 0 ? wb.styles.intern(xfStyle(g, rec)) : 0
    styleIds.set(xf, id)
    return id
  }

  // Names: the workbook's, with their formulas decoded once the sheets are known.
  const nameTexts = g.names.map((n) => (n.builtin ? `_xlnm.${n.name}` : n.name))
  const ctx = {
    sheetNames: worksheets.map((s) => s.name),
    xti: g.xti,
    names: nameTexts,
    externNames: (ixti: number) => {
      const book = g.books[g.xti[ixti]?.supbook ?? -1]
      return book?.self ? nameTexts : (book?.names ?? [])
    },
  }

  const offsets = new Map(records.map((r, i) => [r.offset, i]))
  const views: SheetView[] = []
  for (const [i, info] of worksheets.entries()) {
    const sheet = wb.sheets[i]
    sheet.hidden = info.state !== 0
    const from = offsets.get(info.offset)
    if (info.type !== 0 || from === undefined) {
      if (info.type !== 0)
        warnings.push(`${sheet.name} is a chart sheet or macro sheet; Cells opens it as an empty sheet.`)
      views.push({ ...DEFAULT_VIEW })
      continue
    }
    let to = from + 1
    for (let depth = 1; to < records.length && depth; to++) {
      if (records[to].type === R.BOF) depth++
      else if (records[to].type === R.EOF) depth--
    }
    views.push(readSheet(records, from + 1, to - 1, sheet, ctx, styleId, g, warnings))
    await options.progress?.(0.1 + (0.85 * (i + 1)) / worksheets.length)
  }

  const names: DefinedName[] = []
  for (const n of g.names) {
    let text: string
    try {
      text = decodeFormula(new Cursor(n.record.data, n.at, n.record.breaks), n.cce, { ...ctx, row: 0, col: 0 })
    } catch {
      continue
    }
    const sheetOf = n.sheet > 0 ? wb.sheets[n.sheet - 1] : undefined
    if (n.builtin === "Print_Area" && sheetOf) {
      const area = readPrintArea(text)
      if (area) sheetOf.pageSetup.printArea = area
      continue
    }
    if (n.builtin === "Print_Titles" && sheetOf) {
      const titles = readPrintTitles(text)
      if (titles) {
        sheetOf.pageSetup.titleRows = titles.rows
        sheetOf.pageSetup.titleCols = titles.cols
      }
      continue
    }
    if (n.builtin || n.hidden || n.sheet > 0) continue
    names.push({ name: n.name, ref: text })
  }
  wb.names.replace(names)

  let activeSheet = g.activeTab
  if (!(activeSheet >= 0 && activeSheet < wb.sheets.length) || wb.sheets[activeSheet].hidden)
    activeSheet = Math.max(
      0,
      wb.sheets.findIndex((s) => !s.hidden),
    )

  await options.progress?.(0.95)
  // Everything calculates again, as Excel does with a file from an older version: the saved results stand in until
  // then (some writers save none for text results). The worker that calculates on its own does it there.
  wb.calc.rebuild(false)
  if (options.calculate !== false) wb.calc.calculateNow(true)
  wb.clearHistory()
  return { workbook: wb, views, activeSheet, warnings: [...new Set(warnings)] }
}
