import { columnName } from "@cells/core"
import { escapeXml, readZip, storedZip, type ZipEntry } from "@cells/ooxml"
import { Cursor } from "../xls/biff"
import { isCompoundFile } from "../xls/cfb"
import { errorText, FormulaSkip } from "../xls/formula"
import { rkValue } from "../xls/read"
import { openXlsx, type OpenOptions, type XlsxRead } from "../xlsx/read"
import { decodeXlsbFormula, type XlsbFormulaContext, type XlsbTable } from "./formula"

/**
 * Excel Binary Workbooks (.xlsb, [MS-XLSB]): the same package as an .xlsx with the workbook, sheets, styles, strings,
 * comments and tables kept as records instead of XML. Cells writes each of those parts out as the XML an .xlsx has
 * and opens the result as one, so everything the .xlsx reader knows applies. Charts, drawings and the theme are XML
 * in an .xlsb already and come along as they are. Parts with no XML here (pivot tables, query tables, external links,
 * connections) are left out, with a warning.
 */

export class XlsbError extends Error {
  constructor(
    message: string,
    readonly kind: "encrypted" | "damaged",
  ) {
    super(message)
    this.name = "XlsbError"
  }
}

type Rec = { type: number; data: Uint8Array }

/** A part's records: a 1–2 byte type and a 1–4 byte size, 7 bits a byte ([MS-XLSB] 2.1.4). */
export function xlsbRecords(bytes: Uint8Array): Rec[] {
  const out: Rec[] = []
  let i = 0
  while (i < bytes.length) {
    let type = bytes[i++]
    if (type & 0x80) type = (type & 0x7f) | ((bytes[i++] & 0x7f) << 7)
    let size = 0
    for (let k = 0, shift = 0; k < 4; k++, shift += 7) {
      const b = bytes[i++]
      size += (b & 0x7f) * 2 ** shift
      if (!(b & 0x80)) break
    }
    if (i + size > bytes.length) throw new XlsbError("A record runs past the end of its part.", "damaged")
    out.push({ type, data: bytes.subarray(i, i + size) })
    i += size
  }
  return out
}

/** XLWideString: a 4-byte count of UTF-16 characters. */
function wide(c: Cursor): string {
  const n = c.u32()
  let s = ""
  for (let i = 0; i < n && c.left >= 2; i++) s += String.fromCharCode(c.u16())
  return s
}

/** XLNullableWideString: 0xFFFFFFFF for none. */
function nullable(c: Cursor): string | null {
  if (c.left >= 4 && c.data[c.pos] === 0xff && c.data[c.pos + 1] === 0xff && c.data[c.pos + 2] === 0xff) {
    if (c.data[c.pos + 3] === 0xff) {
      c.skip(4)
      return null
    }
  }
  return wide(c)
}

/** RichStr / a shared string: flags, the text (formatting runs and phonetic text are left out). */
function richText(c: Cursor): string {
  c.u8()
  return wide(c)
}

const cellRef = (row: number, col: number) => `${columnName(col)}${row + 1}`
const rangeRef = (r0: number, r1: number, c0: number, c1: number) =>
  r0 === r1 && c0 === c1 ? cellRef(r0, c0) : `${cellRef(r0, c0)}:${cellRef(r1, c1)}`

function rfx(c: Cursor) {
  const r0 = c.u32()
  const r1 = c.u32()
  const c0 = c.u32()
  const c1 = c.u32()
  return { r0, r1, c0, c1 }
}

const x = (s: string) => escapeXml(s)

/** BrtColor ([MS-XLSB] 2.5.24) as an .xlsx colour element's attributes. */
function color(c: Cursor): string {
  const flags = c.u8()
  const index = c.u8()
  const tint = c.i16()
  const r = c.u8()
  const g = c.u8()
  const b = c.u8()
  const a = c.u8()
  const kind = flags >> 1
  const tintAttr = tint ? ` tint="${Math.round((tint / 32767) * 1e6) / 1e6}"` : ""
  const hex = (n: number) => n.toString(16).padStart(2, "0").toUpperCase()
  switch (kind) {
    case 0:
      return ` auto="1"`
    case 1:
      return ` indexed="${index}"${tintAttr}`
    case 3:
      return ` theme="${index}"${tintAttr}`
    default:
      return ` rgb="${hex(a)}${hex(r)}${hex(g)}${hex(b)}"${tintAttr}`
  }
}

const PATTERNS = [
  "none",
  "solid",
  "mediumGray",
  "darkGray",
  "lightGray",
  "darkHorizontal",
  "darkVertical",
  "darkDown",
  "darkUp",
  "darkGrid",
  "darkTrellis",
  "lightHorizontal",
  "lightVertical",
  "lightDown",
  "lightUp",
  "lightGrid",
  "lightTrellis",
  "gray125",
  "gray0625",
]
const BORDER_STYLES = [
  "none",
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
const H_ALIGN = ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]
const V_ALIGN = ["top", "center", "bottom", "justify", "distributed"]
const UNDERLINE: Record<number, string> = {
  1: "single",
  2: "double",
  0x21: "singleAccounting",
  0x22: "doubleAccounting",
}

/** styles.bin as styles.xml. */
function stylesXml(bytes: Uint8Array): string {
  const formats: string[] = []
  const fonts: string[] = []
  const fills: string[] = []
  const borders: string[] = []
  const styleXfs: string[] = []
  const cellXfs: string[] = []
  const cellStyles: string[] = []
  let inCellXfs = false
  let inStyleXfs = false
  for (const rec of xlsbRecords(bytes)) {
    const c = new Cursor(rec.data)
    switch (rec.type) {
      case 0x2c: {
        // BrtFmt
        const id = c.u16()
        formats.push(`<numFmt numFmtId="${id}" formatCode="${x(wide(c))}"/>`)
        break
      }
      case 0x2b: {
        // BrtFont
        const height = c.u16()
        const grbit = c.u16()
        const weight = c.u16()
        const script = c.u16()
        const underline = c.u8()
        const family = c.u8()
        const charset = c.u8()
        c.skip(1)
        const col = color(c)
        const scheme = c.u8()
        const name = wide(c)
        let f = "<font>"
        if (weight >= 700) f += "<b/>"
        if (grbit & 0x02) f += "<i/>"
        if (grbit & 0x08) f += "<strike/>"
        if (grbit & 0x10) f += "<outline/>"
        if (grbit & 0x20) f += "<shadow/>"
        if (grbit & 0x40) f += "<condense/>"
        if (grbit & 0x80) f += "<extend/>"
        if (UNDERLINE[underline]) f += underline === 1 ? "<u/>" : `<u val="${UNDERLINE[underline]}"/>`
        if (script === 1) f += `<vertAlign val="superscript"/>`
        if (script === 2) f += `<vertAlign val="subscript"/>`
        f += `<sz val="${height / 20}"/><color${col}/><name val="${x(name)}"/>`
        if (family) f += `<family val="${family}"/>`
        if (charset) f += `<charset val="${charset}"/>`
        if (scheme === 1) f += `<scheme val="major"/>`
        if (scheme === 2) f += `<scheme val="minor"/>`
        fonts.push(f + "</font>")
        break
      }
      case 0x2d: {
        // BrtFill: pattern fills (a gradient opens as no fill)
        const pattern = c.u32()
        const fore = color(c)
        const back = color(c)
        const type = PATTERNS[pattern] ?? "none"
        fills.push(
          type === "none"
            ? `<fill><patternFill patternType="none"/></fill>`
            : `<fill><patternFill patternType="${type}"><fgColor${fore}/><bgColor${back}/></patternFill></fill>`,
        )
        break
      }
      case 0x2e: {
        // BrtBorder: top, bottom, left, right, diagonal
        const flags = c.u8()
        const side = () => {
          const dg = c.u8()
          c.skip(1)
          const col = color(c)
          return { style: BORDER_STYLES[dg] ?? "none", col }
        }
        const top = side()
        const bottom = side()
        const left = side()
        const right = side()
        const diag = side()
        const el = (tag: string, s: { style: string; col: string }) =>
          s.style === "none" ? `<${tag}/>` : `<${tag} style="${s.style}"><color${s.col}/></${tag}>`
        const diagAttrs = `${flags & 0x01 ? ` diagonalDown="1"` : ""}${flags & 0x02 ? ` diagonalUp="1"` : ""}`
        borders.push(
          `<border${diagAttrs}>${el("left", left)}${el("right", right)}${el("top", top)}${el("bottom", bottom)}${el("diagonal", diag)}</border>`,
        )
        break
      }
      case 0x272:
        inStyleXfs = true
        break
      case 0x273:
        inStyleXfs = false
        break
      case 0x269:
        inCellXfs = true
        break
      case 0x26a:
        inCellXfs = false
        break
      case 0x2f: {
        // BrtXF
        const parent = c.u16()
        const fmt = c.u16()
        const font = c.u16()
        const fill = c.u16()
        const border = c.u16()
        const rotation = c.u8()
        const indent = c.u8()
        const flags = c.u16()
        const h = H_ALIGN[flags & 0x07] ?? "general"
        const v = V_ALIGN[(flags >> 3) & 0x07] ?? "bottom"
        let align = ""
        if (h !== "general") align += ` horizontal="${h}"`
        if (v !== "bottom") align += ` vertical="${v}"`
        if (flags & 0x40) align += ` wrapText="1"`
        if (flags & 0x80) align += ` justifyLastLine="1"`
        if (flags & 0x100) align += ` shrinkToFit="1"`
        if ((flags >> 10) & 0x03) align += ` readingOrder="${(flags >> 10) & 0x03}"`
        if (rotation) align += ` textRotation="${rotation}"`
        if (indent) align += ` indent="${indent}"`
        const locked = (flags & 0x1000) !== 0
        const hidden = (flags & 0x2000) !== 0
        const protection =
          !locked || hidden ? `<protection${locked ? "" : ` locked="0"`}${hidden ? ` hidden="1"` : ""}/>` : ""
        const inner = (align ? `<alignment${align}/>` : "") + protection
        const xfId = inCellXfs && parent !== 0xffff ? ` xfId="${parent}"` : ""
        const xf = `<xf numFmtId="${fmt}" fontId="${font}" fillId="${fill}" borderId="${border}"${xfId}${inner ? ` applyAlignment="1" applyProtection="1">${inner}</xf>` : "/>"}`
        if (inCellXfs) cellXfs.push(xf)
        else if (inStyleXfs) styleXfs.push(xf)
        break
      }
      case 0x30: {
        // BrtStyle
        const xf = c.u32()
        const grbit = c.u16()
        const builtin = c.u8()
        const level = c.u8()
        const name = wide(c)
        const b =
          grbit & 0x01 ? ` builtinId="${builtin}"${builtin === 1 || builtin === 2 ? ` iLevel="${level}"` : ""}` : ""
        cellStyles.push(`<cellStyle name="${x(name)}" xfId="${xf}"${b}${grbit & 0x02 ? ` hidden="1"` : ""}/>`)
        break
      }
    }
  }
  const list = (tag: string, items: string[]) =>
    items.length ? `<${tag} count="${items.length}">${items.join("")}</${tag}>` : ""
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    list("numFmts", formats) +
    list("fonts", fonts) +
    list("fills", fills) +
    list("borders", borders) +
    list("cellStyleXfs", styleXfs) +
    list("cellXfs", cellXfs) +
    list("cellStyles", cellStyles) +
    `</styleSheet>`
  )
}

function sharedStrings(bytes: Uint8Array): string[] {
  const out: string[] = []
  for (const rec of xlsbRecords(bytes)) if (rec.type === 0x13) out.push(richText(new Cursor(rec.data)))
  return out
}

const preserve = (s: string) => (/^\s|\s$|\n/.test(s) ? ` xml:space="preserve"` : "")
const tText = (s: string) => `<t${preserve(s)}>${x(s)}</t>`

type Workbook = {
  sheets: { name: string; rel: string; state: number }[]
  names: { name: string; scope: number; hidden: boolean; func: boolean; rgce: Uint8Array; cce: number }[]
  xti: ([number, number] | null)[]
  activeTab: number
  firstTab: number
  date1904: boolean
  bookView: string
  calc: string
}

function readWorkbook(bytes: Uint8Array): Workbook {
  const wb: Workbook = {
    sheets: [],
    names: [],
    xti: [],
    activeTab: 0,
    firstTab: 0,
    date1904: false,
    bookView: "",
    calc: "",
  }
  const books: ("self" | "other")[] = []
  for (const rec of xlsbRecords(bytes)) {
    const c = new Cursor(rec.data)
    switch (rec.type) {
      case 0x9c: {
        // BrtBundleSh
        const state = c.u32()
        c.u32()
        const rel = wide(c)
        wb.sheets.push({ state, rel, name: wide(c) })
        break
      }
      case 0x99:
        // BrtWbProp
        wb.date1904 = (c.u32() & 0x01) !== 0
        break
      case 0x9e: {
        // BrtBookView
        const xWn = c.u32() | 0
        const yWn = c.u32() | 0
        const dx = c.u32()
        const dy = c.u32()
        const ratio = c.u32()
        wb.firstTab = c.u32()
        wb.activeTab = c.u32()
        wb.bookView = `<workbookView xWindow="${xWn}" yWindow="${yWn}" windowWidth="${dx}" windowHeight="${dy}" tabRatio="${ratio}"${wb.firstTab ? ` firstSheet="${wb.firstTab}"` : ""}${wb.activeTab ? ` activeTab="${wb.activeTab}"` : ""}/>`
        break
      }
      case 0x9d: {
        // BrtCalcProp: the calculation id, then the mode (0 manual, 1 automatic, 2 automatic except tables)
        c.u32()
        const mode = c.u32()
        if (mode === 0) wb.calc = `<calcPr calcMode="manual"/>`
        else if (mode === 2) wb.calc = `<calcPr calcMode="autoNoTable"/>`
        break
      }
      case 0x165: // BrtSupSelf
      case 0x166: // BrtSupSame
        books.push("self")
        break
      case 0x163: // BrtSupBookSrc: another workbook
      case 0x29a: // BrtSupAddin
        books.push("other")
        break
      case 0x16a: {
        // BrtExternSheet
        const n = c.u32()
        for (let i = 0; i < n; i++) {
          const book = c.u32()
          const first = c.u32() | 0
          const last = c.u32() | 0
          wb.xti.push(books[book] === "other" || first < 0 ? null : [first, last])
        }
        break
      }
      case 0x27: {
        // BrtName
        const flags = c.u32()
        c.u8()
        const scope = c.u32()
        const name = wide(c)
        const cce = c.u32()
        const at = c.pos
        c.skip(cce)
        const cb = c.u32()
        wb.names.push({
          // Built-in names (Print_Area, Print_Titles, _FilterDatabase) are _xlnm. names in an .xlsx.
          name: flags & 0x20 ? `_xlnm.${name}` : name,
          scope: scope === 0xffffffff ? -1 : scope,
          hidden: (flags & 0x01) !== 0,
          func: (flags & 0x02) !== 0,
          cce,
          rgce: rec.data.subarray(at, c.pos + cb),
        })
        break
      }
    }
  }
  return wb
}

type Table = XlsbTable & { xml: string; id: number }

const TOTALS = ["none", "average", "count", "countNums", "max", "min", "sum", "stdDev", "var", "custom"]

function readTable(bytes: Uint8Array): Table | null {
  let table: Table | null = null
  const cols: string[] = []
  let style = ""
  let filter = ""
  for (const rec of xlsbRecords(bytes)) {
    const c = new Cursor(rec.data)
    switch (rec.type) {
      case 0x157: {
        // BrtBeginList
        const r = rfx(c)
        c.u32()
        const id = c.u32()
        const headers = c.u32()
        const totals = c.u32()
        c.skip(4 + 4 * 6 + 4)
        const name = nullable(c) ?? `Table${id}`
        const display = nullable(c) ?? name
        const comment = nullable(c)
        table = {
          id,
          name: display,
          columns: [],
          xml: `id="${id}" name="${x(name)}" displayName="${x(display)}" ref="${rangeRef(r.r0, r.r1, r.c0, r.c1)}"${headers ? "" : ` headerRowCount="0"`}${totals ? ` totalsRowCount="${totals}"` : ` totalsRowShown="0"`}${comment ? ` comment="${x(comment)}"` : ""}`,
        }
        break
      }
      case 0x15b: {
        // BrtBeginListCol
        const id = c.u32()
        const total = c.u32()
        c.skip(16)
        const name = nullable(c)
        const caption = nullable(c) ?? name ?? `Column${id}`
        const label = nullable(c)
        table?.columns.push(caption)
        const fn = TOTALS[total] && total ? ` totalsRowFunction="${TOTALS[total]}"` : ""
        cols.push(`<tableColumn id="${id}" name="${x(caption)}"${label ? ` totalsRowLabel="${x(label)}"` : ""}${fn}/>`)
        break
      }
      case 0xa1: {
        // BrtBeginAFilter
        const r = rfx(c)
        filter = `<autoFilter ref="${rangeRef(r.r0, r.r1, r.c0, r.c1)}"/>`
        break
      }
      case 0x201: {
        // BrtTableStyleClient
        const flags = c.u16()
        const name = nullable(c)
        style = `<tableStyleInfo${name ? ` name="${x(name)}"` : ""} showFirstColumn="${flags & 0x01 ? 1 : 0}" showLastColumn="${flags & 0x02 ? 1 : 0}" showRowStripes="${flags & 0x04 ? 1 : 0}" showColumnStripes="${flags & 0x08 ? 1 : 0}"/>`
        break
      }
    }
  }
  if (!table) return null
  table.xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${table.xml}>` +
    filter +
    `<tableColumns count="${cols.length}">${cols.join("")}</tableColumns>${style}</table>`
  return table
}

function commentsXml(bytes: Uint8Array): string {
  const authors: string[] = []
  const list: string[] = []
  let current: { ref: string; author: number } | null = null
  for (const rec of xlsbRecords(bytes)) {
    const c = new Cursor(rec.data)
    switch (rec.type) {
      case 0x278:
        authors.push(wide(c))
        break
      case 0x27b: {
        const author = c.u32()
        const r = rfx(c)
        current = { author, ref: cellRef(r.r0, r.c0) }
        break
      }
      case 0x27d:
        if (current)
          list.push(
            `<comment ref="${current.ref}" authorId="${current.author}"><text>${tText(richText(c))}</text></comment>`,
          )
        current = null
        break
    }
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<authors>${authors.map((a) => `<author>${x(a)}</author>`).join("")}</authors><commentList>${list.join("")}</commentList></comments>`
  )
}

type SheetResult = { xml: string; skipped: Set<string>; formulasLost: number }

const PANES = ["bottomRight", "topRight", "bottomLeft", "topLeft"]

function sheetXml(bytes: Uint8Array, ctx: Omit<XlsbFormulaContext, "row" | "col">): SheetResult {
  const recs = xlsbRecords(bytes)
  const skipped = new Set<string>()
  let formulasLost = 0
  const shared = new Map<string, { r0: number; r1: number; c0: number; c1: number; data: Uint8Array; at: number }>()
  const arrays = new Map<string, { r0: number; r1: number; c0: number; c1: number; data: Uint8Array; at: number }>()
  // Shared and array formulas follow their first cell: collect them first.
  for (const rec of recs) {
    if (rec.type !== 0x1ab && rec.type !== 0x1aa) continue
    const c = new Cursor(rec.data)
    const r = rfx(c)
    if (rec.type === 0x1aa) {
      c.u8()
      arrays.set(`${r.r0},${r.c0}`, { ...r, data: rec.data, at: c.pos })
    } else shared.set(`${r.r0},${r.c0}`, { ...r, data: rec.data, at: c.pos })
  }

  const decode = (data: Uint8Array, at: number, row: number, col: number): string | null => {
    const c = new Cursor(data, at)
    const cce = c.u32()
    try {
      return decodeXlsbFormula(c, cce, { ...ctx, row, col })
    } catch (error) {
      if (error instanceof FormulaSkip) return null
      throw error
    }
  }

  let sheetPr = ""
  let dimension = ""
  let view = ""
  let pane = ""
  const selections: string[] = []
  let format = ""
  let defaultHeight = 300
  const cols: string[] = []
  const rows: string[] = []
  let row = -1
  let rowAttrs = ""
  let cells: string[] = []
  let lastCol = -1
  const merges: string[] = []
  const links: string[] = []
  let margins = ""
  let printOptions = ""
  let pageSetup = ""
  let headerFooter = ""
  let drawing = ""
  let legacy = ""
  let legacyHF = ""
  let autoFilter = ""
  const tableParts: string[] = []

  const flushRow = () => {
    if (row >= 0 && (cells.length || rowAttrs)) rows.push(`<row r="${row + 1}"${rowAttrs}>${cells.join("")}</row>`)
    cells = []
    rowAttrs = ""
  }

  for (const rec of recs) {
    const c = new Cursor(rec.data)
    const t = rec.type
    if ((t >= 0x01 && t <= 0x0b) || (t >= 0x0c && t <= 0x12)) {
      // A cell: BrtCell* carry their column; BrtShort* are the cell after the last
      const short = t >= 0x0c
      const col = short ? lastCol + 1 : c.u32()
      lastCol = col
      const style = c.u32() & 0xffffff
      const s = style ? ` s="${style}"` : ""
      const ref = cellRef(row, col)
      const kind = short ? t - 0x0b : t
      let value = ""
      let type = ""
      switch (kind) {
        case 0x01:
          break
        case 0x02:
          value = String(rkValue(c.u32()))
          break
        case 0x03:
        case 0x0b:
          type = "e"
          value = errorText(c.u8())
          break
        case 0x04:
        case 0x0a:
          type = "b"
          value = c.u8() ? "1" : "0"
          break
        case 0x05:
        case 0x09:
          value = String(c.f64())
          break
        case 0x06:
        case 0x08:
          type = kind === 0x06 ? "inlineStr" : "str"
          value = wide(c)
          break
        case 0x07:
          type = "s"
          value = String(c.u32())
          break
      }
      let f = ""
      if (!short && kind >= 0x08) {
        c.u16() // grbitFlags
        const at = c.pos
        c.u32()
        const first = rec.data[at + 4]
        let text: string | null = null
        if (first === 0x01) {
          // PtgExp: part of a shared or array formula whose first cell is in the row it names
          const anchorRow = new DataView(rec.data.buffer, rec.data.byteOffset + at + 5, 4).getUint32(0, true)
          const inside = (v: { r0: number; r1: number; c0: number; c1: number }) =>
            row >= v.r0 && row <= v.r1 && col >= v.c0 && col <= v.c1
          const array = arrays.get(`${row},${col}`)
          if (array) {
            const body = decode(array.data, array.at, row, col)
            if (body !== null)
              f = `<f t="array" ref="${rangeRef(array.r0, array.r1, array.c0, array.c1)}">${x(body)}</f>`
            else formulasLost++
          } else if (![...arrays.values()].some(inside)) {
            const sh = [...shared.values()].find((v) => v.r0 === anchorRow && inside(v))
            text = sh ? decode(sh.data, sh.at, row, col) : null
            if (text === null) formulasLost++
          }
        } else {
          text = decode(rec.data, at, row, col)
          if (text === null) formulasLost++
        }
        if (text !== null) f = `<f>${x(text)}</f>`
      }
      const v =
        type === "inlineStr" ? `<is>${tText(value)}</is>` : value !== "" || type === "str" ? `<v>${x(value)}</v>` : ""
      if (kind === 0x01 && !style) continue
      cells.push(`<c r="${ref}"${s}${type ? ` t="${type}"` : ""}>${f}${v}</c>`)
      continue
    }
    switch (t) {
      case 0x00: {
        // BrtRowHdr
        flushRow()
        row = c.u32()
        const xf = c.u32()
        const height = c.u16()
        c.u8()
        const flags = c.u8()
        lastCol = -1
        let a = ""
        if (flags & 0x40) a += ` s="${xf}" customFormat="1"`
        if (flags & 0x20 || height !== defaultHeight)
          a += ` ht="${height / 20}"${flags & 0x20 ? ` customHeight="1"` : ""}`
        if (flags & 0x10) a += ` hidden="1"`
        if (flags & 0x07) a += ` outlineLevel="${flags & 0x07}"`
        if (flags & 0x08) a += ` collapsed="1"`
        rowAttrs = a
        break
      }
      case 0x93: {
        // BrtWsProp: the tab colour after two flag bytes and one more
        c.skip(3)
        const flags = rec.data[3]
        const col = color(c)
        if (flags >> 1) sheetPr = `<sheetPr><tabColor${col}/></sheetPr>`
        break
      }
      case 0x94: {
        const r = rfx(c)
        dimension = `<dimension ref="${rangeRef(r.r0, r.r1, r.c0, r.c1)}"/>`
        break
      }
      case 0x89: {
        // BrtBeginWsView
        const flags = c.u16()
        const mode = c.u32()
        const top = c.u32()
        const left = c.u32()
        c.skip(4)
        const scale = c.u16()
        const normal = c.u16()
        const pbp = c.u16()
        const layout = c.u16()
        let a = ""
        if (flags & 0x02) a += ` showFormulas="1"`
        if (!(flags & 0x04)) a += ` showGridLines="0"`
        if (!(flags & 0x08)) a += ` showRowColHeaders="0"`
        if (!(flags & 0x10)) a += ` showZeros="0"`
        if (flags & 0x20) a += ` rightToLeft="1"`
        if (flags & 0x40) a += ` tabSelected="1"`
        if (!(flags & 0x80)) a += ` showRuler="0"`
        if (mode === 1) a += ` view="pageBreakPreview"`
        if (mode === 2) a += ` view="pageLayout"`
        if (top || left) a += ` topLeftCell="${cellRef(top, left)}"`
        if (scale && scale !== 100) a += ` zoomScale="${scale}"`
        if (normal) a += ` zoomScaleNormal="${normal}"`
        if (pbp) a += ` zoomScaleSheetLayoutView="${pbp}"`
        if (layout) a += ` zoomScalePageLayoutView="${layout}"`
        if (!view) view = a
        break
      }
      case 0x97: {
        // BrtPane
        const xs = c.f64()
        const ys = c.f64()
        const top = c.u32()
        const left = c.u32()
        const active = c.u32()
        const flags = c.u8()
        const state = flags & 0x01 ? (flags & 0x02 ? "frozen" : "frozenSplit") : "split"
        pane = `<pane${xs ? ` xSplit="${xs}"` : ""}${ys ? ` ySplit="${ys}"` : ""} topLeftCell="${cellRef(top, left)}" activePane="${PANES[active] ?? "bottomRight"}" state="${state}"/>`
        break
      }
      case 0x98: {
        // BrtSel
        const p = c.u32()
        const r = c.u32()
        const col = c.u32()
        c.u32()
        const n = c.u32()
        const areas: string[] = []
        for (let i = 0; i < n; i++) {
          const a = rfx(c)
          areas.push(rangeRef(a.r0, a.r1, a.c0, a.c1))
        }
        selections.push(
          `<selection${pane ? ` pane="${PANES[p] ?? "bottomRight"}"` : ""} activeCell="${cellRef(r, col)}" sqref="${areas.join(" ")}"/>`,
        )
        break
      }
      case 0x1e5: {
        // BrtWsFmtInfo
        const width = c.u32()
        const base = c.u16()
        const height = c.u16()
        const flags = c.u16()
        defaultHeight = height
        format = `<sheetFormatPr baseColWidth="${base}"${width !== 0xffffffff ? ` defaultColWidth="${width / 256}"` : ""} defaultRowHeight="${height / 20}"${flags & 0x01 ? ` customHeight="1"` : ""}${flags & 0x02 ? ` zeroHeight="1"` : ""}/>`
        break
      }
      case 0x3c: {
        // BrtColInfo
        const first = c.u32()
        const last = c.u32()
        const width = c.u32()
        const xf = c.u32()
        const flags = c.u16()
        let a = ` min="${first + 1}" max="${last + 1}" width="${width / 256}"`
        if (xf) a += ` style="${xf}"`
        if (flags & 0x01) a += ` hidden="1"`
        if (flags & 0x02) a += ` customWidth="1"`
        if (flags & 0x04) a += ` bestFit="1"`
        if ((flags >> 8) & 0x07) a += ` outlineLevel="${(flags >> 8) & 0x07}"`
        if (flags & 0x1000) a += ` collapsed="1"`
        cols.push(`<col${a}/>`)
        break
      }
      case 0xb0: {
        const r = rfx(c)
        merges.push(`<mergeCell ref="${rangeRef(r.r0, r.r1, r.c0, r.c1)}"/>`)
        break
      }
      case 0x1ee: {
        // BrtHLink
        const r = rfx(c)
        const rel = nullable(c)
        const location = wide(c)
        const tip = wide(c)
        const display = wide(c)
        links.push(
          `<hyperlink ref="${rangeRef(r.r0, r.r1, r.c0, r.c1)}"${rel ? ` r:id="${x(rel)}"` : ""}${location ? ` location="${x(location)}"` : ""}${tip ? ` tooltip="${x(tip)}"` : ""}${display ? ` display="${x(display)}"` : ""}/>`,
        )
        break
      }
      case 0x1dc: {
        const [l, r, tp, b, h, f] = [c.f64(), c.f64(), c.f64(), c.f64(), c.f64(), c.f64()]
        margins = `<pageMargins left="${l}" right="${r}" top="${tp}" bottom="${b}" header="${h}" footer="${f}"/>`
        break
      }
      case 0x1dd: {
        const flags = c.u16()
        let a = ""
        if (flags & 0x01) a += ` horizontalCentered="1"`
        if (flags & 0x02) a += ` verticalCentered="1"`
        if (flags & 0x04) a += ` headings="1"`
        if (flags & 0x08) a += ` gridLines="1"`
        if (a) printOptions = `<printOptions${a}/>`
        break
      }
      case 0x1de: {
        // BrtPageSetup
        const paper = c.u32()
        const scale = c.u32()
        c.skip(8)
        const copies = c.u32()
        const firstPage = c.u32() | 0
        const fitW = c.u32()
        const fitH = c.u32()
        const flags = c.u16()
        let a = ""
        if (paper && paper !== 1) a += ` paperSize="${paper}"`
        if (scale && scale !== 100) a += ` scale="${scale}"`
        if (copies > 1) a += ` copies="${copies}"`
        if (flags & 0x80) a += ` firstPageNumber="${firstPage}" useFirstPageNumber="1"`
        if (fitW !== 1) a += ` fitToWidth="${fitW}"`
        if (fitH !== 1) a += ` fitToHeight="${fitH}"`
        if (!(flags & 0x40)) a += ` orientation="${flags & 0x02 ? "landscape" : "portrait"}"`
        if (flags & 0x01) a += ` pageOrder="overThenDown"`
        if (flags & 0x08) a += ` blackAndWhite="1"`
        if (flags & 0x10) a += ` draft="1"`
        pageSetup = a ? `<pageSetup${a}/>` : ""
        break
      }
      case 0x1df: {
        // BrtBeginHeaderFooter
        const flags = c.u16()
        const [h, f, he, fe, hf, ff] = [nullable(c), nullable(c), nullable(c), nullable(c), nullable(c), nullable(c)]
        const el = (tag: string, s: string | null) => (s ? `<${tag}>${x(s)}</${tag}>` : "")
        const inner =
          el("oddHeader", h) +
          el("oddFooter", f) +
          el("evenHeader", he) +
          el("evenFooter", fe) +
          el("firstHeader", hf) +
          el("firstFooter", ff)
        let a = ""
        if (flags & 0x01) a += ` differentOddEven="1"`
        if (flags & 0x02) a += ` differentFirst="1"`
        if (!(flags & 0x04)) a += ` scaleWithDoc="0"`
        if (!(flags & 0x08)) a += ` alignWithMargins="0"`
        if (inner || a) headerFooter = `<headerFooter${a}>${inner}</headerFooter>`
        break
      }
      case 0x226:
        drawing = `<drawing r:id="${x(wide(c))}"/>`
        break
      case 0x227:
        legacy = `<legacyDrawing r:id="${x(wide(c))}"/>`
        break
      case 0x228:
        legacyHF = `<legacyDrawingHF r:id="${x(wide(c))}"/>`
        break
      case 0x295:
        tableParts.push(`<tablePart r:id="${x(wide(c))}"/>`)
        break
      case 0xa1: {
        const r = rfx(c)
        autoFilter = `<autoFilter ref="${rangeRef(r.r0, r.r1, r.c0, r.c1)}"/>`
        break
      }
      case 0x1cd:
        skipped.add("conditional formatting")
        break
      case 0x23d:
        skipped.add("data validation")
        break
      case 0x217:
        break
    }
  }
  flushRow()
  const sheetView = `<sheetViews><sheetView workbookViewId="0"${view}>${pane}${selections.join("")}</sheetView></sheetViews>`
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    sheetPr +
    dimension +
    sheetView +
    format +
    (cols.length ? `<cols>${cols.join("")}</cols>` : "") +
    `<sheetData>${rows.join("")}</sheetData>` +
    autoFilter +
    (merges.length ? `<mergeCells count="${merges.length}">${merges.join("")}</mergeCells>` : "") +
    (links.length ? `<hyperlinks>${links.join("")}</hyperlinks>` : "") +
    printOptions +
    margins +
    pageSetup +
    headerFooter +
    drawing +
    legacy +
    legacyHF +
    (tableParts.length ? `<tableParts count="${tableParts.length}">${tableParts.join("")}</tableParts>` : "") +
    `</worksheet>`
  return { xml, skipped, formulasLost }
}

const decoder = new TextDecoder()
const encoder = new TextEncoder()

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const RELATIONSHIPS = /<Relationship\b[^>]*\/?>/g
const attr = (tag: string, name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1]

/** Resolves a relationship target against the folder of the part that refers to it. */
function resolve(from: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1)
  const parts = from.split("/").slice(0, -1)
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop()
    else if (seg !== ".") parts.push(seg)
  }
  return parts.join("/")
}

const relsOf = (part: string) => {
  const i = part.lastIndexOf("/")
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`
}

const unescape = (s: string) =>
  s.replace(/&(lt|gt|quot|apos|amp);/g, (_, e: string) => ({ lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" })[e]!)

const MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml"
const TYPES: Record<string, string> = {
  workbook: `${MAIN}.sheet.main+xml`,
  workbookMacro: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  worksheet: `${MAIN}.worksheet+xml`,
  styles: `${MAIN}.styles+xml`,
  sharedStrings: `${MAIN}.sharedStrings+xml`,
  comments: `${MAIN}.comments+xml`,
  table: `${MAIN}.table+xml`,
}

/** The .xlsb's bytes as an .xlsx package (and what couldn't come along). */
export async function xlsbToXlsx(
  bytes: Uint8Array,
): Promise<{ bytes: Uint8Array; warnings: string[]; macros: boolean }> {
  if (isCompoundFile(bytes)) throw new XlsbError("The workbook is protected with a password to open.", "encrypted")
  let entries: ZipEntry[]
  try {
    entries = await readZip(bytes)
  } catch (error) {
    throw new XlsbError(error instanceof Error ? error.message : String(error), "damaged")
  }
  const files = new Map(entries.map((e) => [e.normalizedName, e.data]))
  const lower = new Map(entries.map((e) => [e.normalizedName.toLowerCase(), e.normalizedName]))
  const get = (name: string) => files.get(lower.get(name.toLowerCase()) ?? name)
  const text = (name: string) => {
    const b = get(name)
    return b ? decoder.decode(b) : null
  }
  const rootRels = text("_rels/.rels") ?? ""
  const mainTag = [...rootRels.matchAll(RELATIONSHIPS)].find((m) => attr(m[0], "Type")?.endsWith("/officeDocument"))
  const mainPart = mainTag ? resolve("", unescape(attr(mainTag[0], "Target") ?? "")) : "xl/workbook.bin"
  const wbBytes = get(mainPart)
  if (!wbBytes) throw new XlsbError("The file has no workbook part.", "damaged")

  const warnings: string[] = []
  const out = new Map<string, Uint8Array>()
  const types = new Map<string, string>()
  const put = (name: string, xml: string, type?: string) => {
    out.set(name, encoder.encode(xml))
    if (type) types.set(name, type)
  }
  const renamed = new Map<string, string>()
  const xmlName = (part: string) => part.replace(/\.bin$/i, ".xml")

  // Each part's relationships, by kind.
  type Rel = { id: string; type: string; target: string; external: boolean; tag: string }
  const rels = (part: string): Rel[] =>
    [...(text(relsOf(part)) ?? "").matchAll(RELATIONSHIPS)].map((m) => {
      const external = attr(m[0], "TargetMode") === "External"
      const target = unescape(attr(m[0], "Target") ?? "")
      return {
        id: attr(m[0], "Id") ?? "",
        type: attr(m[0], "Type") ?? "",
        external,
        target: external ? target : resolve(part, target),
        tag: m[0],
      }
    })
  const kind = (r: Rel) => r.type.slice(r.type.lastIndexOf("/") + 1)

  const book = readWorkbook(wbBytes)
  const wbRels = rels(mainPart)
  const relById = new Map(wbRels.map((r) => [r.id, r]))
  const strings = (() => {
    const r = wbRels.find((r) => kind(r) === "sharedStrings")
    const b = r && get(r.target)
    return b ? sharedStrings(b) : []
  })()

  // Tables first: any sheet's formulas may refer to any table.
  const tables = new Map<number, Table>()
  const tableFor = new Map<string, Table>()
  for (const s of book.sheets) {
    const r = relById.get(s.rel)
    if (!r) continue
    for (const t of rels(r.target)) {
      if (kind(t) !== "table") continue
      const b = get(t.target)
      const table = b && readTable(b)
      if (table) {
        tables.set(table.id, table)
        tableFor.set(t.target, table)
      }
    }
  }

  const ctxNames = book.names.map((n) => n.name)
  const ctx = {
    sheetNames: book.sheets.map((s) => s.name),
    xti: book.xti,
    names: ctxNames,
    tables: new Map([...tables].map(([id, t]) => [id, { name: t.name, columns: t.columns }])),
  }

  let lost = 0
  const sheetsXml: string[] = []
  const kept: Rel[] = []
  for (const [i, s] of book.sheets.entries()) {
    const r = relById.get(s.rel)
    if (!r) continue
    const state = s.state === 1 ? ` state="hidden"` : s.state === 2 ? ` state="veryHidden"` : ""
    const sheetKind = kind(r)
    if (sheetKind !== "worksheet" || !/\.bin$/i.test(r.target)) {
      // A chart sheet, dialog or macro sheet: an empty sheet stands in.
      warnings.push(`${s.name} is a chart sheet or macro sheet; Cells opens it as an empty sheet.`)
      const name = `xl/worksheets/cells-empty${i + 1}.xml`
      put(
        name,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
        TYPES.worksheet,
      )
      kept.push({ ...r, type: `${REL}/worksheet`, target: name })
      sheetsXml.push(`<sheet name="${x(s.name)}" sheetId="${i + 1}"${state} r:id="${x(s.rel)}"/>`)
      continue
    }
    const b = get(r.target)
    if (!b) continue
    const result = sheetXml(b, ctx)
    lost += result.formulasLost
    for (const what of result.skipped)
      warnings.push(`${s.name} has ${what}, which Cells doesn't open from .xlsb files yet.`)
    const name = xmlName(r.target)
    renamed.set(r.target, name)
    put(name, result.xml, TYPES.worksheet)
    kept.push({ ...r, target: name })
    sheetsXml.push(`<sheet name="${x(s.name)}" sheetId="${i + 1}"${state} r:id="${x(s.rel)}"/>`)

    // The sheet's own parts: comments and tables written as XML; drawings and VML as they are.
    const own: Rel[] = []
    for (const t of rels(r.target)) {
      const k = kind(t)
      if (t.external || k === "hyperlink") own.push(t)
      else if (k === "comments") {
        const cb = get(t.target)
        if (!cb) continue
        const cname = xmlName(t.target)
        put(cname, commentsXml(cb), TYPES.comments)
        own.push({ ...t, target: cname })
      } else if (k === "table") {
        const table = tableFor.get(t.target)
        if (!table) continue
        const tname = xmlName(t.target)
        put(tname, table.xml, TYPES.table)
        own.push({ ...t, target: tname })
      } else if (/\.bin$/i.test(t.target) && k !== "printerSettings") {
        if (k !== "xlBinaryIndex")
          warnings.push(`${s.name} has a ${k} part, which Cells doesn't open from .xlsb files.`)
      } else own.push(t)
    }
    put(relsOf(name), relsXml(name, own))
  }

  // Names: the workbook's, decoded now that every sheet and table is known.
  const definedNames: string[] = []
  for (const n of book.names) {
    if (/^_xl(fn|pm|ws)\./i.test(n.name)) continue
    const c = new Cursor(n.rgce)
    let body: string
    try {
      body = decodeXlsbFormula(c, n.cce, { ...ctx, row: 0, col: 0 })
    } catch {
      continue
    }
    definedNames.push(
      `<definedName name="${x(n.name)}"${n.scope >= 0 ? ` localSheetId="${n.scope}"` : ""}${n.hidden ? ` hidden="1"` : ""}>${x(body)}</definedName>`,
    )
  }
  if (lost)
    warnings.push(
      `${lost === 1 ? "A formula" : `${lost} formulas`} couldn't be read; ${lost === 1 ? "its cell keeps its" : "their cells keep their"} value.`,
    )

  // The workbook and its parts.
  const macros = wbRels.some((r) => kind(r) === "vbaProject")
  const wbName = xmlName(mainPart)
  put(
    wbName,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}">` +
      (book.date1904 ? `<workbookPr date1904="1"/>` : "") +
      (book.bookView ? `<bookViews>${book.bookView}</bookViews>` : "") +
      `<sheets>${sheetsXml.join("")}</sheets>` +
      (definedNames.length ? `<definedNames>${definedNames.join("")}</definedNames>` : "") +
      book.calc +
      `</workbook>`,
    macros ? TYPES.workbookMacro : TYPES.workbook,
  )
  const wbOwn: Rel[] = []
  for (const r of wbRels) {
    const k = kind(r)
    if (k === "worksheet" || k === "chartsheet" || k === "dialogsheet" || k === "xlMacrosheet") {
      const k2 = kept.find((s) => s.id === r.id)
      if (k2) wbOwn.push(k2)
    } else if (k === "styles") {
      const b = get(r.target)
      if (!b) continue
      const name = xmlName(r.target)
      put(name, stylesXml(b), TYPES.styles)
      wbOwn.push({ ...r, target: name })
    } else if (k === "sharedStrings") {
      const name = xmlName(r.target)
      put(
        name,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si>${tText(s)}</si>`).join("")}</sst>`,
        TYPES.sharedStrings,
      )
      wbOwn.push({ ...r, target: name })
    } else if (/\.bin$/i.test(r.target) && k !== "vbaProject") {
      if (k === "pivotCacheDefinition")
        warnings.push("The workbook has PivotTables, which Cells doesn't open from .xlsb files yet.")
      else if (k === "externalLink")
        warnings.push("The workbook has links to other workbooks, which Cells doesn't open from .xlsb files yet.")
      else if (k === "connections")
        warnings.push("The workbook has data connections, which Cells doesn't open from .xlsb files yet.")
    } else wbOwn.push(r)
  }
  put(relsOf(wbName), relsXml(wbName, wbOwn))
  put(
    "_rels/.rels",
    rootRels.replace(RELATIONSHIPS, (tag) =>
      attr(tag, "Type")?.endsWith("/officeDocument") ? tag.replace(/Target="[^"]*"/, `Target="${x(wbName)}"`) : tag,
    ),
  )

  // Everything else that's XML or media (theme, drawings, charts, pictures, VML, document properties) as it is.
  for (const [name, data] of files) {
    if (name === "[Content_Types].xml" || out.has(name) || /(^|\/)_rels\//.test(name)) continue
    if (/\.bin$/i.test(name) && !/vbaProject\.bin$|printerSettings\d*\.bin$/i.test(name)) continue
    out.set(name, data)
  }
  for (const [name, data] of files) {
    if (!/(^|\/)_rels\/[^/]+\.rels$/.test(name) || out.has(name)) continue
    const owner = name.replace(/_rels\/([^/]+)\.rels$/, "$1")
    if (/\.bin$/i.test(owner) && !/vbaProject\.bin$/i.test(owner)) continue
    out.set(name, data)
  }

  // Content types: the original's for parts kept, the XML ones for parts written.
  const ct = text("[Content_Types].xml") ?? ""
  const defaults = [...ct.matchAll(/<Default\b[^>]*\/>/g)]
    .map((m) => m[0])
    .filter((d) => attr(d, "Extension")?.toLowerCase() !== "bin")
  const overrides = [...ct.matchAll(/<Override\b[^>]*\/>/g)]
    .map((m) => m[0])
    .filter((o) => {
      const part = (attr(o, "PartName") ?? "").replace(/^\//, "")
      return out.has(part) && !types.has(part)
    })
  const binKept = [...out.keys()].filter((n) => /\.bin$/i.test(n))
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    defaults.join("") +
    (binKept.some((n) => /printerSettings/i.test(n))
      ? `<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>`
      : "") +
    binKept
      .filter((n) => /vbaProject\.bin$/i.test(n))
      .map((n) => `<Override PartName="/${x(n)}" ContentType="application/vnd.ms-office.vbaProject"/>`)
      .join("") +
    overrides.join("") +
    [...types].map(([name, type]) => `<Override PartName="/${x(name)}" ContentType="${type}"/>`).join("") +
    `</Types>`
  out.set("[Content_Types].xml", encoder.encode(contentTypes))

  const order = ["[Content_Types].xml", ...[...out.keys()].filter((n) => n !== "[Content_Types].xml")]
  return { bytes: storedZip(order.map((name) => ({ name, data: out.get(name)! }))), warnings, macros }
}

function relsXml(part: string, list: { id: string; type: string; target: string; external: boolean }[]): string {
  const from = part.split("/").slice(0, -1)
  const relative = (target: string) => {
    const to = target.split("/")
    let i = 0
    while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++
    return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/")
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    list
      .map(
        (r) =>
          `<Relationship Id="${x(r.id)}" Type="${x(r.type)}" Target="${x(r.external ? r.target : relative(r.target))}"${r.external ? ` TargetMode="External"` : ""}/>`,
      )
      .join("") +
    `</Relationships>`
  )
}

export type XlsbRead = XlsxRead

/** Opens an .xlsb workbook (as an .xlsx made from it). */
export async function openXlsb(
  bytes: Uint8Array,
  fileName = "Book1.xlsb",
  options: OpenOptions = {},
): Promise<XlsbRead & { macros: boolean }> {
  const converted = await xlsbToXlsx(bytes)
  const name = fileName.replace(/\.xlsb$/i, "") + (converted.macros ? ".xlsm" : ".xlsx")
  const read = await openXlsx(converted.bytes, name, options)
  return { ...read, warnings: [...new Set([...converted.warnings, ...read.warnings])], macros: converted.macros }
}
