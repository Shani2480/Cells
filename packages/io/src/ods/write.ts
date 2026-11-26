import {
  displayText,
  isError,
  type Border,
  type CellStyle,
  type CellValue,
  type Range,
  type Workbook,
} from "@cells/core"
import { escapeXml, writeZip } from "@cells/ooxml"
import type { SheetView } from "../xlsx/sheet-read"
import { DEFAULT_FONT } from "../xlsx/styles"
import { formulaToOds, type OdsFormulaContext } from "./formula"
import { formatToOds } from "./numfmt"

/**
 * Saves a workbook as an OpenDocument Spreadsheet (.ods) that LibreOffice, Excel and Google Sheets open: every
 * sheet's values, formulas (in OpenFormula), cell formats and number formats, column widths, row heights, hidden
 * rows, columns and sheets, merged cells, notes, links, named ranges, print areas, frozen panes and the active
 * sheet. Charts, pictures, shapes, conditional formats, validation, tables (their ranges stay, as plain cells) and
 * PivotTables aren't part of it: Save asks first, as Excel does ("Some features might be lost").
 */

export type OdsSave = {
  views?: SheetView[]
  activeSheet?: number
}

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ` +
  `xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" ` +
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:tableooo="http://openoffice.org/2009/table" ` +
  `xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"`

const x = (s: string) => escapeXml(s)
const MIME = "application/vnd.oasis.opendocument.spreadsheet"

/** Text as ODF paragraphs: runs of spaces as text:s, tabs as text:tab, one text:p a line. */
export function odsParagraphs(text: string, link?: string): string {
  return text
    .split("\n")
    .map((line) => {
      let body = ""
      let i = 0
      while (i < line.length) {
        const ch = line[i]
        if (ch === "\t") {
          body += "<text:tab/>"
          i++
        } else if (ch === " ") {
          let n = 0
          while (line[i + n] === " ") n++
          // One space stands as itself between words; the rest (and a leading one) are counted.
          const lead = i === 0 ? 0 : 1
          body += " ".repeat(lead) + (n - lead > 0 ? `<text:s${n - lead > 1 ? ` text:c="${n - lead}"` : ""}/>` : "")
          i += n
        } else {
          let j = i
          while (j < line.length && line[j] !== " " && line[j] !== "\t") j++
          body += x(line.slice(i, j))
          i = j
        }
      }
      return `<text:p>${link ? `<text:a xlink:href="${x(link)}" xlink:type="simple">${body}</text:a>` : body}</text:p>`
    })
    .join("")
}

const BORDER: Record<string, string> = {
  thin: "0.75pt solid",
  medium: "1.75pt solid",
  thick: "2.5pt solid",
  hair: "0.25pt solid",
  dotted: "0.75pt dotted",
  dashed: "0.75pt dashed",
  mediumDashed: "1.75pt dashed",
  double: "2.5pt double",
  dashDot: "0.75pt dash-dot",
  mediumDashDot: "1.75pt dash-dot",
  slantDashDot: "1.75pt dash-dot",
  dashDotDot: "0.75pt dash-dot-dot",
  mediumDashDotDot: "1.75pt dash-dot-dot",
}

const border = (b: Border | undefined) => (b ? `${BORDER[b.style] ?? "0.75pt solid"} ${b.color ?? "#000000"}` : "none")

const H_ALIGN: Record<string, string> = {
  left: "start",
  right: "end",
  center: "center",
  centerContinuous: "center",
  justify: "justify",
  distributed: "justify",
  fill: "start",
}

/** A Cells format as an ODF cell style's properties. */
function odsCellStyle(
  name: string,
  s: CellStyle,
  dataStyle: string | null,
  fontFace: (name: string) => string,
): string {
  let cell = ""
  if (s.fill) cell += ` fo:background-color="${s.fill}"`
  const sides = [s.borderTop, s.borderBottom, s.borderLeft, s.borderRight]
  if (sides.some(Boolean)) {
    cell += ` fo:border-top="${border(s.borderTop)}" fo:border-bottom="${border(s.borderBottom)}" fo:border-left="${border(s.borderLeft)}" fo:border-right="${border(s.borderRight)}"`
  }
  if (s.borderDiagonal && s.diagonalUp) cell += ` style:diagonal-bl-tr="${border(s.borderDiagonal)}"`
  if (s.borderDiagonal && s.diagonalDown) cell += ` style:diagonal-tl-br="${border(s.borderDiagonal)}"`
  if (s.vAlign === "top") cell += ` style:vertical-align="top"`
  else if (s.vAlign === "center") cell += ` style:vertical-align="middle"`
  else if (s.vAlign) cell += ` style:vertical-align="bottom"`
  if (s.wrap) cell += ` fo:wrap-option="wrap"`
  if (s.shrink) cell += ` style:shrink-to-fit="true"`
  if (s.rotation === 255) cell += ` style:direction="ttb"`
  else if (s.rotation) cell += ` style:rotation-angle="${s.rotation <= 90 ? s.rotation : 360 - (s.rotation - 90)}"`
  if (s.hAlign && H_ALIGN[s.hAlign])
    cell += ` style:text-align-source="fix" style:repeat-content="${s.hAlign === "fill"}"`
  if (s.locked === false || s.hiddenFormula)
    cell += ` style:cell-protect="${s.locked === false ? (s.hiddenFormula ? "formula-hidden" : "none") : "hidden-and-protected"}"`
  let para = ""
  if (s.hAlign && H_ALIGN[s.hAlign]) para += ` fo:text-align="${H_ALIGN[s.hAlign]}"`
  if (s.indent) para += ` fo:margin-left="${s.indent * 10}pt"`
  let text = ""
  if (s.fontName) text += ` style:font-name="${x(fontFace(s.fontName))}"`
  if (s.fontSize) text += ` fo:font-size="${s.fontSize}pt"`
  if (s.bold) text += ` fo:font-weight="bold"`
  if (s.italic) text += ` fo:font-style="italic"`
  if (s.underline)
    text += ` style:text-underline-style="solid" style:text-underline-type="${/double/i.test(s.underline) ? "double" : "single"}" style:text-underline-width="auto" style:text-underline-color="font-color"`
  if (s.strike) text += ` style:text-line-through-style="solid" style:text-line-through-type="single"`
  if (s.fontColor) text += ` fo:color="${s.fontColor}"`
  if (s.vertAlign === "superscript") text += ` style:text-position="super 58%"`
  if (s.vertAlign === "subscript") text += ` style:text-position="sub 58%"`
  return (
    `<style:style style:name="${name}" style:family="table-cell" style:parent-style-name="Default"${dataStyle ? ` style:data-style-name="${dataStyle}"` : ""}>` +
    (cell ? `<style:table-cell-properties${cell}/>` : "") +
    (para ? `<style:paragraph-properties${para}/>` : "") +
    (text ? `<style:text-properties${text}/>` : "") +
    `</style:style>`
  )
}

/** A serial date as ODF's ISO date-time (1899-12-30 is day 0). */
function isoDate(serial: number): string {
  const ms = Math.round((serial - 25569) * 86400000)
  const d = new Date(ms)
  const iso = d.toISOString()
  return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 19) : iso.replace(/\.000Z$/, "").replace(/Z$/, "")
}

/** A day fraction as an ODF duration (PT12H30M00S). */
function isoTime(value: number): string {
  const sign = value < 0 ? "-" : ""
  let secs = Math.round(Math.abs(value) * 86400 * 1000) / 1000
  const h = Math.floor(secs / 3600)
  secs -= h * 3600
  const m = Math.floor(secs / 60)
  secs -= m * 60
  const s = secs.toFixed(Number.isInteger(secs) ? 0 : 3).padStart(2, "0")
  return `${sign}PT${String(h).padStart(2, "0")}H${String(m).padStart(2, "0")}M${s}S`
}

const errorText = (v: CellValue) => (isError(v) ? v.error : "")

/** The sheet name as an ODF cell address writes it. */
const plainSheet = (name: string) => (/^[A-Za-z_][\w]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
/** Absolute ($Sheet1), as LibreOffice writes addresses: it drops names whose addresses aren't. */
const sheetRef = (name: string) => "$" + plainSheet(name)

function colLetters(c: number): string {
  let s = ""
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}

const address = (sheet: string, r: Range) =>
  `${sheetRef(sheet)}.$${colLetters(r.c0)}$${r.r0 + 1}` +
  (r.r0 === r.r1 && r.c0 === r.c1 ? "" : `:${sheetRef(sheet)}.$${colLetters(r.c1)}$${r.r1 + 1}`)

export async function saveOds(wb: Workbook, options: OdsSave = {}): Promise<Uint8Array> {
  const defaultFont = { name: wb.theme.fonts.minor || DEFAULT_FONT.name, size: DEFAULT_FONT.size }
  const faces = new Set<string>([defaultFont.name])
  const fontFace = (name: string) => {
    faces.add(name)
    return name
  }
  const tables = wb.sheets.flatMap((s) => s.tables.map((table) => ({ table, sheet: s.name })))

  // Styles: one cell style per format used, one data style per number format.
  const dataStyles = new Map<string, { name: string; xml: string; type: ReturnType<typeof formatToOds>["type"] }>()
  const dataStyle = (code: string | undefined) => {
    if (!code || code === "General") return null
    let d = dataStyles.get(code)
    if (!d) {
      const name = `N${dataStyles.size + 100}`
      const f = formatToOds(code, name)
      d = { name, xml: f.xml, type: f.type }
      dataStyles.set(code, d)
    }
    return d
  }
  const cellStyles = new Map<number, string>()
  const cellStyleList: string[] = []
  const styleName = (id: number | undefined): string | null => {
    if (!id) return null
    let name = cellStyles.get(id)
    if (name !== undefined) return name
    const s = wb.styles.get(id)
    if (!Object.keys(s).length) return null
    name = `ce${cellStyles.size + 1}`
    cellStyles.set(id, name)
    const d = dataStyle(s.numberFormat)
    cellStyleList.push(odsCellStyle(name, s, d?.name ?? null, fontFace))
    return name
  }

  // TRUE and FALSE are numbers shown with a boolean data style in LibreOffice: their cells' styles carry one.
  const booleanStyles = new Map<number, string>()
  const booleanStyle = (id: number | undefined): string => {
    const key = id ?? 0
    let name = booleanStyles.get(key)
    if (!name) {
      name = `ceb${booleanStyles.size + 1}`
      booleanStyles.set(key, name)
      const s = key ? wb.styles.get(key) : {}
      cellStyleList.push(odsCellStyle(name, { ...s, numberFormat: undefined }, "NB", fontFace))
    }
    return name
  }
  const columnStyles = new Map<string, string>()
  const rowStyles = new Map<string, string>()
  const autoXml: string[] = []
  const columnStyle = (px: number) => {
    const key = String(px)
    let name = columnStyles.get(key)
    if (!name) {
      name = `co${columnStyles.size + 1}`
      columnStyles.set(key, name)
      autoXml.push(
        `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="${Math.round((px / 96) * 10000) / 10000}in"/></style:style>`,
      )
    }
    return name
  }
  const rowStyle = (px: number, custom: boolean) => {
    const key = `${px},${custom}`
    let name = rowStyles.get(key)
    if (!name) {
      name = `ro${rowStyles.size + 1}`
      rowStyles.set(key, name)
      autoXml.push(
        `<style:style style:name="${name}" style:family="table-row"><style:table-row-properties style:row-height="${Math.round(px * 0.75 * 100) / 100}pt" fo:break-before="auto" style:use-optimal-row-height="${!custom}"/></style:style>`,
      )
    }
    return name
  }
  const tableStyles: string[] = []
  const tableStyle = (hidden: boolean, tab?: string) => {
    const name = `ta${tableStyles.length + 1}`
    tableStyles.push(
      `<style:style style:name="${name}" style:family="table" style:master-page-name="Default"><style:table-properties table:display="${!hidden}" style:writing-mode="lr-tb"${tab ? ` tableooo:tab-color="${tab}"` : ""}/></style:style>`,
    )
    return name
  }

  const spills = new Map(wb.calc.spillAreas().map((a) => [`${a.sheet},${a.row},${a.col}`, a.range]))
  const sheetsXml: string[] = []
  for (const [index, sheet] of wb.sheets.entries()) {
    const used = sheet.cells.usedBounds()
    const merges = sheet.merges.all
    const notes = new Map(sheet.annotations.all().map((a) => [`${a.row},${a.col}`, a.value]))
    let lastRow = used?.r1 ?? -1
    let lastCol = used?.c1 ?? -1
    for (const m of merges) {
      lastRow = Math.max(lastRow, m.r1)
      lastCol = Math.max(lastCol, m.c1)
    }
    for (const n of sheet.annotations.all()) {
      lastRow = Math.max(lastRow, n.row)
      lastCol = Math.max(lastCol, n.col)
    }
    for (let c = 0; c < 16384 && c < lastCol + 64; c++)
      if (sheet.cols.hasCustomSize(c) || sheet.cols.isHidden(c)) lastCol = Math.max(lastCol, c)
    for (let r = 0; r < lastRow + 256; r++)
      if (sheet.rows.hasCustomSize(r) || sheet.rows.isHidden(r)) lastRow = Math.max(lastRow, r)

    const tab = sheet.tabColor
    let xml = `<table:table table:name="${x(sheet.name)}" table:style-name="${tableStyle(sheet.hidden, tab)}"`
    const print = sheet.pageSetup.printArea
    if (print?.length) xml += ` table:print-ranges="${x(print.map((r) => address(sheet.name, r)).join(" "))}"`
    xml += ">"

    // Columns, in runs of the same width and visibility; then the rest of the sheet's.
    const defaultCol = sheet.cols.defaultSize
    let c = 0
    while (c <= lastCol) {
      const width = sheet.cols.hasCustomSize(c) ? sheet.cols.storedSize(c) : defaultCol
      const hidden = sheet.cols.isHidden(c)
      let n = 1
      while (
        c + n <= lastCol &&
        (sheet.cols.hasCustomSize(c + n) ? sheet.cols.storedSize(c + n) : defaultCol) === width &&
        sheet.cols.isHidden(c + n) === hidden
      )
        n++
      xml += `<table:table-column table:style-name="${columnStyle(width)}"${n > 1 ? ` table:number-columns-repeated="${n}"` : ""}${hidden ? ` table:visibility="collapse"` : ""} table:default-cell-style-name="Default"/>`
      c += n
    }
    xml += `<table:table-column table:style-name="${columnStyle(defaultCol)}" table:number-columns-repeated="${16384 - Math.max(0, lastCol + 1)}" table:default-cell-style-name="Default"/>`

    // Rows: cells in runs; empty rows of the same height in runs.
    const covered = new Set<string>()
    for (const m of merges)
      for (let r = m.r0; r <= m.r1; r++)
        for (let k = m.c0; k <= m.c1; k++) if (r !== m.r0 || k !== m.c0) covered.add(`${r},${k}`)
    const links = sheet.hyperlinks
    const linkAt = (r: number, k: number) =>
      links.find((l) => r >= l.range.r0 && r <= l.range.r1 && k >= l.range.c0 && k <= l.range.c1)
    const defaultRow = sheet.rows.defaultSize
    const ctx: OdsFormulaContext = { tables, row: 0, sheet: sheet.name }
    let pendingEmpty: { style: string; hidden: boolean; count: number } | null = null
    const flushEmpty = () => {
      if (!pendingEmpty) return
      xml += `<table:table-row table:style-name="${pendingEmpty.style}"${pendingEmpty.count > 1 ? ` table:number-rows-repeated="${pendingEmpty.count}"` : ""}${pendingEmpty.hidden ? ` table:visibility="collapse"` : ""}><table:table-cell table:number-columns-repeated="16384"/></table:table-row>`
      pendingEmpty = null
    }
    for (let r = 0; r <= lastRow; r++) {
      const custom = sheet.rows.hasCustomSize(r) && !sheet.rows.isAuto(r)
      const height = sheet.rows.hasCustomSize(r) ? sheet.rows.storedSize(r) : defaultRow
      const rs = rowStyle(height, custom)
      const hidden = sheet.rows.isHidden(r)
      const cells: string[] = []
      let empty = 0
      let any = false
      const flushCells = () => {
        if (empty) cells.push(`<table:table-cell${empty > 1 ? ` table:number-columns-repeated="${empty}"` : ""}/>`)
        empty = 0
      }
      for (let k = 0; k <= lastCol; k++) {
        const data = sheet.cells.get(r, k)
        const note = notes.get(`${r},${k}`)
        const isCovered = covered.has(`${r},${k}`)
        const style = styleName(data?.style)
        if (
          !data?.value &&
          data?.value !== 0 &&
          data?.value !== false &&
          !data?.input &&
          !note &&
          !style &&
          !isCovered &&
          !merges.some((m) => m.r0 === r && m.c0 === k)
        ) {
          empty++
          continue
        }
        flushCells()
        any = true
        if (isCovered) {
          cells.push(`<table:covered-table-cell${style ? ` table:style-name="${style}"` : ""}/>`)
          continue
        }
        cells.push(
          cellXml(
            r,
            k,
            data?.value,
            data?.input,
            style,
            data?.style,
            note,
            merges.find((m) => m.r0 === r && m.c0 === k),
            linkAt(r, k),
          ),
        )
      }
      flushCells()
      if (!any) {
        const key = { style: rs, hidden, count: 1 }
        if (pendingEmpty && pendingEmpty.style === rs && pendingEmpty.hidden === hidden) pendingEmpty.count++
        else {
          flushEmpty()
          pendingEmpty = key
        }
        continue
      }
      flushEmpty()
      xml += `<table:table-row table:style-name="${rs}"${hidden ? ` table:visibility="collapse"` : ""}>${cells.join("")}<table:table-cell table:number-columns-repeated="${16384 - lastCol - 1}"/></table:table-row>`
    }
    flushEmpty()
    xml += `<table:table-row table:style-name="${rowStyle(defaultRow, false)}" table:number-rows-repeated="${1048576 - lastRow - 1}"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>`
    xml += `</table:table>`
    sheetsXml.push(xml)

    function cellXml(
      r: number,
      k: number,
      value: CellValue | undefined,
      input: string | undefined,
      style: string | null,
      styleId: number | undefined,
      note: ReturnType<typeof notes.get>,
      merge: Range | undefined,
      link: (typeof links)[number] | undefined,
    ): string {
      const fmt = styleId ? wb.styles.get(styleId).numberFormat : undefined
      let attrs = style ? ` table:style-name="${style}"` : ""
      if (merge)
        attrs += ` table:number-columns-spanned="${merge.c1 - merge.c0 + 1}" table:number-rows-spanned="${merge.r1 - merge.r0 + 1}"`
      if (input?.startsWith("=")) {
        ctx.row = r
        const f = formulaToOds(input, ctx)
        if (f) attrs += ` table:formula="${x(f)}"`
        // An array formula's range (a spill's, or a legacy one's).
        const spill = spills.get(`${index},${r},${k}`)
        if (f && spill && (spill.r1 > spill.r0 || spill.c1 > spill.c0))
          attrs += ` table:number-matrix-columns-spanned="${spill.c1 - spill.c0 + 1}" table:number-matrix-rows-spanned="${spill.r1 - spill.r0 + 1}"`
      }
      let shown = ""
      if (typeof value === "number") {
        const type = dataStyle(fmt)?.type ?? "float"
        if (type === "date")
          attrs += ` office:value-type="date" office:date-value="${isoDate(value)}" calcext:value-type="date"`
        else if (type === "time")
          attrs += ` office:value-type="time" office:time-value="${isoTime(value)}" calcext:value-type="time"`
        else if (type === "percentage")
          attrs += ` office:value-type="percentage" office:value="${value}" calcext:value-type="percentage"`
        else attrs += ` office:value-type="float" office:value="${value}" calcext:value-type="float"`
        shown = displayText(value, fmt)
      } else if (typeof value === "boolean") {
        attrs = attrs.replace(/^ table:style-name="[^"]*"|^/, ` table:style-name="${booleanStyle(styleId)}"`)
        attrs += ` office:value-type="boolean" office:boolean-value="${value}" calcext:value-type="boolean"`
        shown = value ? "TRUE" : "FALSE"
      } else if (typeof value === "string") {
        attrs += ` office:value-type="string" calcext:value-type="string"`
        shown = value
      } else if (value !== undefined && isError(value)) {
        attrs += ` office:value-type="string" office:string-value="" calcext:value-type="error"`
        shown = errorText(value)
      }
      let inner = ""
      if (note) {
        const text =
          note.kind === "note"
            ? note.text
            : note.posts.map((p) => (p.author ? `${p.author}: ${p.text}` : p.text)).join("\n")
        const author = note.kind === "note" ? note.author : (note.posts[0]?.author ?? "")
        inner += `<office:annotation${note.kind === "note" && note.shown ? ` office:display="true"` : ""}>${author ? `<dc:creator>${x(author)}</dc:creator>` : ""}${odsParagraphs(text)}</office:annotation>`
      }
      const href = link
        ? link.address
          ? link.address + (link.location ? `#${link.location}` : "")
          : link.location
            ? `#${odsLocation(link.location)}`
            : undefined
        : undefined
      if (value !== undefined && value !== "") inner += odsParagraphs(shown, href)
      return `<table:table-cell${attrs}>${inner}</table:table-cell>`
    }
  }

  // Names: ranges as named ranges, the rest as named expressions.
  const names = wb.names.all().filter((n) => !/^_xlnm\./i.test(n.name))
  let namesXml = ""
  for (const n of names) {
    const f = formulaToOds("=" + n.ref, { tables, row: 0, sheet: wb.sheets[0]?.name ?? "" })
    if (!f) continue
    const body = f.slice(4)
    const range = /^\[([^\]]+)\]$/.exec(body)
    namesXml += range
      ? `<table:named-range table:name="${x(n.name)}" table:base-cell-address="${x(`${sheetRef(wb.sheets[0]?.name ?? "Sheet1")}.$A$1`)}" table:cell-range-address="${x(odsAddress(range[1]))}"/>`
      : `<table:named-expression table:name="${x(n.name)}" table:base-cell-address="${x(`${sheetRef(wb.sheets[0]?.name ?? "Sheet1")}.$A$1`)}" table:expression="${x(f)}"/>`
  }

  const faceXml = `<office:font-face-decls>${[...faces].map((f) => `<style:font-face style:name="${x(f)}" svg:font-family="${x(/\s/.test(f) ? `'${f}'` : f)}"/>`).join("")}</office:font-face-decls>`
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS} office:version="1.3">` +
    faceXml +
    `<office:automatic-styles>${autoXml.join("")}${tableStyles.join("")}${booleanStyles.size ? `<number:boolean-style style:name="NB"><number:boolean/></number:boolean-style>` : ""}${[...dataStyles.values()].map((d) => d.xml).join("")}${cellStyleList.join("")}</office:automatic-styles>` +
    `<office:body><office:spreadsheet><table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/>` +
    sheetsXml.join("") +
    (namesXml ? `<table:named-expressions>${namesXml}</table:named-expressions>` : "") +
    `</office:spreadsheet></office:body></office:document-content>`

  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS} office:version="1.3">` +
    faceXml +
    `<office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="${x(defaultFont.name)}" fo:font-size="${defaultFont.size}pt"/></style:default-style>` +
    `<number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>` +
    `<style:style style:name="Default" style:family="table-cell" style:data-style-name="N0"><style:text-properties style:font-name="${x(defaultFont.name)}" fo:font-size="${defaultFont.size}pt"/></style:style>` +
    `</office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:margin-top="0.75in" fo:margin-bottom="0.75in" fo:margin-left="0.7in" fo:margin-right="0.7in"/></style:page-layout></office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>`

  const meta =
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ${NS} office:version="1.3"><office:meta><meta:generator>Cells</meta:generator>` +
    `<meta:document-statistic meta:table-count="${wb.sheets.length}"/></office:meta></office:document-meta>`

  const settings = settingsXml(wb, options)
  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${MIME}"/>` +
    ["content.xml", "styles.xml", "meta.xml", "settings.xml"]
      .map((p) => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="text/xml"/>`)
      .join("") +
    `</manifest:manifest>`

  const enc = new TextEncoder()
  return writeZip([
    // The media type comes first, stored, so the file is known by its first bytes.
    { name: "mimetype", data: enc.encode(MIME), stored: true },
    { name: "content.xml", data: enc.encode(content), stored: false },
    { name: "styles.xml", data: enc.encode(styles), stored: false },
    { name: "meta.xml", data: enc.encode(meta), stored: false },
    { name: "settings.xml", data: enc.encode(settings), stored: false },
    { name: "META-INF/manifest.xml", data: enc.encode(manifest), stored: false },
  ])
}

/** A link's place in this workbook ("Sheet2!A1") as ODF writes it ("Sheet2.A1"). */
function odsLocation(location: string): string {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(location)
  if (!m) return location
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]
  return `${plainSheet(sheet)}.${m[3].replace(/\$/g, "")}`
}

/** [Sheet.$A$1:.$B$2] (a formula's reference) as a cell-range-address (Sheet.$A$1:Sheet.$B$2). */
function odsAddress(inner: string): string {
  const [a, b] = inner.split(":")
  const sheet = a.slice(0, a.lastIndexOf("."))
  return b === undefined ? a : `${a}:${b.startsWith(".") ? sheet + b : b}`
}

function settingsXml(wb: Workbook, options: OdsSave): string {
  const item = (name: string, type: string, value: string | number | boolean) =>
    `<config:config-item config:name="${name}" config:type="${type}">${value}</config:config-item>`
  const active = wb.sheets[options.activeSheet ?? 0]?.name ?? wb.sheets[0]?.name ?? ""
  const tables = wb.sheets
    .map((s, i) => {
      const v = options.views?.[i]
      const f = v?.freeze
      const cursor = v?.active ?? { row: 0, col: 0 }
      return (
        `<config:config-item-map-entry config:name="${x(s.name)}">` +
        item("CursorPositionX", "int", cursor.col) +
        item("CursorPositionY", "int", cursor.row) +
        item("HorizontalSplitMode", "short", f && f.col > 0 ? 2 : 0) +
        item("VerticalSplitMode", "short", f && f.row > 0 ? 2 : 0) +
        item("HorizontalSplitPosition", "int", f?.col ?? 0) +
        item("VerticalSplitPosition", "int", f?.row ?? 0) +
        item("ActiveSplitRange", "short", f && f.row > 0 ? 2 : 3) +
        item("PositionLeft", "int", 0) +
        item("PositionRight", "int", f?.left ?? f?.col ?? 0) +
        item("PositionTop", "int", 0) +
        item("PositionBottom", "int", f?.top ?? f?.row ?? 0) +
        item("ZoomType", "short", 0) +
        item("ZoomValue", "int", Math.round((v?.zoom ?? 1) * 100)) +
        item("ShowGrid", "boolean", v?.showGridlines !== false) +
        `</config:config-item-map-entry>`
      )
    })
    .join("")
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-settings ${NS} office:version="1.3"><office:settings>` +
    `<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
    item("ViewId", "string", "view1") +
    `<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>` +
    item("ActiveTable", "string", x(active)) +
    item("ShowGrid", "boolean", true) +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`
  )
}
