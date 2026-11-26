import { columnName } from "@cells/core"
import { escapeXml, readZip, storedZip, type ZipEntry } from "@cells/ooxml"
import { at, kids, parseTree, type XNode } from "../xlsx/dom"
import { openXlsx, type OpenOptions, type XlsxRead } from "../xlsx/read"
import { formulaToFile } from "../xlsx/formula"
import { pxToColWidth } from "../xlsx/units"
import { formulaFromOds } from "./formula"
import { formatFromOds } from "./numfmt"

/**
 * OpenDocument Spreadsheets (.ods, as LibreOffice, OpenOffice, Google Sheets and Excel write them). Cells writes the
 * document out as the .xlsx it would be (values, formulas, cell styles and number formats, column widths and row
 * heights, hidden rows, columns and sheets, merges, notes, links, named ranges, print ranges, frozen panes) and
 * opens that, so everything the .xlsx reader knows applies. Charts, pictures, conditional formats, validation and
 * pivot tables (DataPilot) aren't brought over yet; a warning says so.
 */

export class OdsError extends Error {
  constructor(
    message: string,
    readonly kind: "encrypted" | "damaged",
  ) {
    super(message)
    this.name = "OdsError"
  }
}

const NS = {
  office: "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
  style: "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
  table: "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
  text: "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
  fo: "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
  number: "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
  svg: "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
  xlink: "http://www.w3.org/1999/xlink",
  dc: "http://purl.org/dc/elements/1.1/",
  config: "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
  calcext: "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
  tableooo: "http://openoffice.org/2009/table",
}

const x = (s: string) => escapeXml(s)
const cellRef = (row: number, col: number) => `${columnName(col)}${row + 1}`
const rangeRef = (r0: number, c0: number, r1: number, c1: number) =>
  r0 === r1 && c0 === c1 ? cellRef(r0, c0) : `${cellRef(r0, c0)}:${cellRef(r1, c1)}`

const MAX_ROWS = 1048576
const MAX_COLS = 16384
/** Empty formatted cells and rows repeated more than this are left out (a style run to the sheet's end). */
const REPEAT_LIMIT = 1000

/** A length ("2.35cm", "0.889in", "18.75pt", "64px") in points. */
function points(length: string | undefined): number | null {
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px)?$/.exec(length?.trim() ?? "")
  if (!m) return null
  const v = Number(m[1])
  switch (m[2]) {
    case "cm":
      return (v / 2.54) * 72
    case "mm":
      return (v / 25.4) * 72
    case "in":
      return v * 72
    case "pc":
      return v * 12
    case "px":
      return v * 0.75
    default:
      return v
  }
}

/** A cell style's properties, merged along its parents. */
type Props = Record<string, string>

function propsOf(style: XNode): Props {
  const out: Props = {}
  for (const c of style.children) {
    if (!/-properties$/.test(c.local)) continue
    for (const [k, v] of Object.entries(c.nsAttrs)) {
      const [uri, local] = k.includes("|") ? k.split("|") : ["", k]
      const prefix = Object.entries(NS).find(([, u]) => u === uri)?.[0] ?? ""
      out[`${prefix}:${local}`] = v
    }
  }
  const data = at(style, "data-style-name", NS.style)
  if (data) out["data-style"] = data
  return out
}

const BORDER_STYLE = (spec: string | undefined): { style: string; color: string } | null => {
  if (!spec || spec === "none" || /^\s*0(\.0+)?\w*\s/.test(spec) || /hidden/.test(spec)) return null
  const color = /#([0-9a-f]{6})/i.exec(spec)?.[1]?.toUpperCase() ?? "000000"
  const width = /([\d.]+(?:pt|cm|mm|in|px))/.exec(spec)?.[1]
  const w = width ? (points(width) ?? 0.75) : /thick/.test(spec) ? 2.5 : /medium/.test(spec) ? 1.75 : 0.75
  let style: string
  if (/double/.test(spec)) style = "double"
  else if (/dash-dot-dot/.test(spec)) style = w > 1.1 ? "mediumDashDotDot" : "dashDotDot"
  else if (/dash-dot/.test(spec)) style = w > 1.1 ? "mediumDashDot" : "dashDot"
  else if (/dashed|fine-dashed|long-dash/.test(spec)) style = w > 1.1 ? "mediumDashed" : "dashed"
  else if (/dotted/.test(spec)) style = w < 0.3 ? "hair" : "dotted"
  else style = w > 2 ? "thick" : w > 1.1 ? "medium" : w < 0.3 ? "hair" : "thin"
  return { style, color }
}

const fontFamily = (s: string) => s.replace(/^['"]|['"]$/g, "")

type Font = {
  name: string
  size: number
  bold: boolean
  italic: boolean
  underline: string
  strike: boolean
  color: string
  script: string
}

const fontKey = (f: Font) => JSON.stringify(f)

/** The styles of the workbook as styles.xml, and the index of each ODF cell style. */
class Styles {
  readonly numFmts: string[] = []
  private readonly formatIds = new Map<string, number>()
  readonly fonts: string[] = []
  private readonly fontIds = new Map<string, number>()
  readonly fills = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
  ]
  private readonly fillIds = new Map<string, number>()
  readonly borders = [`<border><left/><right/><top/><bottom/><diagonal/></border>`]
  private readonly borderIds = new Map<string, number>()
  readonly xfs: string[] = []
  private readonly xfIds = new Map<string, number>()
  private readonly cellStyles = new Map<string, XNode>()
  private readonly dataStyles = new Map<string, XNode>()
  private readonly resolved = new Map<string, number>()
  private readonly fontFaces = new Map<string, string>()
  readonly base: Font
  private readonly defaultProps: Props

  constructor(docs: XNode[]) {
    for (const doc of docs) {
      for (const section of doc.children) {
        if (section.local === "font-face-decls")
          for (const f of kids(section, "font-face"))
            this.fontFaces.set(
              at(f, "name", NS.style) ?? "",
              fontFamily(at(f, "font-family", NS.svg) ?? at(f, "name", NS.style) ?? ""),
            )
        if (section.local !== "styles" && section.local !== "automatic-styles") continue
        for (const s of section.children) {
          const name = at(s, "name", NS.style)
          if (s.local === "style" && at(s, "family", NS.style) === "table-cell" && name) this.cellStyles.set(name, s)
          else if (/-style$/.test(s.local) && s.uri === NS.number && name) this.dataStyles.set(name, s)
          else if (s.local === "default-style" && at(s, "family", NS.style) === "table-cell")
            this.cellStyles.set("\u0000default", s)
        }
      }
    }
    // The base: the default cell style, then "Default" (Excel's Normal).
    // LibreOffice's default style says "wrap" but its cells don't: only the default style's font counts.
    const fallback = this.cellStyles.get("\u0000default")
    const fontOnly = fallback
      ? { ...fallback, children: fallback.children.filter((c) => c.local === "text-properties") }
      : undefined
    const chain = [fontOnly, this.cellStyles.get("Default")].filter((s): s is XNode => !!s)
    this.defaultProps = Object.assign({}, ...chain.map(propsOf))
    this.base = this.font(this.defaultProps, null)
    this.fontId(this.base)
    this.xfs.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`)
    this.xfIds.set("", 0)
  }

  /** A cell style's merged properties. */
  private props(name: string): Props {
    const chain: XNode[] = []
    let s = this.cellStyles.get(name)
    for (let guard = 0; s && guard < 20; guard++) {
      chain.unshift(s)
      const parent = at(s, "parent-style-name", NS.style)
      s = parent ? this.cellStyles.get(parent) : undefined
    }
    return Object.assign({}, this.defaultProps, ...chain.map(propsOf))
  }

  private font(p: Props, base: Font | null): Font {
    const face = p["style:font-name"]
    const name = face
      ? (this.fontFaces.get(face) ?? face)
      : p["fo:font-family"]
        ? fontFamily(p["fo:font-family"])
        : (base?.name ?? "Liberation Sans")
    const sizePt = points(p["fo:font-size"]) ?? base?.size ?? 10
    const underline =
      p["style:text-underline-style"] && p["style:text-underline-style"] !== "none"
        ? p["style:text-underline-type"] === "double"
          ? "double"
          : "single"
        : ""
    const position = p["style:text-position"] ?? ""
    return {
      name,
      size: Math.round(sizePt * 100) / 100,
      bold: p["fo:font-weight"] === "bold" || Number(p["fo:font-weight"]) >= 600,
      italic: p["fo:font-style"] === "italic" || p["fo:font-style"] === "oblique",
      underline,
      strike: !!p["style:text-line-through-style"] && p["style:text-line-through-style"] !== "none",
      // Black is what Excel's Automatic shows; files write it out.
      color:
        /^#[0-9a-f]{6}$/i.test(p["fo:color"] ?? "") && p["fo:color"] !== "#000000"
          ? p["fo:color"].slice(1).toUpperCase()
          : "",
      script:
        /^super|^\d/.test(position) && !/^0%?\s/.test(position)
          ? "superscript"
          : /^sub|^-/.test(position)
            ? "subscript"
            : "",
    }
  }

  private fontId(f: Font): number {
    const key = fontKey(f)
    let id = this.fontIds.get(key)
    if (id !== undefined) return id
    id = this.fonts.length
    this.fontIds.set(key, id)
    this.fonts.push(
      `<font>${f.bold ? "<b/>" : ""}${f.italic ? "<i/>" : ""}${f.strike ? "<strike/>" : ""}${f.underline ? (f.underline === "double" ? `<u val="double"/>` : "<u/>") : ""}${f.script ? `<vertAlign val="${f.script}"/>` : ""}<sz val="${f.size}"/>${f.color ? `<color rgb="FF${f.color}"/>` : ""}<name val="${x(f.name)}"/></font>`,
    )
    return id
  }

  private formatId(dataStyle: string | undefined): number {
    if (!dataStyle) return 0
    const node = this.dataStyles.get(dataStyle)
    if (!node) return 0
    const code = formatFromOds(node, (n) => this.dataStyles.get(n))
    if (!code || code === "General") return 0
    let id = this.formatIds.get(code)
    if (id !== undefined) return id
    id = 164 + this.formatIds.size
    this.formatIds.set(code, id)
    this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${x(code)}"/>`)
    return id
  }

  /** The format code of a cell style (for telling dates from times from numbers). */
  formatCode(name: string | undefined): string {
    if (!name) return ""
    const data = this.props(name)["data-style"]
    const node = data ? this.dataStyles.get(data) : undefined
    return node ? formatFromOds(node, (n) => this.dataStyles.get(n)) : ""
  }

  /** The xf index of an ODF cell style (0 for none or the default). */
  xf(name: string | undefined): number {
    if (!name) return 0
    const known = this.resolved.get(name)
    if (known !== undefined) return known
    const p = this.props(name)
    const font = this.fontId(this.font(p, this.base))
    const fmt = this.formatId(p["data-style"])
    const bg = /^#[0-9a-f]{6}$/i.test(p["fo:background-color"] ?? "")
      ? p["fo:background-color"].slice(1).toUpperCase()
      : ""
    let fill = 0
    if (bg) {
      fill = this.fillIds.get(bg) ?? this.fills.length
      if (fill === this.fills.length) {
        this.fillIds.set(bg, fill)
        this.fills.push(
          `<fill><patternFill patternType="solid"><fgColor rgb="FF${bg}"/><bgColor indexed="64"/></patternFill></fill>`,
        )
      }
    }
    const side = (which: string) => BORDER_STYLE(p[`fo:border-${which}`] ?? p["fo:border"])
    const sides = { left: side("left"), right: side("right"), top: side("top"), bottom: side("bottom") }
    const up = BORDER_STYLE(p["style:diagonal-bl-tr"])
    const down = BORDER_STYLE(p["style:diagonal-tl-br"])
    const el = (tag: string, s: { style: string; color: string } | null) =>
      s
        ? `<${tag} style="${s.style}">${s.color === "000000" ? "" : `<color rgb="FF${s.color}"/>`}</${tag}>`
        : `<${tag}/>`
    const borderXml = `<border${up ? ` diagonalUp="1"` : ""}${down ? ` diagonalDown="1"` : ""}>${el("left", sides.left)}${el("right", sides.right)}${el("top", sides.top)}${el("bottom", sides.bottom)}${el("diagonal", up ?? down)}</border>`
    let border = 0
    if (borderXml !== this.borders[0]) {
      border = this.borderIds.get(borderXml) ?? this.borders.length
      if (border === this.borders.length) {
        this.borderIds.set(borderXml, border)
        this.borders.push(borderXml)
      }
    }
    // Alignment: a fixed horizontal one; vertical; wrap, shrink, indent, rotation.
    let align = ""
    const h = p["fo:text-align"]
    if (h && p["style:text-align-source"] !== "value-type") {
      const map: Record<string, string> = {
        start: "left",
        left: "left",
        end: "right",
        right: "right",
        center: "center",
        justify: "justify",
      }
      if (map[h]) align += ` horizontal="${map[h]}"`
    }
    const v = p["style:vertical-align"]
    if (v === "top") align += ` vertical="top"`
    else if (v === "middle") align += ` vertical="center"`
    if (p["fo:wrap-option"] === "wrap") align += ` wrapText="1"`
    if (p["style:shrink-to-fit"] === "true") align += ` shrinkToFit="1"`
    const margin = points(p["fo:margin-left"])
    if (margin && margin > 0 && /horizontal="left"/.test(align))
      align += ` indent="${Math.max(1, Math.round(margin / 7.5))}"`
    const angle = Number(p["style:rotation-angle"] ?? 0) % 360
    if (p["style:direction"] === "ttb") align += ` textRotation="255"`
    else if (angle > 0 && angle <= 90) align += ` textRotation="${Math.round(angle)}"`
    else if (angle >= 270) align += ` textRotation="${Math.round(90 + 360 - angle)}"`
    const protect = p["style:cell-protect"]
    const unlocked = protect === "none" || protect === "formula-hidden"
    const hidden = protect === "formula-hidden" || protect === "hidden-and-protected"
    const protection =
      unlocked || hidden ? `<protection${unlocked ? ` locked="0"` : ""}${hidden ? ` hidden="1"` : ""}/>` : ""
    const xml = `<xf numFmtId="${fmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${fmt ? ` applyNumberFormat="1"` : ""}${font ? ` applyFont="1"` : ""}${fill ? ` applyFill="1"` : ""}${border ? ` applyBorder="1"` : ""}${align ? ` applyAlignment="1"` : ""}${protection ? ` applyProtection="1"` : ""}${align || protection ? `>${align ? `<alignment${align}/>` : ""}${protection}</xf>` : "/>"}`
    let id = this.xfIds.get(xml)
    if (id === undefined) {
      id = fmt || font || fill || border || align || protection ? this.xfs.length : 0
      if (id) this.xfs.push(xml)
      this.xfIds.set(xml, id)
    }
    this.resolved.set(name, id)
    return id
  }

  xml(): string {
    const list = (tag: string, items: string[]) => `<${tag} count="${items.length}">${items.join("")}</${tag}>`
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      (this.numFmts.length ? list("numFmts", this.numFmts) : "") +
      list("fonts", this.fonts) +
      list("fills", this.fills) +
      list("borders", this.borders) +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      list("cellXfs", this.xfs) +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
    )
  }
}

/** The text of a text:p (spaces, tabs, line breaks, spans and links), and the first link in it. */
function paragraph(node: XNode, source: string): { text: string; href: string | null } {
  let href: string | null = null
  const walk = (n: XNode): string => {
    if (n.tagEnd >= n.end || source[n.tagEnd - 2] === "/") return ""
    // Mixed content: the text between children comes from the source.
    const close = source.lastIndexOf("<", n.end - 1)
    let out = ""
    let pos = n.tagEnd
    for (const c of n.children) {
      out += unescapeXml(source.slice(pos, c.start))
      switch (c.local) {
        case "s":
          out += " ".repeat(Number(at(c, "c", NS.text) ?? 1))
          break
        case "tab":
          out += "\t"
          break
        case "line-break":
          out += "\n"
          break
        case "a":
          href ??= at(c, "href", NS.xlink) ?? null
          out += walk(c)
          break
        case "annotation":
        case "note":
        case "bookmark":
        case "bookmark-start":
        case "bookmark-end":
        case "soft-page-break":
          break
        default:
          out += walk(c)
      }
      pos = c.end
    }
    return out + unescapeXml(source.slice(pos, Math.max(pos, close)))
  }
  return { text: walk(node), href }
}

const unescapeXml = (s: string) =>
  s
    .replace(/<[^>]*>/g, "")
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(Number(d)))
    .replace(/&(lt|gt|quot|apos|amp);/g, (_, e: string) => ({ lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" })[e]!)

/** Days from the document's null date (1899-12-30 unless set) to an ISO date or date-time. */
function serial(iso: string, nullDate: number): number | null {
  const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?/.exec(iso)
  if (!m) return null
  const day = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86400000
  const time = m[4] ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6])) / 86400 : 0
  return day - nullDate + time
}

/** An ISO duration (PT12H30M15S, -PT1H) as a fraction of a day. */
function duration(iso: string): number | null {
  const m = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(
    iso,
  )
  if (!m) return null
  const v = Number(m[2] ?? 0) + (Number(m[3] ?? 0) * 3600 + Number(m[4] ?? 0) * 60 + Number(m[5] ?? 0)) / 86400
  return m[1] ? -v : v
}

/** The rows of a table, flattening row groups and header rows. */
function rowsOf(table: XNode): XNode[] {
  const out: XNode[] = []
  for (const c of table.children) {
    if (c.local === "table-row") out.push(c)
    else if (c.local === "table-row-group" || c.local === "table-header-rows" || c.local === "table-rows")
      out.push(...rowsOf(c))
  }
  return out
}

function columnsOf(table: XNode): XNode[] {
  const out: XNode[] = []
  for (const c of table.children) {
    if (c.local === "table-column") out.push(c)
    else if (c.local === "table-column-group" || c.local === "table-header-columns" || c.local === "table-columns")
      out.push(...columnsOf(c))
  }
  return out
}

const ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"]

type SheetView = { freezeRows: number; freezeCols: number; cursor: [number, number]; grid: boolean; zoom: number }

function settingsViews(settings: XNode | null): {
  active: string | null
  views: Map<string, SheetView>
  grid: boolean
} {
  const out = { active: null as string | null, views: new Map<string, SheetView>(), grid: true }
  if (!settings) return out
  const items = (node: XNode) => {
    const m = new Map<string, string>()
    for (const c of node.children) if (c.local === "config-item") m.set(at(c, "name", NS.config) ?? "", c.text)
    return m
  }
  const find = (node: XNode, name: string): XNode | undefined => {
    for (const c of node.children) {
      if (at(c, "name", NS.config) === name) return c
      const deeper = find(c, name)
      if (deeper) return deeper
    }
    return undefined
  }
  const views = find(settings, "Views")
  const first = views?.children.find((c) => c.local === "config-item-map-entry")
  if (!first) return out
  const view = items(first)
  out.active = view.get("ActiveTable") ?? null
  out.grid = view.get("ShowGrid") !== "false"
  const tables = first.children.find((c) => at(c, "name", NS.config) === "Tables")
  for (const t of tables?.children ?? []) {
    const m = items(t)
    const n = (k: string) => Number(m.get(k) ?? 0)
    out.views.set(at(t, "name", NS.config) ?? "", {
      freezeCols: n("HorizontalSplitMode") === 2 ? n("HorizontalSplitPosition") : 0,
      freezeRows: n("VerticalSplitMode") === 2 ? n("VerticalSplitPosition") : 0,
      cursor: [n("CursorPositionY"), n("CursorPositionX")],
      grid: m.get("ShowGrid") !== "false",
      zoom: n("ZoomValue") || 100,
    })
  }
  return out
}

const encoder = new TextEncoder()
const decoder = new TextDecoder()

/** An address list ("Sheet1.A1:Sheet1.B2 Sheet1.D4") as Excel's ("Sheet1!$A$1:$B$2,Sheet1!$D$4"). */
function addressFromOds(text: string): string | null {
  const out: string[] = []
  for (const part of text.trim().split(/\s+/)) {
    const f = formulaFromOds(`of:=[${part.replace(/^\$/, "")}]`)
    if (f === null) return null
    out.push(f)
  }
  return out.join(",")
}

/** The .ods (or .ots) as an .xlsx package, and what didn't come along. */
export async function odsToXlsx(bytes: Uint8Array): Promise<{ bytes: Uint8Array; warnings: string[] }> {
  let entries: ZipEntry[]
  try {
    entries = await readZip(bytes)
  } catch (error) {
    throw new OdsError(error instanceof Error ? error.message : String(error), "damaged")
  }
  const files = new Map(entries.map((e) => [e.normalizedName, e.data]))
  const manifest = files.get("META-INF/manifest.xml")
  if (manifest && /encryption-data/.test(decoder.decode(manifest)))
    throw new OdsError("The document is protected with a password to open.", "encrypted")
  const contentBytes = files.get("content.xml")
  if (!contentBytes) throw new OdsError("The file has no content.xml.", "damaged")
  const content = parseTree(contentBytes)
  const stylesDoc = files.get("styles.xml") ? parseTree(files.get("styles.xml")!) : null
  const settingsDoc = files.get("settings.xml") ? parseTree(files.get("settings.xml")!) : null
  const source = content.source
  const warnings = new Set<string>()

  const styles = new Styles([...(stylesDoc ? [stylesDoc.root] : []), content.root])
  const automatic = new Map<string, XNode>()
  for (const s of kids(
    content.root.children.find((c) => c.local === "automatic-styles"),
    "style",
  ))
    automatic.set(at(s, "name", NS.style) ?? "", s)
  const styleProps = (name: string | undefined) =>
    name ? propsOf(automatic.get(name) ?? ({ children: [] } as unknown as XNode)) : {}

  const body = content.root.children.find((c) => c.local === "body")
  const spreadsheet = body?.children.find((c) => c.local === "spreadsheet")
  if (!spreadsheet) throw new OdsError("The document isn't a spreadsheet.", "damaged")
  const nullDateNode = kids(spreadsheet, "calculation-settings")[0]?.children.find((c) => c.local === "null-date")
  const nullIso = at(nullDateNode, "date-value", NS.table) ?? "1899-12-30"
  const nullDate =
    Date.UTC(Number(nullIso.slice(0, 4)), Number(nullIso.slice(5, 7)) - 1, Number(nullIso.slice(8, 10))) / 86400000
  const date1904 = nullIso.startsWith("1904")
  const settings = settingsViews(settingsDoc?.root ?? null)

  const out = new Map<string, Uint8Array>()
  const put = (name: string, xml: string) => out.set(name, encoder.encode(xml))
  const tables = kids(spreadsheet, "table")
  if (!tables.length) throw new OdsError("The spreadsheet has no sheets.", "damaged")
  const sheetNames = tables.map((t, i) => at(t, "name", NS.table) ?? `Sheet${i + 1}`)
  const sheets: string[] = []
  const definedNames: string[] = []
  const overrides: string[] = []
  let active = 0

  for (const [index, table] of tables.entries()) {
    const name = sheetNames[index]
    if (settings.active === name) active = index
    const tableProps = styleProps(at(table, "style-name", NS.table))
    const hidden = tableProps["table:display"] === "false"
    const tabColor = /^#[0-9a-f]{6}$/i.test(tableProps["tableooo:tab-color"] ?? "")
      ? tableProps["tableooo:tab-color"].slice(1).toUpperCase()
      : ""

    // The sheet's usual column width and row height (the most common ones) stand as its defaults.
    const common = (list: [number | null, number][]) => {
      const counts = new Map<number, number>()
      for (const [v, n] of list)
        if (v !== null) counts.set(Math.round(v * 100) / 100, (counts.get(Math.round(v * 100) / 100) ?? 0) + n)
      return [...counts].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null
    }
    const defaultWidthPt =
      common(
        columnsOf(table).map((c) => [
          points(styleProps(at(c, "style-name", NS.table))["style:column-width"]),
          Number(at(c, "number-columns-repeated", NS.table) ?? 1),
        ]),
      ) ?? 48
    const defaultHeightPt =
      common(
        rowsOf(table).map((r) => [
          points(styleProps(at(r, "style-name", NS.table))["style:row-height"]),
          Number(at(r, "number-rows-repeated", NS.table) ?? 1),
        ]),
      ) ?? 15
    const defaultWidthPx = Math.round((defaultWidthPt * 4) / 3)

    // Columns: widths, hidden ones, and their cells' default styles.
    const colStyle: (string | undefined)[] = []
    const cols: string[] = []
    let col = 0
    for (const c of columnsOf(table)) {
      const repeat = Math.min(Number(at(c, "number-columns-repeated", NS.table) ?? 1), MAX_COLS - col)
      const props = styleProps(at(c, "style-name", NS.table))
      const widthPt = points(props["style:column-width"])
      const isHidden = (at(c, "visibility", NS.table) ?? "visible") !== "visible"
      const def = at(c, "default-cell-style-name", NS.table)
      if (def && def !== "Default" && repeat <= REPEAT_LIMIT) for (let k = 0; k < repeat; k++) colStyle[col + k] = def
      const xf = def && def !== "Default" && repeat <= REPEAT_LIMIT ? styles.xf(def) : 0
      if (widthPt !== null || isHidden || xf) {
        const px = widthPt !== null ? (widthPt * 4) / 3 : defaultWidthPx
        const width = pxToColWidth(Math.round(px))
        if (Math.round(px) !== defaultWidthPx || isHidden || xf)
          cols.push(
            `<col min="${col + 1}" max="${col + repeat}" width="${width}" customWidth="1"${isHidden ? ` hidden="1"` : ""}${xf ? ` style="${xf}"` : ""}/>`,
          )
      }
      col += repeat
      if (col >= MAX_COLS) break
    }

    const rowsXml: string[] = []
    const merges: string[] = []
    const mergeAreas: [number, number, number, number][] = []
    const links: { ref: string; href: string }[] = []
    const notes: { ref: string; row: number; col: number; author: string; text: string; shown: boolean }[] = []
    let row = 0
    for (const r of rowsOf(table)) {
      if (row >= MAX_ROWS) break
      const repeat = Math.min(Number(at(r, "number-rows-repeated", NS.table) ?? 1), MAX_ROWS - row)
      const rprops = styleProps(at(r, "style-name", NS.table))
      const heightPt = points(rprops["style:row-height"])
      const customHeight = rprops["style:use-optimal-row-height"] === "false"
      const rowHidden = (at(r, "visibility", NS.table) ?? "visible") !== "visible"
      const rowDefault = at(r, "default-cell-style-name", NS.table)

      // The row's cells, once; repeated rows repeat them.
      const cells: { col: number; xml: (row: number) => string }[] = []
      let c = 0
      for (const cell of r.children) {
        if (cell.local !== "table-cell" && cell.local !== "covered-table-cell") continue
        const span = Math.min(Number(at(cell, "number-columns-repeated", NS.table) ?? 1), MAX_COLS - c)
        const covered = cell.local === "covered-table-cell"
        const own = at(cell, "style-name", NS.table)
        const styleName = own ?? colStyle[c] ?? rowDefault
        const type = at(cell, "value-type", NS.office) ?? at(cell, "value-type", NS.calcext)
        const formula = at(cell, "formula", NS.table)
        const paragraphs = kids(cell, "p").map((p) => paragraph(p, source))
        const annotation = kids(cell, "annotation")[0]
        const hasContent = !!type || !!formula || paragraphs.length > 0 || !!annotation
        const xf = styleName && styleName !== "Default" ? styles.xf(styleName) : 0
        if (!hasContent && !xf && !at(cell, "number-columns-spanned", NS.table)) {
          c += span
          if (c >= MAX_COLS) break
          continue
        }
        if (!hasContent && span > REPEAT_LIMIT) {
          c += span
          continue
        }
        const colsSpanned = Number(at(cell, "number-columns-spanned", NS.table) ?? 1)
        const rowsSpanned = Number(at(cell, "number-rows-spanned", NS.table) ?? 1)
        const matrixCols = Number(at(cell, "number-matrix-columns-spanned", NS.table) ?? 0)
        const matrixRows = Number(at(cell, "number-matrix-rows-spanned", NS.table) ?? 0)

        // The value, typed as the .xlsx keeps it.
        let v = ""
        let t = ""
        const shown = paragraphs.map((p) => p.text).join("\n")
        const textValue = at(cell, "string-value", NS.office) || shown
        const calcType = at(cell, "value-type", NS.calcext)
        if (type === "error" || calcType === "error") {
          t = "e"
          v = ERRORS.find((e) => e === textValue.trim() || e === shown.trim()) ?? "#N/A"
        } else if (type === "float" || type === "percentage" || type === "currency")
          v = at(cell, "value", NS.office) ?? ""
        else if (type === "date") {
          const s = serial(at(cell, "date-value", NS.office) ?? "", nullDate)
          v = s === null ? "" : String(s)
          if (s === null) t = "inlineStr"
        } else if (type === "time") {
          const d = duration(at(cell, "time-value", NS.office) ?? "")
          v = d === null ? "" : String(d)
        } else if (type === "boolean") {
          t = "b"
          v = (at(cell, "boolean-value", NS.office) ?? "false") === "true" ? "1" : "0"
        } else if (type === "string" || (!type && paragraphs.length)) {
          t = formula ? "str" : "inlineStr"
          v = textValue
        }
        if (t === "inlineStr" && !v) t = ""
        const link = paragraphs.find((p) => p.href)?.href ?? null

        let f = ""
        if (formula) {
          const text = formulaFromOds(formula)
          if (text === null) warnings.add("Some formulas refer to other files; their cells keep their values.")
          else {
            f = x(formulaToFile("=" + text))
            if (t === "inlineStr") t = "str"
          }
        }
        const valueXml = (r: number, k: number) => {
          const fx = !f
            ? ""
            : matrixCols || matrixRows
              ? `<f t="array" ref="${rangeRef(r, k, r + Math.max(1, matrixRows) - 1, k + Math.max(1, matrixCols) - 1)}">${f}</f>`
              : `<f>${f}</f>`
          const vx =
            t === "inlineStr"
              ? `<is><t xml:space="preserve">${x(v)}</t></is>`
              : v !== "" || t === "str"
                ? `<v>${x(v)}</v>`
                : ""
          return `<c r="${cellRef(r, k)}"${xf ? ` s="${xf}"` : ""}${t ? ` t="${t}"` : ""}>${fx}${vx}</c>`
        }
        if (!covered || hasContent || xf)
          for (let k = 0; k < span; k++) {
            const at0 = c + k
            cells.push({ col: at0, xml: (rr) => valueXml(rr, at0) })
          }
        // What belongs to the first copy of the cell only (repeated cells and rows repeat it).
        for (let rr = 0; rr < (hasContent ? repeat : 1); rr++) {
          for (let k = 0; k < span; k++) {
            const r0 = row + rr
            const c0 = c + k
            if (!covered && (colsSpanned > 1 || rowsSpanned > 1))
              merges.push(`<mergeCell ref="${rangeRef(r0, c0, r0 + rowsSpanned - 1, c0 + colsSpanned - 1)}"/>`)
            if (!covered && (colsSpanned > 1 || rowsSpanned > 1))
              mergeAreas.push([r0, c0, r0 + rowsSpanned - 1, c0 + colsSpanned - 1])
            if (link) links.push({ ref: cellRef(r0, c0), href: link })
            if (annotation) {
              const author = kids(annotation, "creator")[0]?.text ?? ""
              const text = kids(annotation, "p")
                .map((p) => paragraph(p, source).text)
                .join("\n")
              notes.push({
                ref: cellRef(r0, c0),
                row: r0,
                col: c0,
                author,
                text,
                shown: at(annotation, "display", NS.office) === "true",
              })
            }
          }
          if (!hasContent) break
        }
        c += span
        if (c >= MAX_COLS) break
      }

      const anyContent = cells.length > 0
      const rowAttrs = (rr: number) =>
        `<row r="${rr + 1}"${heightPt !== null && (customHeight || Math.abs(heightPt - defaultHeightPt) > 0.3) ? ` ht="${Math.round(heightPt * 100) / 100}"${customHeight ? ` customHeight="1"` : ""}` : ""}${rowHidden ? ` hidden="1"` : ""}>`
      const special = rowHidden || (customHeight && heightPt !== null)
      if (anyContent || (special && repeat <= REPEAT_LIMIT)) {
        const count = anyContent || special ? (repeat > REPEAT_LIMIT && !anyContent ? 0 : repeat) : 0
        for (let rr = 0; rr < count; rr++)
          rowsXml.push(rowAttrs(row + rr) + cells.map((cl) => cl.xml(row + rr)).join("") + "</row>")
      }
      row += repeat
    }

    // Notes as comments (and a VML part for the shown ones), links as hyperlinks.
    const sheetPath = `xl/worksheets/sheet${index + 1}.xml`
    const rels: string[] = []
    const linkXml: string[] = []
    for (const [i, l] of links.entries()) {
      if (l.href.startsWith("#")) {
        const loc = formulaFromOds(`of:=[${l.href.slice(1).replace(/^\$/, "")}]`) ?? l.href.slice(1)
        linkXml.push(`<hyperlink ref="${l.ref}" location="${x(loc)}"/>`)
      } else {
        rels.push(
          `<Relationship Id="rIdL${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${x(l.href)}" TargetMode="External"/>`,
        )
        linkXml.push(`<hyperlink ref="${l.ref}" r:id="rIdL${i}"/>`)
      }
    }
    let legacy = ""
    if (notes.length) {
      const authors = [...new Set(notes.map((n) => n.author))]
      put(
        `xl/comments${index + 1}.xml`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authors.map((a) => `<author>${x(a)}</author>`).join("")}</authors><commentList>${notes.map((n) => `<comment ref="${n.ref}" authorId="${authors.indexOf(n.author)}"><text><t xml:space="preserve">${x(n.text)}</t></text></comment>`).join("")}</commentList></comments>`,
      )
      overrides.push(
        `<Override PartName="/xl/comments${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`,
      )
      rels.push(
        `<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${index + 1}.xml"/>`,
      )
      put(
        `xl/drawings/vmlDrawing${index + 1}.vml`,
        `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">${notes.map((n, i) => `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;visibility:${n.shown ? "visible" : "hidden"}"><x:ClientData ObjectType="Note"><x:Row>${n.row}</x:Row><x:Column>${n.col}</x:Column>${n.shown ? "<x:Visible/>" : ""}</x:ClientData></v:shape>`).join("")}</xml>`,
      )
      rels.push(
        `<Relationship Id="rIdV" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${index + 1}.vml"/>`,
      )
      legacy = `<legacyDrawing r:id="rIdV"/>`
    }

    // The view: frozen panes and the cursor from settings.xml.
    const v = settings.views.get(name)
    let pane = ""
    let selection = ""
    // The cursor on a merged cell selects the whole merge, as Excel saves it.
    const at0 = v ? cellRef(v.cursor[0], v.cursor[1]) : ""
    const merged =
      v &&
      mergeAreas.find(
        ([r0, c0, r1, c1]) => v.cursor[0] >= r0 && v.cursor[0] <= r1 && v.cursor[1] >= c0 && v.cursor[1] <= c1,
      )
    const sqref = merged ? rangeRef(...merged) : at0
    if (v && (v.freezeRows || v.freezeCols)) {
      const activePane = v.freezeRows && v.freezeCols ? "bottomRight" : v.freezeRows ? "bottomLeft" : "topRight"
      pane = `<pane${v.freezeCols ? ` xSplit="${v.freezeCols}"` : ""}${v.freezeRows ? ` ySplit="${v.freezeRows}"` : ""} topLeftCell="${cellRef(v.freezeRows, v.freezeCols)}" activePane="${activePane}" state="frozen"/>`
      selection = `<selection pane="${activePane}" activeCell="${at0}" sqref="${sqref}"/>`
    } else if (v) selection = `<selection activeCell="${at0}" sqref="${sqref}"/>`
    const grid = v ? v.grid && settings.grid : settings.grid
    const zoom = v && v.zoom !== 100 ? ` zoomScale="${v.zoom}"` : ""
    const view = `<sheetViews><sheetView workbookViewId="0"${grid ? "" : ` showGridLines="0"`}${zoom}${index === active ? ` tabSelected="1"` : ""}>${pane}${selection}</sheetView></sheetViews>`

    put(
      sheetPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        (tabColor ? `<sheetPr><tabColor rgb="FF${tabColor}"/></sheetPr>` : "") +
        view +
        `<sheetFormatPr defaultColWidth="${pxToColWidth(defaultWidthPx)}" defaultRowHeight="${Math.round(defaultHeightPt * 100) / 100}" customHeight="1"/>` +
        (cols.length ? `<cols>${cols.join("")}</cols>` : "") +
        `<sheetData>${rowsXml.join("")}</sheetData>` +
        (merges.length ? `<mergeCells count="${merges.length}">${merges.join("")}</mergeCells>` : "") +
        (linkXml.length ? `<hyperlinks>${linkXml.join("")}</hyperlinks>` : "") +
        legacy +
        `</worksheet>`,
    )
    if (rels.length)
      put(
        `xl/worksheets/_rels/sheet${index + 1}.xml.rels`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`,
      )
    overrides.push(
      `<Override PartName="/${sheetPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    sheets.push(
      `<sheet name="${x(name)}" sheetId="${index + 1}"${hidden ? ` state="hidden"` : ""} r:id="rId${index + 1}"/>`,
    )

    // Print ranges, and names that belong to this sheet.
    const print = at(table, "print-ranges", NS.table)
    const area = print && addressFromOds(print)
    if (area) definedNames.push(`<definedName name="_xlnm.Print_Area" localSheetId="${index}">${x(area)}</definedName>`)
    for (const exprs of kids(table, "named-expressions")) definedNames.push(...namedExpressions(exprs, index))

    // What doesn't come along yet.
  }
  for (const exprs of kids(spreadsheet, "named-expressions")) definedNames.push(...namedExpressions(exprs, -1))
  if (kids(spreadsheet, "data-pilot-tables").length)
    warnings.add("Pivot tables (DataPilot) in .ods files don't open in Cells yet.")
  if (kids(spreadsheet, "content-validations").length)
    warnings.add("Data validation in .ods files doesn't open in Cells yet.")
  // calcext:conditional-formats live inside each table in LibreOffice's files.
  if (/<calcext:conditional-formats/.test(source))
    warnings.add("Conditional formatting in .ods files doesn't open in Cells yet.")
  if (/<draw:frame|<draw:custom-shape/.test(source))
    warnings.add("Charts, pictures and shapes in .ods files don't open in Cells yet.")

  put(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      (date1904 ? `<workbookPr date1904="1"/>` : "") +
      `<bookViews><workbookView${active ? ` activeTab="${active}"` : ""}/></bookViews>` +
      `<sheets>${sheets.join("")}</sheets>` +
      (definedNames.length ? `<definedNames>${definedNames.join("")}</definedNames>` : "") +
      `</workbook>`,
  )
  put("xl/styles.xml", styles.xml())
  put(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      tables
        .map(
          (_, i) =>
            `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join("") +
      `<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  put(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    overrides.join("") +
    `</Types>`
  const order = ["[Content_Types].xml", ...out.keys()]
  out.set("[Content_Types].xml", encoder.encode(contentTypes))
  return { bytes: storedZip(order.map((name) => ({ name, data: out.get(name)! }))), warnings: [...warnings] }

  function namedExpressions(exprs: XNode, sheet: number): string[] {
    const list: string[] = []
    const scope = sheet >= 0 ? ` localSheetId="${sheet}"` : ""
    for (const n of exprs.children) {
      const name = at(n, "name", NS.table)
      if (!name) continue
      let ref: string | null = null
      if (n.local === "named-range") ref = addressFromOds(at(n, "cell-range-address", NS.table) ?? "")
      else if (n.local === "named-expression") ref = formulaFromOds(at(n, "expression", NS.table) ?? "")
      if (ref === null) continue
      list.push(`<definedName name="${x(name)}"${scope}>${x(formulaToFile("=" + ref))}</definedName>`)
    }
    return list
  }
}

export type OdsRead = XlsxRead

/** Opens an .ods (or .ots) spreadsheet. */
export async function openOds(bytes: Uint8Array, fileName = "Book1.ods", options: OpenOptions = {}): Promise<OdsRead> {
  const converted = await odsToXlsx(bytes)
  const read = await openXlsx(converted.bytes, fileName.replace(/\.o[dt]s$/i, "") + ".xlsx", {
    ...options,
    calculate: false,
  })
  // Everything calculates again, as Excel does with an .ods: the saved results stand in until then.
  read.workbook.calc.rebuild(false)
  if (options.calculate !== false) read.workbook.calc.calculateNow(true)
  read.workbook.clearHistory()
  return { ...read, warnings: [...new Set([...converted.warnings, ...read.warnings])] }
}
