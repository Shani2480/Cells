import type { Border, BorderStyle, CellStyle, CFStyle, HAlign, Underline, VAlign } from "@cells/core"
import { escapeXml } from "../xml"
import { argb, DEFAULT_THEME_COLORS, resolveColor, type ColorAttrs, type ThemeColors } from "./colors"
import { BUILTIN_CELL_STYLES, BUILTIN_NUMBER_FORMATS, INDEXED_COLORS, NS_MAIN } from "./constants"
import { at, attrText, kid, kids, parseTree, raw, startTag, type XDoc, type XNode } from "./dom"

/** The font a new workbook uses (Excel 365): Aptos Narrow 11 pt, the theme's minor font. */
export const DEFAULT_FONT = { name: "Aptos Narrow", size: 11 }

const BORDER_STYLES = new Set<string>([
  "thin", "medium", "thick", "hair", "dotted", "dashed", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
  "mediumDashDotDot", "slantDashDot", "double",
]) // prettier-ignore

const colorAttrs = (n: XNode | undefined): ColorAttrs | null =>
  n
    ? { rgb: at(n, "rgb"), theme: at(n, "theme"), tint: at(n, "tint"), indexed: at(n, "indexed"), auto: at(n, "auto") }
    : null

const on = (n: XNode | undefined) => !!n && at(n, "val") !== "0" && at(n, "val") !== "false"
const flag = (v: string | undefined) => v === "1" || v === "true"

type FontFields = Pick<
  CellStyle,
  "fontName" | "fontSize" | "bold" | "italic" | "underline" | "strike" | "vertAlign" | "fontColor"
>

/** styles.xml as read: Cells formats for each cell format (xf) and differential format (dxf), plus the source. */
export class StylesIn {
  readonly doc: XDoc
  readonly numFmts = new Map<number, string>()
  readonly palette: string[]
  readonly defaultFont: { name: string; size: number; color: string | null }
  private readonly fonts: XNode[]
  private readonly fills: XNode[]
  private readonly borders: XNode[]
  readonly cellXfs: XNode[]
  readonly dxfs: XNode[]
  private readonly styleNames = new Map<number, string>()
  private readonly cache = new Map<number, CellStyle>()

  constructor(
    source: Uint8Array | string,
    readonly theme: ThemeColors = DEFAULT_THEME_COLORS,
  ) {
    this.doc = parseTree(source)
    const root = this.doc.root
    for (const n of kids(kid(root, "numFmts"), "numFmt")) {
      const id = Number(at(n, "numFmtId"))
      const code = at(n, "formatCode")
      if (Number.isInteger(id) && code !== undefined) this.numFmts.set(id, code)
    }
    const custom = kids(kid(kid(root, "colors"), "indexedColors"), "rgbColor").map((c) =>
      (at(c, "rgb") ?? "").slice(-6),
    )
    this.palette = custom.length ? custom.concat(INDEXED_COLORS.slice(custom.length)) : INDEXED_COLORS
    this.fonts = kids(kid(root, "fonts"), "font")
    this.fills = kids(kid(root, "fills"), "fill")
    this.borders = kids(kid(root, "borders"), "border")
    this.cellXfs = kids(kid(root, "cellXfs"), "xf")
    this.dxfs = kids(kid(root, "dxfs"), "dxf")
    for (const s of kids(kid(root, "cellStyles"), "cellStyle")) {
      const name = at(s, "name")
      if (name) this.styleNames.set(Number(at(s, "xfId") ?? 0), name)
    }
    const f0 = this.fonts[0]
    this.defaultFont = {
      name: at(kid(f0, "name"), "val") ?? DEFAULT_FONT.name,
      size: Number(at(kid(f0, "sz"), "val") ?? DEFAULT_FONT.size),
      color: this.color(kid(f0, "color")),
    }
  }

  private color(n: XNode | undefined): string | null {
    const attrs = colorAttrs(n)
    return attrs ? resolveColor(attrs, this.theme, this.palette) : null
  }

  numberFormat(id: number): string | undefined {
    return this.numFmts.get(id) ?? BUILTIN_NUMBER_FORMATS[id]
  }

  private font(n: XNode | undefined, relative: boolean): FontFields {
    const out: FontFields = {}
    if (!n) return out
    const name = at(kid(n, "name"), "val")
    const size = at(kid(n, "sz"), "val")
    if (name && (!relative || name !== this.defaultFont.name)) out.fontName = name
    if (size && (!relative || Number(size) !== this.defaultFont.size)) out.fontSize = Number(size)
    if (on(kid(n, "b"))) out.bold = true
    if (on(kid(n, "i"))) out.italic = true
    if (on(kid(n, "strike"))) out.strike = true
    const u = kid(n, "u")
    if (u && at(u, "val") !== "none") out.underline = (at(u, "val") ?? "single") as Underline
    const va = at(kid(n, "vertAlign"), "val")
    if (va === "superscript" || va === "subscript") out.vertAlign = va
    const color = this.color(kid(n, "color"))
    if (color && (!relative || color !== this.defaultFont.color)) out.fontColor = color
    return out
  }

  private border(n: XNode | undefined): Partial<CellStyle> {
    const out: Partial<CellStyle> = {}
    if (!n) return out
    const side = (local: string): Border | undefined => {
      const s = kid(n, local)
      const style = at(s, "style")
      if (!style || !BORDER_STYLES.has(style)) return undefined
      const color = this.color(kid(s, "color"))
      return color ? { style: style as BorderStyle, color } : { style: style as BorderStyle }
    }
    const sides = [
      ["borderLeft", side("left") ?? side("start")],
      ["borderRight", side("right") ?? side("end")],
      ["borderTop", side("top")],
      ["borderBottom", side("bottom")],
    ] as const
    for (const [key, b] of sides) if (b) out[key] = b
    const diagonal = side("diagonal")
    if (diagonal) {
      out.borderDiagonal = diagonal
      if (flag(at(n, "diagonalUp"))) out.diagonalUp = true
      if (flag(at(n, "diagonalDown"))) out.diagonalDown = true
    }
    return out
  }

  private fill(n: XNode | undefined, differential: boolean): string | undefined {
    const p = kid(n, "patternFill")
    if (!p) return undefined
    const type = at(p, "patternType")
    if (differential) {
      // A dxf's solid fill is its bgColor (fgColor when Excel wrote a pattern type).
      const c = this.color(kid(p, "bgColor")) ?? this.color(kid(p, "fgColor"))
      return c ?? undefined
    }
    if (type !== "solid") return undefined
    return this.color(kid(p, "fgColor")) ?? undefined
  }

  /** The Cells format of cell format (xf) `index`. */
  xfStyle(index: number): CellStyle {
    const cached = this.cache.get(index)
    if (cached) return cached
    const xf = this.cellXfs[index]
    const style: CellStyle = {}
    if (xf) {
      const fmt = this.numberFormat(Number(at(xf, "numFmtId") ?? 0))
      if (fmt && fmt !== "General") style.numberFormat = fmt
      Object.assign(style, this.font(this.fonts[Number(at(xf, "fontId") ?? 0)], true))
      const fill = this.fill(this.fills[Number(at(xf, "fillId") ?? 0)], false)
      if (fill) style.fill = fill
      Object.assign(style, this.border(this.borders[Number(at(xf, "borderId") ?? 0)]))
      const a = kid(xf, "alignment")
      if (a) {
        const h = at(a, "horizontal")
        if (h && h !== "general") style.hAlign = h as HAlign
        const v = at(a, "vertical")
        if (v && v !== "bottom") style.vAlign = v as VAlign
        const indent = Number(at(a, "indent") ?? 0)
        if (indent) style.indent = indent
        if (flag(at(a, "wrapText"))) style.wrap = true
        if (flag(at(a, "shrinkToFit"))) style.shrink = true
        const rotation = Number(at(a, "textRotation") ?? 0)
        if (rotation) style.rotation = rotation
      }
      const p = kid(xf, "protection")
      if (p) {
        if (at(p, "locked") === "0" || at(p, "locked") === "false") style.locked = false
        if (flag(at(p, "hidden"))) style.hiddenFormula = true
      }
      const named = this.styleNames.get(Number(at(xf, "xfId") ?? 0))
      if (named && named !== "Normal") style.cellStyle = named
    }
    this.cache.set(index, style)
    return style
  }

  /** The Cells differential format (conditional formatting) of dxf `index`. */
  dxfStyle(index: number): CFStyle {
    const dxf = this.dxfs[index]
    const out: CFStyle = {}
    if (!dxf) return out
    const font = this.font(kid(dxf, "font"), false)
    for (const k of ["bold", "italic", "underline", "strike", "fontColor"] as const)
      if (font[k] !== undefined) (out as Record<string, unknown>)[k] = font[k]
    const nf = kid(dxf, "numFmt")
    const code = nf ? (at(nf, "formatCode") ?? this.numberFormat(Number(at(nf, "numFmtId")))) : undefined
    if (code) out.numberFormat = code
    const fill = this.fill(kid(dxf, "fill"), true)
    if (fill) out.fill = fill
    const b = this.border(kid(dxf, "border"))
    for (const k of ["borderTop", "borderBottom", "borderLeft", "borderRight"] as const) if (b[k]) out[k] = b[k]
    return out
  }
}

// ── Writing ──────────────────────────────────────────────────────────────────────────────────────────────────

const BUILTIN_BY_CODE = new Map(Object.entries(BUILTIN_NUMBER_FORMATS).map(([id, code]) => [code, Number(id)]))

/** A section of styles.xml: the source's entries (as written) and the ones added for new formats. */
class Section {
  readonly items: string[]
  private readonly index = new Map<string, number>()
  constructor(source: string[] = []) {
    this.items = [...source]
    // The source's entries count as written: a format that matches one exactly reuses it.
    source.forEach((xml, i) => {
      if (!this.index.has(xml)) this.index.set(xml, i)
    })
  }
  add(xml: string): number {
    const found = this.index.get(xml)
    if (found !== undefined) return found
    const id = this.items.length
    this.items.push(xml)
    this.index.set(xml, id)
    return id
  }
}

const colorXml = (tag: string, hex: string | undefined) => (hex ? `<${tag} rgb="${argb(hex)}"/>` : "")

/**
 * Builds styles.xml for a save. Formats that came from the file keep their own xf (so theme colours, patterns and
 * theme fonts survive untouched); new formats get new fonts, fills, borders and xfs appended after the source's.
 */
export class StylesOut {
  private readonly numFmts = new Map<number, string>()
  private nextNumFmt = 164
  private readonly fonts: Section
  private readonly fills: Section
  private readonly borders: Section
  private readonly cellStyleXfs: Section
  private readonly cellXfs: Section
  private readonly cellStyles: string[]
  private readonly styleXfByName = new Map<string, number>()
  private readonly dxfs: Section
  private readonly byStyle = new Map<number, number>()
  private readonly font: { name: string; size: number }

  /** `reuse`: Cells style id → the source's xf index, for formats read from the file and not changed since. */
  constructor(
    private readonly source: StylesIn | null,
    private readonly reuse: ReadonlyMap<number, number> = new Map(),
  ) {
    const doc = source?.doc
    const items = (section: string, item: string) =>
      doc ? kids(kid(doc.root, section), item).map((n) => raw(doc, n)) : []
    if (source) {
      for (const [id, code] of source.numFmts) {
        this.numFmts.set(id, code)
        this.nextNumFmt = Math.max(this.nextNumFmt, id + 1)
      }
    }
    this.font = source ? { name: source.defaultFont.name, size: source.defaultFont.size } : DEFAULT_FONT
    this.fonts = new Section(items("fonts", "font"))
    this.fills = new Section(items("fills", "fill"))
    this.borders = new Section(items("borders", "border"))
    this.cellStyleXfs = new Section(items("cellStyleXfs", "xf"))
    this.cellXfs = new Section(items("cellXfs", "xf"))
    this.dxfs = new Section(items("dxfs", "dxf"))
    this.cellStyles = items("cellStyles", "cellStyle")
    if (doc)
      for (const s of kids(kid(doc.root, "cellStyles"), "cellStyle"))
        this.styleXfByName.set(at(s, "name") ?? "", Number(at(s, "xfId") ?? 0))
    if (!this.fonts.items.length)
      this.fonts.add(
        `<font><sz val="${this.font.size}"/><color theme="1"/><name val="${escapeXml(this.font.name)}"/><family val="2"/><scheme val="minor"/></font>`,
      )
    if (!this.fills.items.length) {
      this.fills.add(`<fill><patternFill patternType="none"/></fill>`)
      this.fills.add(`<fill><patternFill patternType="gray125"/></fill>`)
    }
    if (!this.borders.items.length) this.borders.add(`<border><left/><right/><top/><bottom/><diagonal/></border>`)
    if (!this.cellStyleXfs.items.length) this.cellStyleXfs.add(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`)
    if (!this.cellXfs.items.length) this.cellXfs.add(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`)
    if (!this.cellStyles.length) {
      this.cellStyles.push(`<cellStyle name="Normal" xfId="0" builtinId="0"/>`)
      this.styleXfByName.set("Normal", 0)
    }
  }

  private numFmtId(code: string | undefined): number {
    if (!code || code === "General") return 0
    for (const [id, c] of this.numFmts) if (c === code) return id
    const builtin = BUILTIN_BY_CODE.get(code)
    if (builtin !== undefined) return builtin
    const id = this.nextNumFmt++
    this.numFmts.set(id, code)
    return id
  }

  private fontXml(s: Partial<CellStyle>, differential: boolean): string {
    let x = ""
    if (s.bold) x += "<b/>"
    if (s.italic) x += "<i/>"
    if (s.strike) x += "<strike/>"
    if (s.underline) x += s.underline === "single" ? "<u/>" : `<u val="${s.underline}"/>`
    if (s.vertAlign) x += `<vertAlign val="${s.vertAlign}"/>`
    if (!differential) x += `<sz val="${s.fontSize ?? this.font.size}"/>`
    x += s.fontColor ? colorXml("color", s.fontColor) : differential ? "" : `<color theme="1"/>`
    if (!differential) {
      x += `<name val="${escapeXml(s.fontName ?? this.font.name)}"/><family val="2"/>`
      if (!s.fontName || s.fontName === this.font.name) x += `<scheme val="minor"/>`
    }
    return `<font>${x}</font>`
  }

  private borderXml(s: Partial<CellStyle>): string {
    const side = (tag: string, b: Border | undefined) =>
      b
        ? `<${tag} style="${b.style}">${b.color ? colorXml("color", b.color) : `<color indexed="64"/>`}</${tag}>`
        : `<${tag}/>`
    return (
      `<border${attrText({ diagonalUp: s.diagonalUp, diagonalDown: s.diagonalDown })}>` +
      side("left", s.borderLeft) +
      side("right", s.borderRight) +
      side("top", s.borderTop) +
      side("bottom", s.borderBottom) +
      side("diagonal", s.borderDiagonal) +
      `</border>`
    )
  }

  private styleXf(name: string | undefined): number {
    if (!name) return 0
    const found = this.styleXfByName.get(name)
    if (found !== undefined) return found
    // A gallery style the file didn't define: a style record with the default look; the cells carry the format.
    const id = this.cellStyleXfs.items.length
    this.cellStyleXfs.items.push(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`)
    const builtin = BUILTIN_CELL_STYLES[name]
    this.cellStyles.push(`<cellStyle${attrText({ name, xfId: id, builtinId: builtin })}/>`)
    this.styleXfByName.set(name, id)
    return id
  }

  /** The xf index for a Cells style (id and its format). */
  xf(id: number, style: CellStyle): number {
    if (id === 0) return 0
    const known = this.byStyle.get(id)
    if (known !== undefined) return known
    const reused = this.reuse.get(id)
    if (reused !== undefined) {
      this.byStyle.set(id, reused)
      return reused
    }
    const numFmtId = this.numFmtId(style.numberFormat)
    const fontId = this.fonts.add(this.fontXml(style, false))
    const fillId = style.fill
      ? this.fills.add(
          `<fill><patternFill patternType="solid">${colorXml("fgColor", style.fill)}<bgColor indexed="64"/></patternFill></fill>`,
        )
      : 0
    const borderId = this.borders.add(this.borderXml(style))
    const align = attrText({
      horizontal: style.hAlign,
      vertical: style.vAlign,
      textRotation: style.rotation,
      wrapText: style.wrap,
      indent: style.indent,
      shrinkToFit: style.shrink,
    })
    const protect = attrText({ locked: style.locked === false ? "0" : undefined, hidden: style.hiddenFormula })
    const xml =
      `<xf${attrText({
        numFmtId,
        fontId,
        fillId,
        borderId,
        xfId: this.styleXf(style.cellStyle),
        applyNumberFormat: numFmtId !== 0,
        applyFont: fontId !== 0,
        applyFill: fillId !== 0,
        applyBorder: borderId !== 0,
        applyAlignment: align !== "",
        applyProtection: protect !== "",
      })}` +
      (align || protect
        ? `>${align ? `<alignment${align}/>` : ""}${protect ? `<protection${protect}/>` : ""}</xf>`
        : "/>")
    const index = this.cellXfs.add(xml)
    this.byStyle.set(id, index)
    return index
  }

  /** The dxf index for a conditional format's look. */
  dxf(style: CFStyle): number {
    let x = ""
    if (style.bold || style.italic || style.underline || style.strike || style.fontColor) x += this.fontXml(style, true)
    if (style.numberFormat) {
      const id = this.numFmtId(style.numberFormat)
      x += `<numFmt numFmtId="${id}" formatCode="${escapeXml(style.numberFormat)}"/>`
    }
    if (style.fill) x += `<fill><patternFill>${colorXml("bgColor", style.fill)}</patternFill></fill>`
    if (style.borderLeft || style.borderRight || style.borderTop || style.borderBottom) x += this.borderXml(style)
    return this.dxfs.add(`<dxf>${x}</dxf>`)
  }

  xml(): string {
    const doc = this.source?.doc
    const open = doc ? startTag(doc, doc.root) : `<styleSheet xmlns="${NS_MAIN}"`
    const section = (tag: string, items: string[]) =>
      items.length ? `<${tag} count="${items.length}">${items.join("")}</${tag}>` : ""
    const numFmts = [...this.numFmts].map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`)
    const keep = (local: string) => {
      const n = doc ? kid(doc.root, local) : undefined
      return n && doc ? raw(doc, n) : ""
    }
    return (
      `${open}>` +
      section("numFmts", numFmts) +
      section("fonts", this.fonts.items) +
      section("fills", this.fills.items) +
      section("borders", this.borders.items) +
      section("cellStyleXfs", this.cellStyleXfs.items) +
      section("cellXfs", this.cellXfs.items) +
      section("cellStyles", this.cellStyles) +
      (this.dxfs.items.length ? section("dxfs", this.dxfs.items) : `<dxfs count="0"/>`) +
      (keep("tableStyles") ||
        `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>`) +
      keep("colors") +
      keep("extLst") +
      `</styleSheet>`
    )
  }
}
