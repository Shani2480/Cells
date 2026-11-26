import type {
  ChildBox,
  DashStyle,
  DrawingObject,
  LineEnd,
  ShapeColor,
  ShapeFill,
  ShapeLine,
  ShapeText,
  SheetGroup,
  SheetPicture,
  SheetShape,
  TextParagraph,
  TextRun,
  ThemeSlot,
} from "@cells/core"
import { NS_R } from "./constants"
import { at, kid, kids, type XDoc, type XNode } from "./dom"
import { mathFallbackText, mathLook, mathParaXml, NS_A14, readMath } from "./omml"
import type { MathList, Model3D } from "@cells/core"
import { model3dObjectXml } from "./model3d"
import { adjustBlipXml, imgPropsExtXml, readPictureAdjust } from "./picture-adjust"

/**
 * DrawingML shapes, pictures and groups in a sheet's drawing (xdr:sp, xdr:cxnSp, xdr:pic, xdr:grpSp): read into
 * Cells' objects and written from them. Sizes in the file are EMUs (9,525 to a pixel), line widths EMUs too (12,700
 * to a point), angles 60,000ths of a degree and percentages 1,000ths.
 */

export const EMU = 9525
/** Office's SVG picture extension on a blip. */
const SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
const NS_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
const PT = 12700
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

const num = (v: string | undefined, fallback = 0) => {
  const n = Number(v)
  return Number.isFinite(n) ? n : fallback
}

// ── Colours ──────────────────────────────────────────────────────────────────────────────────────────────────

/** DrawingML's other names for the theme's text and background colours. */
const SLOT_ALIASES: Record<string, ThemeSlot> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" }
const SLOTS = new Set([
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
])

/** A colour element's colour (srgbClr, schemeClr, sysClr, prstClr); `ph` stands in for phClr (the style's colour). */
export function readColor(node: XNode | undefined, ph?: ShapeColor): ShapeColor | undefined {
  if (!node) return undefined
  let color: ShapeColor | undefined
  const val = at(node, "val")
  if (node.local === "srgbClr" && val) color = { rgb: val.toUpperCase() }
  else if (node.local === "sysClr")
    color = { rgb: (at(node, "lastClr") ?? (val === "window" ? "FFFFFF" : "000000")).toUpperCase() }
  else if (node.local === "prstClr") color = { rgb: PRESET_COLORS[val ?? ""] ?? "000000" }
  else if (node.local === "schemeClr" && val) {
    if (val === "phClr") color = ph ? { ...ph } : { theme: "accent1" }
    else {
      const slot = SLOT_ALIASES[val] ?? val
      color = SLOTS.has(slot) ? { theme: slot as ThemeSlot } : { theme: "accent1" }
    }
  }
  if (!color) return undefined
  for (const m of node.children) {
    const v = num(at(m, "val")) / 100000
    if (m.local === "lumMod") color.lumMod = v
    else if (m.local === "lumOff") color.lumOff = v
    else if (m.local === "alpha") color.alpha = v
    // Shade darkens toward black and tint lightens toward white: near enough as luminance changes.
    else if (m.local === "shade") color.lumMod = (color.lumMod ?? 1) * v
    else if (m.local === "tint") {
      color.lumMod = (color.lumMod ?? 1) * v
      color.lumOff = (color.lumOff ?? 0) + (1 - v)
    }
  }
  return color
}

const PRESET_COLORS: Record<string, string> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "008000",
  blue: "0000FF",
  yellow: "FFFF00",
  gray: "808080",
  orange: "FFA500",
}

const COLOR_ELEMENTS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"])
const colorIn = (node: XNode | undefined, ph?: ShapeColor) =>
  readColor(
    node?.children.find((c) => COLOR_ELEMENTS.has(c.local)),
    ph,
  )

export function colorXml(c: ShapeColor): string {
  const mods =
    (c.lumMod !== undefined && c.lumMod !== 1 ? `<a:lumMod val="${Math.round(c.lumMod * 100000)}"/>` : "") +
    (c.lumOff ? `<a:lumOff val="${Math.round(c.lumOff * 100000)}"/>` : "") +
    (c.alpha !== undefined && c.alpha < 1 ? `<a:alpha val="${Math.round(c.alpha * 100000)}"/>` : "")
  const el = c.theme ? `a:schemeClr val="${c.theme}"` : `a:srgbClr val="${c.rgb ?? "000000"}"`
  const name = c.theme ? "a:schemeClr" : "a:srgbClr"
  return mods ? `<${el}>${mods}</${name}>` : `<${el}/>`
}

// ── Fill, line, text ─────────────────────────────────────────────────────────────────────────────────────────

/** A fill element among a node's children; undefined when none is given (the style decides). */
function readFill(parent: XNode | undefined, ph?: ShapeColor): ShapeFill | undefined {
  for (const c of parent?.children ?? []) {
    if (c.local === "noFill") return { kind: "none" }
    if (c.local === "solidFill") {
      const color = colorIn(c, ph)
      return color ? { kind: "solid", color } : { kind: "none" }
    }
    // A gradient or pattern shows as its first colour (the file's own stays while the shape is unchanged).
    if (c.local === "gradFill") {
      const color = colorIn(kid(kid(c, "gsLst"), "gs"), ph)
      if (color) return { kind: "solid", color }
    }
    if (c.local === "pattFill") {
      const color = colorIn(kid(c, "fgClr"), ph)
      if (color) return { kind: "solid", color }
    }
  }
  return undefined
}

export const fillXml = (f: ShapeFill) =>
  f.kind === "none" ? "<a:noFill/>" : `<a:solidFill>${colorXml(f.color)}</a:solidFill>`

const DASHES = new Set(["solid", "dot", "dash", "lgDash", "dashDot", "sysDot", "sysDash"])
const ENDS = new Set(["none", "triangle", "arrow", "stealth", "diamond", "oval"])

function readLine(ln: XNode | undefined, base: ShapeLine, ph?: ShapeColor): ShapeLine {
  if (!ln) return base
  const out: ShapeLine = { ...base }
  const w = at(ln, "w")
  if (w !== undefined) out.width = num(w) / PT
  const fill = readFill(ln, ph)
  if (fill) out.color = fill.kind === "solid" ? fill.color : null
  const dash = at(kid(ln, "prstDash"), "val")
  if (dash && DASHES.has(dash)) out.dash = dash as DashStyle
  const head = at(kid(ln, "headEnd"), "type")
  const tail = at(kid(ln, "tailEnd"), "type")
  if (head && ENDS.has(head) && head !== "none") out.head = head as LineEnd
  if (tail && ENDS.has(tail) && tail !== "none") out.tail = tail as LineEnd
  return out
}

export function lineXml(l: ShapeLine): string {
  const fill = l.color ? `<a:solidFill>${colorXml(l.color)}</a:solidFill>` : "<a:noFill/>"
  const dash = l.dash && l.dash !== "solid" ? `<a:prstDash val="${l.dash}"/>` : ""
  const head = l.head && l.head !== "none" ? `<a:headEnd type="${l.head}"/>` : ""
  const tail = l.tail && l.tail !== "none" ? `<a:tailEnd type="${l.tail}"/>` : ""
  return `<a:ln w="${Math.round(l.width * PT)}">${fill}${dash}${head}${tail}</a:ln>`
}

function readRun(rPr: XNode | undefined, text: string, fontColor?: ShapeColor): TextRun {
  const run: TextRun = { text }
  if (!rPr) return fontColor ? { ...run, color: fontColor } : run
  if (at(rPr, "b") === "1") run.bold = true
  if (at(rPr, "i") === "1") run.italic = true
  const u = at(rPr, "u")
  if (u && u !== "none") run.underline = true
  const strike = at(rPr, "strike")
  if (strike && strike !== "noStrike") run.strike = true
  const sz = at(rPr, "sz")
  if (sz) run.size = num(sz) / 100
  const font = at(kid(rPr, "latin"), "typeface")
  if (font && !font.startsWith("+")) run.font = font
  const fill = readFill(rPr)
  const color = fill?.kind === "solid" ? fill.color : fontColor
  if (color) run.color = color
  return run
}

function readText(body: XNode | undefined, fontColor?: ShapeColor, doc?: XDoc): ShapeText | undefined {
  if (!body) return undefined
  const pr = kid(body, "bodyPr")
  const text: ShapeText = { paragraphs: [] }
  const anchor = at(pr, "anchor")
  if (anchor === "t" || anchor === "ctr" || anchor === "b") text.anchor = anchor
  const vert = at(pr, "vert")
  if (vert === "vert" || vert === "vert270") text.vertical = vert
  if (at(pr, "wrap") === "none") text.wrap = false
  else text.wrap = true
  if (kid(pr, "spAutoFit")) text.autoFit = true
  const ins = (name: string, d: number) => num(at(pr, name), d * EMU) / EMU
  if (pr && ["lIns", "tIns", "rIns", "bIns"].some((n) => at(pr, n) !== undefined))
    text.insets = { l: ins("lIns", 9.6), t: ins("tIns", 4.8), r: ins("rIns", 9.6), b: ins("bIns", 4.8) }
  for (const p of kids(body, "p")) {
    const para: TextParagraph = { runs: [] }
    const algn = at(kid(p, "pPr"), "algn")
    if (algn === "l" || algn === "ctr" || algn === "r" || algn === "just") para.align = algn
    const endPr = kid(p, "endParaRPr")
    // An equation (Insert › Equation): Office Math in the paragraph.
    const math = doc && p.children.find((c) => c.local === "m" && c.uri === NS_A14)
    const mathPara = math?.children.find((c) => c.local === "oMathPara" || c.local === "oMath")
    if (doc && mathPara) {
      para.math = readMath(doc, mathPara)
      // Text beside an equation in one paragraph: kept as the equation's normal (upright) text, as Office Math holds
      // text, so none of it is lost.
      const texts = p.children.filter((c) => c.local === "r" || c.local === "fld")
      if (texts.length) {
        const line: MathList = []
        for (const c of p.children) {
          if (c === math) line.push(...para.math.flat())
          else if (c.local === "r" || c.local === "fld") line.push({ t: "r", text: kid(c, "t")?.text ?? "", sty: "p" })
        }
        para.math = [line]
      }
      // The equation's look: its size and colour (math italic and Cambria Math are the equation's own).
      const look: TextRun = { ...mathLook(mathPara, (rPr) => readRun(rPr, "", fontColor)), text: "" }
      delete look.italic
      delete look.font
      para.runs = [look]
      text.paragraphs.push(para)
      continue
    }
    for (const r of p.children) {
      if (r.local === "r" || r.local === "fld")
        para.runs.push(readRun(kid(r, "rPr"), kid(r, "t")?.text ?? "", fontColor))
      // A line break starts a new paragraph in Cells (the file's own stays while the shape is unchanged).
      else if (r.local === "br") {
        text.paragraphs.push({ ...para, runs: para.runs })
        para.runs = []
      }
    }
    // An empty paragraph keeps its size and colour, so typed text looks as Excel would show it.
    if (!para.runs.length && endPr) {
      const look = readRun(endPr, "", fontColor)
      if (look.size || look.color || look.bold) para.runs.push(look)
    }
    text.paragraphs.push(para)
  }
  return text
}

function runXml(r: TextRun): string {
  const attrs =
    ` lang="en-US"` +
    (r.size ? ` sz="${Math.round(r.size * 100)}"` : "") +
    (r.bold ? ` b="1"` : "") +
    (r.italic ? ` i="1"` : "") +
    (r.underline ? ` u="sng"` : "") +
    (r.strike ? ` strike="sngStrike"` : "")
  const inner =
    (r.color ? `<a:solidFill>${colorXml(r.color)}</a:solidFill>` : "") +
    (r.font ? `<a:latin typeface="${esc(r.font)}"/><a:cs typeface="${esc(r.font)}"/>` : "")
  return inner ? `<a:rPr${attrs}>${inner}</a:rPr>` : `<a:rPr${attrs}/>`
}

export function textXml(t: ShapeText | undefined, wordArt?: SheetShape["wordArt"], mathAsText = false): string {
  const ins = t?.insets
  const body =
    `<a:bodyPr vertOverflow="clip" horzOverflow="clip"` +
    (t?.vertical && t.vertical !== "horz" ? ` vert="${t.vertical}"` : "") +
    (t?.wrap === false ? ` wrap="none"` : ` wrap="square"`) +
    (ins
      ? ` lIns="${Math.round(ins.l * EMU)}" tIns="${Math.round(ins.t * EMU)}" rIns="${Math.round(ins.r * EMU)}" bIns="${Math.round(ins.b * EMU)}"`
      : "") +
    ` rtlCol="0" anchor="${t?.anchor ?? "t"}"${wordArt ? ` fromWordArt="1"` : ""}>` +
    (t?.autoFit ? "<a:spAutoFit/>" : "<a:noAutofit/>") +
    `</a:bodyPr><a:lstStyle/>`
  const paras = (t?.paragraphs.length ? t.paragraphs : [{ runs: [] }]).map((p) => {
    const ppr = p.align ? `<a:pPr algn="${p.align}"/>` : ""
    if (p.math) {
      const look = p.runs[0]
      const end = runXml({ ...look, text: "" })
        .replace("<a:rPr", "<a:endParaRPr")
        .replace("</a:rPr>", "</a:endParaRPr>")
      if (mathAsText) {
        const lines = mathFallbackText(p.math).map(
          (text) => `<a:r>${runXml({ ...look, text })}<a:t>${esc(text)}</a:t></a:r>`,
        )
        return `<a:p>${ppr}${lines.join("<a:br/>")}${end}</a:p>`
      }
      return `<a:p>${ppr || "<a:pPr/>"}${mathParaXml(p.math, look, runXml)}${end}</a:p>`
    }
    const effects = wordArt ? wordArtRunXml(wordArt) : ""
    const runs = p.runs
      .filter((r) => r.text)
      .map((r) => {
        const rpr = runXml(r)
        return `<a:r>${effects ? rpr.replace(/(\/>|>)/, (m) => (m === "/>" ? `>${effects}</a:rPr>` : `>${effects}`)) : rpr}<a:t>${esc(r.text)}</a:t></a:r>`
      })
      .join("")
    const end = p.runs[0]
      ? runXml(p.runs[0]).replace("<a:rPr", "<a:endParaRPr").replace("</a:rPr>", "</a:endParaRPr>")
      : `<a:endParaRPr lang="en-US"/>`
    return `<a:p>${ppr}${runs}${end}</a:p>`
  })
  return `<xdr:txBody>${body}${paras.join("")}</xdr:txBody>`
}

/** WordArt's letter fill and outline, put first in each run's properties. */
function wordArtRunXml(w: NonNullable<SheetShape["wordArt"]>): string {
  const outline = w.effects?.outline ? lineXml(w.effects.outline) : ""
  const fill = w.fill ? fillXml(w.fill) : ""
  const effects: string[] = []
  if (w.effects?.glow)
    effects.push(`<a:glow rad="63500">${colorXml({ ...w.effects.glow, alpha: w.effects.glow.alpha ?? 0.6 })}</a:glow>`)
  if (w.effects?.shadow)
    effects.push(
      `<a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw>`,
    )
  if (w.effects?.reflection)
    effects.push(
      `<a:reflection blurRad="6350" stA="55000" endA="300" endPos="45500" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`,
    )
  return outline + fill + (effects.length ? `<a:effectLst>${effects.join("")}</a:effectLst>` : "")
}

// ── Objects ──────────────────────────────────────────────────────────────────────────────────────────────────

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")

/** What reading a drawing object needs from around it. */
export type ObjectReader = {
  nextId(): number
  /** Each object's id in the file → its id in Cells (connectors name the shapes they're glued to by file id). */
  ids?: Map<number, number>
  /** The drawing as read (an equation keeps the math elements Cells doesn't model as they came). */
  doc?: XDoc
  /** A picture's data URL from its relationship id (undefined: missing). */
  image(rid: string): string | undefined
  /** A corrected picture's original (its a14:imgLayer, an HD Photo or any picture) from its relationship id. */
  layer?(rid: string): string | undefined
}

type Common = Pick<DrawingObject, "id" | "name" | "rotation" | "flipH" | "flipV" | "hidden" | "altText" | "decorative">

function readCommon(nvPr: XNode | undefined, xfrm: XNode | undefined, r: ObjectReader, decorative: boolean): Common {
  const c = kid(nvPr, "cNvPr")
  const id = r.nextId()
  const fileId = Number(at(c, "id"))
  if (r.ids && Number.isFinite(fileId)) r.ids.set(fileId, id)
  const out: Common = { id, name: at(c, "name") || `Object ${id}` }
  const descr = at(c, "descr")
  if (descr) out.altText = descr
  if (at(c, "hidden") === "1") out.hidden = true
  if (decorative) out.decorative = true
  const rot = num(at(xfrm, "rot")) / 60000
  if (rot) out.rotation = rot
  if (at(xfrm, "flipH") === "1") out.flipH = true
  if (at(xfrm, "flipV") === "1") out.flipV = true
  return out
}

/** An xfrm's box (off/ext) in pixels, for children of a group. */
function xfrmBox(xfrm: XNode | undefined): ChildBox {
  const off = kid(xfrm, "off")
  const ext = kid(xfrm, "ext")
  return {
    x: num(at(off, "x")) / EMU,
    y: num(at(off, "y")) / EMU,
    w: num(at(ext, "cx")) / EMU,
    h: num(at(ext, "cy")) / EMU,
  }
}

/** The theme colours a shape's style (lnRef, fillRef, fontRef) gives it. */
function styleColors(style: XNode | undefined) {
  const ref = (name: string) => {
    const node = kid(style, name)
    return node && num(at(node, "idx")) > 0 ? colorIn(node) : undefined
  }
  // The Office theme's line styles: idx 1, 2 and 3 are ½, 1 and 1½ points.
  const lineWidth = [0, 0.5, 1, 1.5][num(at(kid(style, "lnRef"), "idx"))] ?? 1
  return { line: ref("lnRef"), lineWidth, fill: ref("fillRef"), font: colorIn(kid(style, "fontRef")) }
}

/**
 * One drawing object (sp, cxnSp, pic or grpSp) as Cells sees it; null when Cells can't (a chart in a group, an
 * object of a kind it doesn't know). `isDecorative` reads Office's decorative mark.
 */
export function readObject(
  node: XNode,
  r: ObjectReader,
  isDecorative: (nv: XNode | undefined) => boolean,
): DrawingObject | null {
  if (node.local === "sp" || node.local === "cxnSp") {
    const nv = kid(node, node.local === "sp" ? "nvSpPr" : "nvCxnSpPr")
    const spPr = kid(node, "spPr")
    const xfrm = kid(spPr, "xfrm")
    const style = styleColors(kid(node, "style"))
    const geom = kid(spPr, "prstGeom")
    const geometry = at(geom, "prst") ?? (kid(spPr, "custGeom") ? "custom" : "rect")
    const adjust: Record<string, number> = {}
    for (const gd of kids(kid(geom, "avLst"), "gd")) {
      const m = /^val\s+(-?\d+)/.exec(at(gd, "fmla") ?? "")
      if (m) adjust[at(gd, "name") ?? "adj"] = Number(m[1])
    }
    const connector = node.local === "cxnSp" || geometry === "line" || geometry.includes("Connector")
    const textBox = at(kid(nv, "cNvSpPr"), "txBox") === "1"
    const body = kid(node, "txBody")
    const fill =
      readFill(spPr, style.fill) ?? (style.fill && !connector ? { kind: "solid", color: style.fill } : { kind: "none" })
    const line = readLine(
      kid(spPr, "ln"),
      { color: style.line ?? null, width: style.line ? style.lineWidth : 0 },
      style.line,
    )
    if (line.color && !line.width) line.width = 0.75
    const shape: SheetShape = {
      kind: "shape",
      ...readCommon(nv, xfrm, r, isDecorative(kid(nv, "cNvPr"))),
      anchor: ZERO_ANCHOR(),
      geometry,
      fill,
      line,
    }
    if (Object.keys(adjust).length) shape.adjust = adjust
    if (textBox) shape.textBox = true
    if (connector) shape.connector = true
    // Glued ends, by the file's ids for now (the drawing's reader turns them into Cells' ids).
    const cxn = kid(nv, "cNvCxnSpPr")
    for (const [tag, end] of [
      ["stCxn", "from"],
      ["endCxn", "to"],
    ] as const) {
      const g = kid(cxn, tag)
      const id = Number(at(g, "id"))
      const site = Number(at(g, "idx"))
      if (g && Number.isFinite(id) && Number.isFinite(site)) shape[end] = { id, site }
    }
    const text = readText(body, style.font, r.doc)
    if (text && (text.paragraphs.some((p) => p.runs.some((x) => x.text)) || !connector)) shape.text = text
    const pr = kid(body, "bodyPr")
    const rPr = kid(kid(kid(body, "p"), "r"), "rPr")
    // WordArt: marked so, or (as Excel 365 writes it) a shape with no fill and no style whose letters have their own
    // outline or effects.
    const letters = !kid(node, "style") && fill.kind === "none" && (kid(rPr, "ln") || kid(rPr, "effectLst"))
    if (pr && (at(pr, "fromWordArt") === "1" || kid(pr, "prstTxWarp") || letters)) {
      const effects = kid(rPr, "effectLst")
      shape.wordArt = {
        fill: readFill(rPr),
        effects: {
          ...(kid(rPr, "ln") ? { outline: readLine(kid(rPr, "ln"), { color: null, width: 0.75 }) } : {}),
          ...(kid(effects, "outerShdw") ? { shadow: true } : {}),
          ...(kid(effects, "glow") ? { glow: colorIn(kid(effects, "glow")) } : {}),
          ...(kid(effects, "reflection") ? { reflection: true } : {}),
        },
      }
    }
    if (kid(kid(spPr, "effectLst"), "outerShdw")) shape.shadow = true
    // The theme's effect, when the shape's style names one and its own properties don't set effects.
    const effectRef = num(at(kid(kid(node, "style"), "effectRef"), "idx"))
    if (effectRef >= 1 && effectRef <= 3 && !kid(spPr, "effectLst") && !shape.wordArt)
      shape.themeEffect = effectRef as 1 | 2 | 3
    return shape
  }
  if (node.local === "pic") {
    const nv = kid(node, "nvPicPr")
    const blipFill = kid(node, "blipFill")
    const blip = kid(blipFill, "blip")
    const rid = at(blip, "embed", NS_R)
    const src = rid ? r.image(rid) : undefined
    if (!src) return null
    const spPr = kid(node, "spPr")
    const pic: SheetPicture = {
      kind: "picture",
      ...readCommon(nv, kid(spPr, "xfrm"), r, isDecorative(kid(nv, "cNvPr"))),
      anchor: ZERO_ANCHOR(),
      src,
    }
    const rect = kid(blipFill, "srcRect")
    if (rect) {
      const side = (n: string) => num(at(rect, n)) / 100000
      const crop = { l: side("l"), t: side("t"), r: side("r"), b: side("b") }
      if (crop.l || crop.t || crop.r || crop.b) pic.crop = crop
    }
    // Crop to Shape: the picture cut to a preset shape.
    const prst = at(kid(spPr, "prstGeom"), "prst")
    if (prst && prst !== "rect") pic.geometry = prst
    const ln = kid(spPr, "ln")
    if (ln) {
      const line = readLine(ln, { color: null, width: 0.75 })
      if (line.color) pic.line = line
    }
    // Office's SVG pictures (icons): the SVG in an extension, the PNG in the blip for older apps.
    for (const ext of kids(kid(blip, "extLst"), "ext")) {
      if (at(ext, "uri")?.toUpperCase() !== SVG_EXT) continue
      const svgRid = at(kid(ext, "svgBlip"), "embed", NS_R)
      const svg = svgRid ? r.image(svgRid) : undefined
      if (svg) pic.svg = svg
    }
    if (pic.svg || src.startsWith("data:image/svg")) pic.icon = true
    // Corrections, Color and Transparency.
    const { adjust, layerRid } = readPictureAdjust(blip, r.doc)
    if (adjust) pic.adjust = adjust
    const original = layerRid ? r.layer?.(layerRid) : undefined
    if (original) pic.original = original
    return pic
  }
  if (node.local === "grpSp") {
    const nv = kid(node, "nvGrpSpPr")
    const xfrm = kid(kid(node, "grpSpPr"), "xfrm")
    const items: DrawingObject[] = []
    const chOff = kid(xfrm, "chOff")
    const chExt = kid(xfrm, "chExt")
    const ox = num(at(chOff, "x")) / EMU
    const oy = num(at(chOff, "y")) / EMU
    const group: SheetGroup = {
      kind: "group",
      ...readCommon(nv, xfrm, r, isDecorative(kid(nv, "cNvPr"))),
      anchor: ZERO_ANCHOR(),
      child: { w: num(at(chExt, "cx")) / EMU || 1, h: num(at(chExt, "cy")) / EMU || 1 },
      items,
    }
    for (const c of node.children) {
      if (c.local === "nvGrpSpPr" || c.local === "grpSpPr") continue
      const item = readObject(c, r, isDecorative)
      // A group with something Cells can't read (a chart, an ink stroke) is kept whole.
      if (!item) return null
      const box = childXfrm(c)
      item.box = { x: box.x - ox, y: box.y - oy, w: box.w, h: box.h }
      items.push(item)
    }
    return group
  }
  return null
}

export function childXfrm(node: XNode): ChildBox {
  const pr = node.local === "grpSp" ? kid(node, "grpSpPr") : node.local === "graphicFrame" ? node : kid(node, "spPr")
  return xfrmBox(kid(pr, "xfrm"))
}

const ZERO_ANCHOR = () => ({ from: { row: 0, col: 0, x: 0, y: 0 }, to: { row: 0, col: 0, x: 0, y: 0 } })

/** Writing objects: ids for the drawing, and a picture's relationship (an existing one when it's the same picture). */
export type ObjectWriter = {
  nextId(): number
  /** Equations as their linear text (the fallback for readers without Office Math). */
  mathAsText?: boolean
  imageRel(src: string, was?: string): string
  /** A corrected picture's original as an a14:imgLayer relationship (Office's hdphoto type). */
  layerRel?(original: string, was?: string): string
  /** A 3D model's file as a relationship (the same part when it is the one read as `was`); absent in groups. */
  modelRel?(data: Uint8Array, was?: string): string
  decorativeXml(o: DrawingObject, id: number): string
  /** Notes the id an object is written with (a connector glued to it names it so). */
  written?(cellsId: number, fileId: number): void
}

/** A glued end's shape, named by its Cells id until the drawing is written and its file id known. */
export const glueRef = (cellsId: number) => `@@cxn:${cellsId}@@`

const xfrmXml = (o: DrawingObject, box: ChildBox, group?: { w: number; h: number }) => {
  const attrs =
    (o.rotation ? ` rot="${Math.round(o.rotation * 60000)}"` : "") +
    (o.flipH ? ` flipH="1"` : "") +
    (o.flipV ? ` flipV="1"` : "")
  const e = (v: number) => Math.round(v * EMU)
  const inner = `<a:off x="${e(box.x)}" y="${e(box.y)}"/><a:ext cx="${e(box.w)}" cy="${e(box.h)}"/>`
  const child = group ? `<a:chOff x="0" y="0"/><a:chExt cx="${e(group.w)}" cy="${e(group.h)}"/>` : ""
  return `<a:xfrm${attrs}>${inner}${child}</a:xfrm>`
}

/** The theme colour a shape's effect takes (a glow in the fill's colour): its fill's, or accent 1. */
function effectColor(o: SheetShape): string {
  return (o.fill.kind === "solid" && o.fill.color.theme) || "accent1"
}

/** An object's XML (sp, cxnSp, pic or grpSp) at `box` (pixels: on the sheet, or in its group). */
export function objectXml(o: DrawingObject, box: ChildBox, w: ObjectWriter): string {
  const id = w.nextId()
  w.written?.(o.id, id)
  const descr = o.altText ? ` descr="${esc(o.altText)}"` : ""
  const hidden = o.hidden ? ` hidden="1"` : ""
  const ext = w.decorativeXml(o, id)
  const cNvPr = ext
    ? `<xdr:cNvPr id="${id}" name="${esc(o.name)}"${descr}${hidden}>${ext}</xdr:cNvPr>`
    : `<xdr:cNvPr id="${id}" name="${esc(o.name)}"${descr}${hidden}/>`
  const nv = (tag: string, inner: string) => `<xdr:${tag}>${cNvPr}${inner}</xdr:${tag}>`
  if (o.kind === "shape") {
    const geom = o.geometry === "custom" ? "rect" : o.geometry
    const av = o.adjust
      ? Object.entries(o.adjust)
          .map(([n, v]) => `<a:gd name="${n}" fmla="val ${Math.round(v)}"/>`)
          .join("")
      : ""
    const spPr =
      `<xdr:spPr>${xfrmXml(o, box)}<a:prstGeom prst="${geom}"><a:avLst>${av}</a:avLst></a:prstGeom>` +
      (o.connector ? "" : fillXml(o.fill)) +
      lineXml(o.line) +
      (o.shadow
        ? `<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>`
        : "") +
      `</xdr:spPr>`
    if (o.connector) {
      const glue =
        (o.from ? `<a:stCxn id="${glueRef(o.from.id)}" idx="${o.from.site}"/>` : "") +
        (o.to ? `<a:endCxn id="${glueRef(o.to.id)}" idx="${o.to.site}"/>` : "")
      const cxn = glue ? `<xdr:cNvCxnSpPr>${glue}</xdr:cNvCxnSpPr>` : "<xdr:cNvCxnSpPr/>"
      return `<xdr:cxnSp macro="">${nv("nvCxnSpPr", cxn)}${spPr}</xdr:cxnSp>`
    }
    const txBox = o.textBox ? ` txBox="1"` : ""
    // A theme effect: a style naming it (its fill and line are the shape's own, set in spPr; text stays dark 1).
    const style =
      o.themeEffect && !o.shadow
        ? `<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>` +
          `<a:effectRef idx="${o.themeEffect}"><a:schemeClr val="${effectColor(o)}"/></a:effectRef>` +
          `<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></xdr:style>`
        : ""
    return (
      `<xdr:sp macro="" textlink="">${nv("nvSpPr", `<xdr:cNvSpPr${txBox}/>`)}${spPr}${style}` +
      textXml(o.text, o.wordArt, w.mathAsText) +
      `</xdr:sp>`
    )
  }
  if (o.kind === "picture") {
    const rid = w.imageRel(o.src, o.source?.rels?.embed)
    // The blip's extensions: the corrections with the original (a14:imgLayer), and an icon's SVG.
    const layer =
      o.original && o.adjust && w.layerRel
        ? imgPropsExtXml(o.adjust, w.layerRel(o.original, o.source?.rels?.layer))
        : ""
    const svgExt = o.svg
      ? `<a:ext uri="${SVG_EXT}"><asvg:svgBlip xmlns:asvg="${NS_ASVG}" xmlns:r="${NS_R}" r:embed="${w.imageRel(o.svg, o.source?.rels?.svg)}"/></a:ext>`
      : ""
    const svg = adjustBlipXml(o.adjust) + (layer || svgExt ? `<a:extLst>${layer}${svgExt}</a:extLst>` : "")
    if (o.model3d && w.modelRel) {
      // Office's picture of the model, which nothing can change but the model's view.
      const locks = `noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noCrop="1"`
      const fallback =
        `<xdr:pic>${nv("nvPicPr", `<xdr:cNvPicPr><a:picLocks ${locks}/></xdr:cNvPicPr>`)}` +
        `<xdr:blipFill><a:blip xmlns:r="${NS_R}" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
        `<xdr:spPr>${xfrmXml(o, box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
      const modelRid = w.modelRel(o.model3d.data, o.source?.rels?.model)
      return model3dObjectXml(o as SheetPicture & { model3d: Model3D }, box, cNvPr, modelRid, rid, fallback)
    }
    const crop = o.crop
      ? `<a:srcRect l="${Math.round(o.crop.l * 100000)}" t="${Math.round(o.crop.t * 100000)}" r="${Math.round(o.crop.r * 100000)}" b="${Math.round(o.crop.b * 100000)}"/>`
      : ""
    return (
      `<xdr:pic>${nv("nvPicPr", `<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>`)}` +
      `<xdr:blipFill><a:blip xmlns:r="${NS_R}" r:embed="${rid}">${svg}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      `<xdr:spPr>${xfrmXml(o, box)}<a:prstGeom prst="${o.geometry ?? "rect"}"><a:avLst/></a:prstGeom>${o.line ? lineXml(o.line) : ""}</xdr:spPr></xdr:pic>`
    )
  }
  // A slicer is written by the drawing (its frame needs the sheet's box); it never sits in a group.
  if (o.kind === "slicer" || o.kind === "timeline") return ""
  const items = o.items
    .filter((i): i is DrawingObject => "kind" in i)
    .map((i) => objectXml(i, i.box ?? { x: 0, y: 0, w: o.child.w, h: o.child.h }, { ...w, modelRel: undefined }))
    .join("")
  return (
    `<xdr:grpSp>${nv("nvGrpSpPr", "<xdr:cNvGrpSpPr/>")}<xdr:grpSpPr>${xfrmXml(o, box, o.child)}</xdr:grpSpPr>` +
    `${items}</xdr:grpSp>`
  )
}

export { NS_A }

const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

/**
 * An object holding an equation goes in Excel 2010's form, with a fallback showing the equation as linear text
 * (the same shape ids in both, as Excel writes them).
 */
export function objectWithMathXml(o: DrawingObject, box: ChildBox, w: ObjectWriter): string {
  const ids: number[] = []
  const choice = objectXml(o, box, {
    ...w,
    nextId: () => {
      const id = w.nextId()
      ids.push(id)
      return id
    },
  })
  if (!choice.includes("<a14:m>")) return choice
  let i = 0
  const fallback = objectXml(o, box, { ...w, nextId: () => ids[i++]!, mathAsText: true })
  return (
    `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:a14="${NS_A14}" Requires="a14">${choice}</mc:Choice>` +
    `<mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
  )
}

/** An object as Cells knows it, to tell whether it changed since it was read (pictures by a sample of their data). */
export function objectKey(o: DrawingObject): string {
  return JSON.stringify(o, (k, v: unknown) =>
    k === "source" || k === "anchor" || k === "id"
      ? undefined
      : (k === "src" || k === "svg" || k === "original") && typeof v === "string"
        ? `${v.length}:${v.slice(-64)}`
        : v instanceof Uint8Array
          ? `${v.length}:${Array.from(v.subarray(-32)).join(",")}`
          : v,
  )
}
