import {
  newSparklineGroup,
  parseCell,
  storedTint,
  type SparkAxisScale,
  type SparkColor,
  type SparkColors,
  type Sparkline,
  type SparklineGroup,
} from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import { argb } from "./colors"
import { INDEXED_COLORS } from "./constants"
import { at, attrText, kid, kids, parseFragment, type XNode } from "./dom"

/**
 * Sparklines in a worksheet's extLst, as Excel 2010 and later write them: `<x14:sparklineGroups>` of
 * `<x14:sparklineGroup>` (type, options, colours, an optional date range) each holding `<x14:sparkline>`s with the
 * data (`xm:f`) and the cell (`xm:sqref`).
 */

export const X14_SPARKLINE_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"

const flag = (v: string | undefined) => v === "1" || v === "true"
const SCALES = new Set(["individual", "group", "custom"])
const COLOR_TAGS: [keyof SparkColors, string][] = [
  ["series", "colorSeries"],
  ["negative", "colorNegative"],
  ["axis", "colorAxis"],
  ["markers", "colorMarkers"],
  ["first", "colorFirst"],
  ["last", "colorLast"],
  ["high", "colorHigh"],
  ["low", "colorLow"],
]

function readColor(n: XNode | undefined): SparkColor | null {
  if (!n) return null
  const rgb = at(n, "rgb")
  if (rgb) return "#" + (rgb.length === 8 ? rgb.slice(2) : rgb).toUpperCase()
  const theme = at(n, "theme")
  if (theme !== undefined) {
    const tint = Number(at(n, "tint") ?? 0)
    return tint ? { theme: Number(theme), tint } : { theme: Number(theme) }
  }
  const indexed = at(n, "indexed")
  if (indexed !== undefined && INDEXED_COLORS[Number(indexed)]) return "#" + INDEXED_COLORS[Number(indexed)]
  return null
}

function readGroup(n: XNode): SparklineGroup | null {
  const sparklines: Sparkline[] = []
  for (const s of kids(kid(n, "sparklines"), "sparkline")) {
    const cell = parseCell(kid(s, "sqref")?.text?.trim() ?? "")
    if (cell) sparklines.push({ row: cell.row, col: cell.col, source: kid(s, "f")?.text?.trim() ?? "" })
  }
  if (!sparklines.length) return null
  const type = at(n, "type")
  const g = newSparklineGroup(type === "column" || type === "stacked" ? type : "line", sparklines)
  for (const [key, tag] of COLOR_TAGS) {
    const c = readColor(kid(n, tag))
    if (c) g.colors[key] = c
  }
  const weight = Number(at(n, "lineWeight"))
  if (Number.isFinite(weight) && weight > 0) g.lineWeight = weight
  for (const k of ["markers", "high", "low", "first", "last", "negative", "displayXAxis", "rightToLeft"] as const)
    g[k] = flag(at(n, k))
  g.showHidden = flag(at(n, "displayHidden"))
  const empty = at(n, "displayEmptyCellsAs")
  // Excel's default when the attribute is missing is "zero".
  g.emptyAs = empty === "gap" || empty === "span" ? empty : "zero"
  const min = at(n, "minAxisType")
  const max = at(n, "maxAxisType")
  if (min && SCALES.has(min)) g.minAxis = min as SparkAxisScale
  if (max && SCALES.has(max)) g.maxAxis = max as SparkAxisScale
  const manualMin = at(n, "manualMin")
  const manualMax = at(n, "manualMax")
  if (manualMin !== undefined) g.manualMin = Number(manualMin)
  if (manualMax !== undefined) g.manualMax = Number(manualMax)
  const date = kid(n, "f")?.text?.trim()
  if (flag(at(n, "dateAxis")) && date) g.dateRange = date
  const uid = at(n, "xr2:uid") ?? at(n, "uid")
  if (uid) g.uid = uid
  return g
}

/** The sparkline groups in a worksheet's extLst, and the extLst without them. */
export function readSparklineExt(source: string, rootTag: string): { groups: SparklineGroup[]; rest: string } {
  if (!source.includes(X14_SPARKLINE_URI)) return { groups: [], rest: source }
  const groups: SparklineGroup[] = []
  const doc = parseFragment(source, rootTag)
  for (const ext of kids(doc.root, "ext")) {
    if (at(ext, "uri") !== X14_SPARKLINE_URI) continue
    for (const node of kids(kid(ext, "sparklineGroups"), "sparklineGroup")) {
      const g = readGroup(node)
      if (g) groups.push(g)
    }
  }
  const rest = source.replace(/<(\w+:)?ext\b[^>]*\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}[\s\S]*?<\/\1?ext>/gi, "")
  return { groups, rest: /<(\w+:)?ext\b/.test(rest) ? rest : "" }
}

function colorXml(tag: string, c: SparkColor): string {
  if (typeof c === "string") return `<x14:${tag} rgb="${argb(c)}"/>`
  return `<x14:${tag}${attrText({ theme: c.theme, tint: c.tint ? storedTint(c.tint) : undefined })}/>`
}

const cellName = (row: number, col: number) => {
  let name = ""
  for (let c = col + 1; c > 0; c = Math.floor((c - 1) / 26)) name = String.fromCharCode(65 + ((c - 1) % 26)) + name
  return `${name}${row + 1}`
}

/** The `<ext>` with a sheet's sparkline groups ("" when it has none). */
export function sparklineExtXml(groups: readonly SparklineGroup[]): string {
  if (!groups.length) return ""
  const parts = groups.map((g) => {
    const attrs = attrText({
      type: g.type === "line" ? undefined : g.type,
      lineWeight: g.lineWeight === 0.75 ? undefined : g.lineWeight,
      dateAxis: g.dateRange ? "1" : undefined,
      manualMax: g.maxAxis === "custom" ? g.manualMax : undefined,
      manualMin: g.minAxis === "custom" ? g.manualMin : undefined,
      displayEmptyCellsAs: g.emptyAs === "zero" ? undefined : g.emptyAs,
      markers: g.markers ? "1" : undefined,
      high: g.high ? "1" : undefined,
      low: g.low ? "1" : undefined,
      first: g.first ? "1" : undefined,
      last: g.last ? "1" : undefined,
      negative: g.negative ? "1" : undefined,
      displayXAxis: g.displayXAxis ? "1" : undefined,
      displayHidden: g.showHidden ? "1" : undefined,
      minAxisType: g.minAxis === "individual" ? undefined : g.minAxis,
      maxAxisType: g.maxAxis === "individual" ? undefined : g.maxAxis,
      rightToLeft: g.rightToLeft ? "1" : undefined,
    })
    const colors = COLOR_TAGS.map(([key, tag]) => colorXml(tag, g.colors[key])).join("")
    const date = g.dateRange ? `<xm:f>${escapeXmlText(g.dateRange)}</xm:f>` : ""
    const lines = g.sparklines
      .map(
        (s) =>
          `<x14:sparkline>${s.source ? `<xm:f>${escapeXmlText(s.source)}</xm:f>` : ""}<xm:sqref>${cellName(s.row, s.col)}</xm:sqref></x14:sparkline>`,
      )
      .join("")
    return `<x14:sparklineGroup${attrs}>${colors}${date}<x14:sparklines>${lines}</x14:sparklines></x14:sparklineGroup>`
  })
  return (
    `<ext uri="${X14_SPARKLINE_URI}" xmlns:x14="${NS_X14}">` +
    `<x14:sparklineGroups xmlns:xm="${NS_XM}">${parts.join("")}</x14:sparklineGroups></ext>`
  )
}
