import {
  CHART_CONTENT_TYPE,
  CHART_RELATIONSHIP,
  chartSpaceXml,
  readChartRefs,
  readChartSpace,
  withSheetData,
  type ChartSpec,
} from "@cells/chart"
import { readChartData, type ChartCorner, type Sheet, type SheetChart, type Workbook } from "@cells/core"
import type { OpcRelationship } from "../opc"
import { NS_R } from "./constants"
import { at, kid, parseTree, raw, startTag, type XNode } from "./dom"

/**
 * A sheet's drawing part: the charts on it, which Cells reads and writes, and everything else (pictures, shapes,
 * charts it can't read), which stays as it came.
 */

export const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
export const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

/** Drawing offsets are in EMUs: 9,525 to a pixel at 96 dpi. */
const EMU = 9525

/** What a save needs of a drawing read from a file. */
export type DrawingSource = {
  partName: string
  /** The root's start tag as read (its namespaces serve the kept anchors). */
  rootTag: string
  /** Anchors Cells doesn't manage, as read. */
  kept: string[]
  /** The relationships of the charts Cells read: dropped, and written anew for the charts still there. */
  chartRels: Set<string>
  /**
   * Each chart read, by id: its part, how it was when read and the numbers Excel cached in it. Unchanged and with
   * the same numbers in its cells, the part goes back as Excel wrote it.
   */
  charts: Map<number, { partName: string; print: string; cache: string }>
}

const print = (chart: SheetChart) => JSON.stringify({ spec: chart.spec, refs: chart.refs })
const cacheOf = (spec: ChartSpec) => JSON.stringify([spec.categories, spec.series.map((s) => s.values)])

function corner(node: XNode | undefined): ChartCorner | null {
  if (!node) return null
  const n = (local: string) => Number(kid(node, local)?.text ?? NaN)
  const c = { col: n("col"), x: n("colOff") / EMU, row: n("row"), y: n("rowOff") / EMU }
  return Object.values(c).every(Number.isFinite) ? c : null
}

/** The chart's relationship id, when an anchor is a chart frame (not chartex, which sits in mc:AlternateContent). */
function chartRelOf(anchor: XNode): string | undefined {
  const graphic = kid(kid(kid(anchor, "graphicFrame"), "graphic"), "graphicData")
  const chart = graphic?.children.find((c) => c.local === "chart" && c.uri === NS_C)
  return chart ? at(chart, "id", NS_R) : undefined
}

/**
 * The charts of a sheet's drawing, onto the sheet; the rest is kept for the save. `text` reads a part,
 * `nextId` numbers the charts across the workbook.
 */
export function readDrawing(
  partName: string,
  xml: string,
  rels: readonly OpcRelationship[],
  sheet: Sheet,
  text: (partName: string | null | undefined) => string | undefined,
  nextId: () => number,
  warnings: string[],
): DrawingSource {
  const doc = parseTree(xml)
  const source: DrawingSource = {
    partName,
    rootTag: startTag(doc, doc.root),
    kept: [],
    chartRels: new Set(),
    charts: new Map(),
  }
  let unread = 0
  for (const anchor of doc.root.children) {
    const rid = anchor.local === "twoCellAnchor" ? chartRelOf(anchor) : undefined
    const rel = rid ? rels.find((r) => r.id === rid && r.type === CHART_RELATIONSHIP) : undefined
    const chartXml = text(rel?.partName)
    const from = corner(kid(anchor, "from"))
    const to = corner(kid(anchor, "to"))
    const spec = chartXml ? readChartSpace(chartXml) : null
    if (!rel?.partName || !chartXml || !spec || !from || !to) {
      if (rid || anchor.local === "AlternateContent") unread++
      source.kept.push(raw(doc, anchor))
      continue
    }
    const refs = readChartRefs(chartXml)
    const nv = kid(kid(kid(anchor, "graphicFrame"), "nvGraphicFramePr"), "cNvPr")
    const chart: SheetChart = {
      id: nextId(),
      name: at(nv, "name") ?? `Chart ${sheet.charts.length + 1}`,
      anchor: { from, to },
      // Read from cells: the numbers in the part are only Excel's cache of them.
      spec: (refs ? keepLook(spec) : spec) as unknown as Record<string, unknown>,
      refs: refs ? { categories: refs.categories, series: refs.series.map((s) => ({ ...s })) } : null,
    }
    sheet.charts.push(chart)
    source.chartRels.add(rid!)
    source.charts.set(chart.id, { partName: rel.partName, print: print(chart), cache: cacheOf(spec) })
  }
  if (unread)
    warnings.push(
      `${unread === 1 ? "A chart" : `${unread} charts`} on ${sheet.name} can't be shown or changed in Cells; ${unread === 1 ? "it is" : "they are"} kept as ${unread === 1 ? "it was" : "they were"} when you save.`,
    )
  return source
}

const keepLook = (spec: ChartSpec): ChartSpec => ({
  ...spec,
  categories: [],
  series: spec.series.map((s) => ({ ...s, values: [], texts: undefined })),
})

/** The package operations a save gives the drawing writer. */
export type DrawingWriter = {
  put(name: string, contentType: string, xml: string): void
  /** A part and the parts only it uses (a chart's style, colours and embedded workbook). */
  removeTree(name: string): void
  addRel(from: string, type: string, to: string): string
  rels(part: string): OpcRelationship[]
  freeName(pattern: (n: number) => string): string
}

/**
 * Writes a sheet's drawing (charts, and what was kept from the file) and returns its relationship id from the sheet,
 * or null when the sheet has nothing to draw.
 */
export function writeDrawing(
  wb: Workbook,
  sheet: Sheet,
  sheetPart: string,
  source: DrawingSource | undefined,
  out: DrawingWriter,
): string | null {
  const used = new Set<number>()
  if (!sheet.charts.length && !source?.kept.length) {
    if (source) {
      for (const c of source.charts.values()) out.removeTree(c.partName)
      out.removeTree(source.partName)
    }
    return null
  }
  const partName = source?.partName ?? out.freeName((n) => `/xl/drawings/drawing${n}.xml`)
  // The drawing's relationships: the kept anchors' (pictures, unread charts), then the charts' anew.
  const rels = out.rels(partName)
  for (const r of [...rels]) if (source?.chartRels.has(r.id)) rels.splice(rels.indexOf(r), 1)
  // Shape ids are unique within a drawing: new ones go above the kept anchors'.
  let shapeId = Math.max(1, ...(source?.kept ?? []).flatMap((k) => [...k.matchAll(/\bid="(\d+)"/g)].map((m) => +m[1])))
  const anchors: string[] = []
  for (const chart of sheet.charts) {
    const was = source?.charts.get(chart.id)
    const spec = chart.refs
      ? withSheetData(chart.spec as Partial<ChartSpec>, readChartData(wb.sheets, wb.styles, chart.refs))
      : (chart.spec as unknown as ChartSpec)
    let part: string
    if (was && was.print === print(chart) && was.cache === cacheOf(spec)) part = was.partName
    else {
      if (was) out.removeTree(was.partName)
      part = was?.partName ?? out.freeName((n) => `/xl/charts/chart${n}.xml`)
      // The package writer adds the XML declaration.
      const xml = chartSpaceXml(spec, undefined, chart.refs ?? undefined).replace(/^<\?xml[^>]*\?>/, "")
      out.put(part, CHART_CONTENT_TYPE, xml)
    }
    if (was) used.add(chart.id)
    const rid = out.addRel(partName, CHART_RELATIONSHIP, part)
    anchors.push(anchorXml(chart, ++shapeId, rid))
  }
  if (source) for (const [id, c] of source.charts) if (!used.has(id)) out.removeTree(c.partName)
  const open = source?.rootTag ?? `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"`
  const name = /^<([^\s>/]+)/.exec(open)![1]
  out.put(partName, CT_DRAWING, `${open}>${source?.kept.join("") ?? ""}${anchors.join("")}</${name}>`)
  return out.addRel(sheetPart, REL_DRAWING, partName)
}

function anchorXml(chart: SheetChart, id: number, rid: string): string {
  const at = (c: ChartCorner) =>
    `<xdr:col>${c.col}</xdr:col><xdr:colOff>${Math.round(c.x * EMU)}</xdr:colOff>` +
    `<xdr:row>${c.row}</xdr:row><xdr:rowOff>${Math.round(c.y * EMU)}</xdr:rowOff>`
  const name = chart.name.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;")
  // Prefixes declared on the anchor, so it reads the same under a kept root that names them otherwise.
  return (
    `<xdr:twoCellAnchor xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}">` +
    `<xdr:from>${at(chart.anchor.from)}</xdr:from><xdr:to>${at(chart.anchor.to)}</xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${rid}"/></a:graphicData></a:graphic>` +
    `</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
  )
}
