import {
  CHART_CONTENT_TYPE,
  CHART_RELATIONSHIP,
  chartSpaceXml,
  readChartRefs,
  readChartSpace,
  withSheetData,
  type ChartSpec,
} from "@cells/chart"
import {
  dataUrl,
  dataUrlBytes,
  isChart,
  PICTURE_TYPES,
  readChartData,
  type ChartAnchor,
  type ChartCorner,
  type DrawingObject,
  type Sheet,
  type SheetChart,
  type SheetObject,
  type SheetPicture,
  type SignatureSetup,
  type Workbook,
} from "@cells/core"
import type { OpcRelationship } from "../opc"
import { NS_R } from "./constants"
import { at, kid, kids, parseTree, raw, startTag, type XNode } from "./dom"
import { objectKey, objectXml, readObject } from "./shapes"
import { isEmbeddedObject } from "./ole-objects"

export const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

/**
 * A sheet's drawing part: its charts, shapes, text boxes, pictures and groups, which Cells reads and writes in their
 * order, and what it can't read (chartex charts, groups with charts in them), which stays as it came.
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
  /** The pictures read, by relationship id: the same picture written again uses the same part. */
  images: Map<string, string>
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
  bytes: (partName: string | null | undefined) => { data: Uint8Array; contentType: string } | undefined = () =>
    undefined,
): DrawingSource {
  const doc = parseTree(xml)
  const source: DrawingSource = {
    partName,
    rootTag: startTag(doc, doc.root),
    kept: [],
    chartRels: new Set(),
    charts: new Map(),
    images: new Map(),
  }
  const image = (rid: string) => {
    if (source.images.has(rid)) return source.images.get(rid)
    const rel = rels.find((r) => r.id === rid && r.type === REL_IMAGE)
    const part = bytes(rel?.partName)
    if (!part) return undefined
    const url = pictureUrl(rel!.partName!, part)
    source.images.set(rid, url)
    return url
  }
  let unread = 0
  for (const anchor of doc.root.children) {
    const placed = anchorOf(sheet, anchor)
    const content = anchor.children.find((c) => OBJECTS.has(c.local))
    if (placed && content && content.local !== "graphicFrame") {
      const object = readObject(content, { nextId, image }, isDecorative)
      if (object) {
        object.anchor = placed.anchor
        if (placed.placement) object.placement = placed.placement
        const embeds: Record<string, string> = {}
        const blip = kid(kid(content, "blipFill"), "blip")
        const embed = at(blip, "embed", NS_R)
        if (embed) embeds.embed = embed
        const svg = kids(kid(blip, "extLst"), "ext")
          .map((e) => at(kid(e, "svgBlip"), "embed", NS_R))
          .find(Boolean)
        if (svg) embeds.svg = svg
        object.source = {
          xml: raw(doc, content),
          key: objectKey(object),
          part: partName,
          ...(Object.keys(embeds).length ? { rels: embeds } : {}),
        }
        sheet.objects.push(object)
        continue
      }
    }
    const rid = anchor.local === "twoCellAnchor" ? chartRelOf(anchor) : undefined
    const rel = rid ? rels.find((r) => r.id === rid && r.type === CHART_RELATIONSHIP) : undefined
    const chartXml = text(rel?.partName)
    const from = corner(kid(anchor, "from"))
    const to = corner(kid(anchor, "to"))
    const spec = chartXml ? readChartSpace(chartXml) : null
    // A stand-in Excel writes for a VML shape (an embedded object, a form control): Cells writes the VML side.
    if (!spec && /compatExt\b[^>]*\bspid="_x0000_s\d+"/.test(raw(doc, anchor))) continue
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
    const descr = at(nv, "descr")
    if (descr) chart.altText = descr
    if (isDecorative(nv)) chart.decorative = true
    sheet.objects.push(chart)
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
/** A picture part as a data URL (its type from the part, else from the file extension). */
export function pictureUrl(partName: string, part: { data: Uint8Array; contentType: string }): string {
  const ext = partName.split(".").pop()!.toLowerCase()
  const mime = part.contentType.startsWith("image/")
    ? part.contentType
    : (Object.keys(PICTURE_TYPES).find((m) => PICTURE_TYPES[m] === ext) ?? "image/png")
  return dataUrl(part.data, mime)
}

/** A picture (data URL) written as a new media part; returns the part's name. */
export function putPicture(src: string, out: Pick<DrawingWriter, "freeName" | "putBytes">): string {
  const data = dataUrlBytes(src)
  const ext = PICTURE_TYPES[data?.mime ?? ""] ?? "png"
  const media = out.freeName((n) => `/xl/media/image${n}.${ext}`)
  out.putBytes(media, data?.mime ?? "image/png", data?.bytes ?? new Uint8Array())
  return media
}

export type DrawingWriter = {
  put(name: string, contentType: string, xml: string): void
  /** A part and the parts only it uses (a chart's style, colours and embedded workbook). */
  removeTree(name: string): void
  addRel(from: string, type: string, to: string): string
  rels(part: string): OpcRelationship[]
  freeName(pattern: (n: number) => string): string
  /** A binary part (a picture). */
  putBytes(name: string, contentType: string, data: Uint8Array): void
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
  // Signature lines and embedded objects go in the VML drawing (vml.ts, ole-objects.ts), not this one.
  const drawn = sheet.objects.filter((o) => !isSignatureLine(o) && !isEmbeddedObject(o))
  if (!drawn.length && !source?.kept.length) {
    if (source) {
      for (const c of source.charts.values()) out.removeTree(c.partName)
      out.removeTree(source.partName)
    }
    return null
  }
  const partName = source?.partName ?? out.freeName((n) => `/xl/drawings/drawing${n}.xml`)
  // The drawing's relationships: the kept anchors' and pictures', then the charts' anew.
  const rels = out.rels(partName)
  for (const r of [...rels]) if (source?.chartRels.has(r.id)) rels.splice(rels.indexOf(r), 1)
  // Shape ids are unique within a drawing: new ones go above the kept anchors'.
  const keptXml = [
    ...(source?.kept ?? []),
    ...sheet.objects.flatMap((o) => (!isChart(o) && o.source?.part === partName ? [o.source.xml] : [])),
  ]
  let shapeId = Math.max(1, ...keptXml.flatMap((k) => [...k.matchAll(/\bid="(\d+)"/g)].map((m) => +m[1])))
  // A copy of an object read from the file is written anew: the file's own goes back once.
  const rawUsed = new Set<string>()
  // A picture written again uses its part; a new one gets a part of its own (one per picture data in this save).
  const newImages = new Map<string, string>()
  const writer = {
    nextId: () => ++shapeId,
    imageRel: (src: string, was?: string) => {
      if (was && source?.images.get(was) === src && rels.some((r) => r.id === was)) return was
      let rid = newImages.get(src)
      if (rid) return rid
      rid = out.addRel(partName, REL_IMAGE, putPicture(src, out))
      newImages.set(src, rid)
      return rid
    },
    decorativeXml: (o: DrawingObject) => (o.decorative ? decorativeExt() : ""),
  }
  const anchors: string[] = []
  for (const object of drawn) {
    if (!isChart(object)) {
      const box = boxOf(sheet, object.anchor)
      // Unchanged since it was read (and still in this drawing): written back exactly as it came.
      const same =
        object.source &&
        object.source.part === partName &&
        !rawUsed.has(object.source.xml) &&
        object.source.key === objectKey(object)
      if (same) rawUsed.add(object.source!.xml)
      const xml = same ? object.source!.xml : objectXml(object, box, writer)
      anchors.push(placedXml(object.anchor, object.placement, xml))
      continue
    }
    const chart = object
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

/** The drawing objects Cells reads (inside an anchor). */
const OBJECTS = new Set(["sp", "cxnSp", "pic", "grpSp", "graphicFrame"])

/** Where an anchor puts its object: two cells, one cell and a size, or a position and a size. */
function anchorOf(sheet: Sheet, anchor: XNode): { anchor: ChartAnchor; placement?: "move" | "fixed" } | null {
  if (anchor.local === "twoCellAnchor") {
    const from = corner(kid(anchor, "from"))
    const to = corner(kid(anchor, "to"))
    if (!from || !to) return null
    const editAs = at(anchor, "editAs")
    return {
      anchor: { from, to },
      placement: editAs === "oneCell" ? "move" : editAs === "absolute" ? "fixed" : undefined,
    }
  }
  const ext = kid(anchor, "ext")
  const w = Number(at(ext, "cx")) / EMU
  const h = Number(at(ext, "cy")) / EMU
  if (!Number.isFinite(w) || !Number.isFinite(h)) return null
  if (anchor.local === "oneCellAnchor") {
    const from = corner(kid(anchor, "from"))
    if (!from) return null
    const x = sheet.cols.offset(from.col) + from.x
    const y = sheet.rows.offset(from.row) + from.y
    return { anchor: { from, to: cornerAt(sheet, x + w, y + h) }, placement: "move" }
  }
  if (anchor.local === "absoluteAnchor") {
    const pos = kid(anchor, "pos")
    const x = Number(at(pos, "x")) / EMU
    const y = Number(at(pos, "y")) / EMU
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null
    return { anchor: { from: cornerAt(sheet, x, y), to: cornerAt(sheet, x + w, y + h) }, placement: "fixed" }
  }
  return null
}

function cornerAt(sheet: Sheet, px: number, py: number): ChartCorner {
  const col = sheet.cols.indexAt(Math.max(0, px))
  const row = sheet.rows.indexAt(Math.max(0, py))
  return { col, row, x: Math.max(0, px) - sheet.cols.offset(col), y: Math.max(0, py) - sheet.rows.offset(row) }
}

/** An anchor's box on the sheet, in pixels. */
/** A signature line (Insert › Signature Line), which files keep in the VML drawing. */
export const isSignatureLine = (o: SheetObject): o is SheetPicture & { signature: SignatureSetup } =>
  !isChart(o) && o.kind === "picture" && o.signature !== undefined

export function boxOf(sheet: Sheet, a: ChartAnchor) {
  const x = sheet.cols.offset(a.from.col) + a.from.x
  const y = sheet.rows.offset(a.from.row) + a.from.y
  const x1 = sheet.cols.offset(a.to.col) + a.to.x
  const y1 = sheet.rows.offset(a.to.row) + a.to.y
  return { x, y, w: Math.max(0, x1 - x), h: Math.max(0, y1 - y) }
}

const cornerXml = (c: ChartCorner) =>
  `<xdr:col>${c.col}</xdr:col><xdr:colOff>${Math.round(c.x * EMU)}</xdr:colOff>` +
  `<xdr:row>${c.row}</xdr:row><xdr:rowOff>${Math.round(c.y * EMU)}</xdr:rowOff>`

/** An object in a two-cell anchor, as Excel writes them (editAs: how it follows the cells). */
function placedXml(a: ChartAnchor, placement: "move" | "fixed" | undefined, xml: string): string {
  const editAs = placement === "move" ? ` editAs="oneCell"` : placement === "fixed" ? ` editAs="absolute"` : ""
  return (
    `<xdr:twoCellAnchor xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"${editAs}>` +
    `<xdr:from>${cornerXml(a.from)}</xdr:from><xdr:to>${cornerXml(a.to)}</xdr:to>` +
    `${xml}<xdr:clientData/></xdr:twoCellAnchor>`
  )
}

const decorativeExt = () =>
  `<a:extLst><a:ext uri="${DECORATIVE_EXT}"><adec:decorative xmlns:adec="${NS_ADEC}" val="1"/></a:ext></a:extLst>`

/** Office's "Mark as decorative" extension on a drawing object's cNvPr. */
const DECORATIVE_EXT = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
const NS_ADEC = "http://schemas.microsoft.com/office/drawing/2017/decorative"

function isDecorative(nv: XNode | undefined): boolean {
  for (const ext of kids(kid(nv, "extLst"), "ext")) {
    if (at(ext, "uri")?.toUpperCase() !== DECORATIVE_EXT) continue
    const v = at(kid(ext, "decorative"), "val")
    return v === "1" || v === "true"
  }
  return false
}

function anchorXml(chart: SheetChart, id: number, rid: string): string {
  const at = (c: ChartCorner) =>
    `<xdr:col>${c.col}</xdr:col><xdr:colOff>${Math.round(c.x * EMU)}</xdr:colOff>` +
    `<xdr:row>${c.row}</xdr:row><xdr:rowOff>${Math.round(c.y * EMU)}</xdr:rowOff>`
  const esc = (s: string) =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;").replace(/\n/g, "&#10;")
  const name = esc(chart.name)
  const descr = chart.altText ? ` descr="${esc(chart.altText)}"` : ""
  const nv = chart.decorative
    ? `<xdr:cNvPr id="${id}" name="${name}"${descr}><a:extLst><a:ext uri="${DECORATIVE_EXT}"><adec:decorative xmlns:adec="${NS_ADEC}" val="1"/></a:ext></a:extLst></xdr:cNvPr>`
    : `<xdr:cNvPr id="${id}" name="${name}"${descr}/>`
  // Prefixes declared on the anchor, so it reads the same under a kept root that names them otherwise.
  return (
    `<xdr:twoCellAnchor xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}">` +
    `<xdr:from>${at(chart.anchor.from)}</xdr:from><xdr:to>${at(chart.anchor.to)}</xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>${nv}<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${rid}"/></a:graphicData></a:graphic>` +
    `</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
  )
}
