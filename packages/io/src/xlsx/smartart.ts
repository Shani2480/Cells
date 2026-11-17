import type { DrawingObject, SheetGroup, SmartArt, SmartArtNode, SmartArtParts } from "@cells/core"
import type { OpcRelationship } from "../opc"
import { NS_R } from "./constants"
import { at, kid, kids, parseTree, textOf, type XNode } from "./dom"
import { childXfrm, readObject, type ObjectReader } from "./shapes"

/**
 * SmartArt in an xlsx: a graphic frame naming four parts (data: the text and how its points connect; the layout
 * definition; the quick style; the colours) and, by the data's extension, Office's drawing of it as laid out
 * (`dsp:drawing`). Cells shows that drawing as a group of shapes and keeps the parts while the SmartArt is unchanged.
 */

export const NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
export const NS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
export const REL_DGM = {
  data: `${REL}diagramData`,
  layout: `${REL}diagramLayout`,
  quickStyle: `${REL}diagramQuickStyle`,
  colors: `${REL}diagramColors`,
  drawing: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
}
const CT = "application/vnd.openxmlformats-officedocument.drawingml."
export const CT_DGM = {
  data: `${CT}diagramData+xml`,
  layout: `${CT}diagramLayout+xml`,
  quickStyle: `${CT}diagramStyle+xml`,
  colors: `${CT}diagramColors+xml`,
  drawing: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
}
const EMU = 9525

/** The `dgm:relIds` of an anchor's graphic frame, when it holds SmartArt. */
export function smartArtRelIds(frame: XNode | undefined): XNode | undefined {
  const data = kid(kid(frame, "graphic"), "graphicData")
  return at(data, "uri") === NS_DGM ? kid(data, "relIds") : undefined
}

/** The text tree of a data part: the document point's children in order, and theirs. */
export function readSmartArtNodes(data: XNode): SmartArtNode[] {
  const points = new Map<string, XNode>()
  for (const p of kids(kid(data, "ptLst"), "pt")) points.set(at(p, "modelId") ?? "", p)
  const children = new Map<string, { ord: number; id: string }[]>()
  for (const c of kids(kid(data, "cxnLst"), "cxn")) {
    const type = at(c, "type") ?? "parOf"
    if (type !== "parOf") continue
    const src = at(c, "srcId") ?? ""
    const list = children.get(src) ?? []
    list.push({ ord: Number(at(c, "srcOrd") ?? 0), id: at(c, "destId") ?? "" })
    children.set(src, list)
  }
  const isNode = (p: XNode | undefined) => !!p && (at(p, "type") ?? "node") === "node"
  const tree = (id: string, depth: number): SmartArtNode[] =>
    depth > 32
      ? []
      : (children.get(id) ?? [])
          .sort((a, b) => a.ord - b.ord)
          .filter((c) => isNode(points.get(c.id)))
          .map((c) => {
            const t = kid(points.get(c.id), "t")
            const text = kids(t, "p")
              .map((p) => textOf(p))
              .join("\n")
            return { text, children: tree(c.id, depth + 1) }
          })
  const doc = [...points.values()].find((p) => at(p, "type") === "doc")
  return doc ? tree(at(doc, "modelId") ?? "", 0) : []
}

/**
 * A SmartArt frame read: a group of the shapes Office drew for it (in pixels within `size`, the frame), its text
 * and its parts. Null when its data part is missing.
 */
export function readSmartArt(
  relIds: XNode,
  frame: XNode,
  size: { w: number; h: number },
  rels: readonly OpcRelationship[],
  text: (partName: string | null | undefined) => string | undefined,
  r: ObjectReader,
  isDecorative: (nv: XNode | undefined) => boolean,
): SheetGroup | null {
  const part = (attr: string) => rels.find((x) => x.id === at(relIds, attr, NS_R))?.partName
  const dataXml = text(part("dm"))
  if (!dataXml) return null
  const data = parseTree(dataXml).root
  const doc = kids(kid(data, "ptLst"), "pt").find((p) => at(p, "type") === "doc")
  const pr = kid(doc, "prSet")
  const layoutXml = text(part("lo")) ?? ""
  const layoutId = /\buniqueId="([^"]*)"/.exec(layoutXml)?.[1] ?? at(pr, "loTypeId") ?? ""
  // Office's drawing: named by the data's extension, its relationship from the sheet's drawing.
  const drawingRel = /<dsp:dataModelExt\b[^>]*\brelId="([^"]+)"/.exec(dataXml)?.[1]
  const drawingXml = text(rels.find((x) => x.id === drawingRel)?.partName)
  const parts: SmartArtParts = {
    data: dataXml,
    layout: layoutXml,
    quickStyle: text(part("qs")) ?? "",
    colors: text(part("cs")) ?? "",
    ...(drawingXml ? { drawing: drawingXml } : {}),
  }
  const items: DrawingObject[] = []
  if (drawingXml) {
    const tree = kid(parseTree(drawingXml).root, "spTree")
    for (const c of tree?.children ?? []) {
      if (c.local === "nvGrpSpPr" || c.local === "grpSpPr") continue
      const item = readObject(c, r, () => false)
      if (!item) continue
      item.box = childXfrm(c)
      items.push(item)
    }
  }
  const nv = kid(frame, "nvGraphicFramePr")
  const c = kid(nv, "cNvPr")
  const id = r.nextId()
  const smartArt: SmartArt = {
    layout: layoutId,
    nodes: readSmartArtNodes(data),
    parts,
  }
  const qs = at(pr, "qsTypeId")
  const cs = at(pr, "csTypeId")
  if (qs) smartArt.quickStyle = qs
  if (cs) smartArt.colors = cs
  const group: SheetGroup = {
    kind: "group",
    id,
    name: at(c, "name") || `Diagram ${id}`,
    anchor: { from: { row: 0, col: 0, x: 0, y: 0 }, to: { row: 0, col: 0, x: 0, y: 0 } },
    child: { w: size.w || 1, h: size.h || 1 },
    items,
    smartArt,
  }
  const descr = at(c, "descr")
  if (descr) group.altText = descr
  if (at(c, "hidden") === "1") group.hidden = true
  if (isDecorative(nv)) group.decorative = true
  return group
}

/** The parts a SmartArt is written with, and the relationship ids they get from the sheet's drawing. */
export type SmartArtWriter = {
  /** Writes a part and gives its relationship id from the drawing. */
  part(kind: keyof SmartArtParts, xml: string): string
}

/** A SmartArt's graphic frame, its parts written anew (a copy, or one changed in Cells). */
export function smartArtFrameXml(parts: SmartArtParts, cNvPr: string, w: SmartArtWriter): string {
  // Office's drawing first: the data names it by its relationship id.
  const drawing = parts.drawing ? w.part("drawing", parts.drawing) : null
  const data = drawing
    ? parts.data.replace(/(<dsp:dataModelExt\b[^>]*\brelId=")[^"]*"/, `$1${drawing}"`)
    : parts.data.replace(
        /<dgm:extLst>\s*<a:ext uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2008\/diagram">[\s\S]*?<\/a:ext>\s*<\/dgm:extLst>/,
        "",
      )
  const dm = w.part("data", data)
  const lo = w.part("layout", parts.layout)
  const qs = w.part("quickStyle", parts.quickStyle)
  const cs = w.part("colors", parts.colors)
  return (
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>${cNvPr}<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_DGM}">` +
    `<dgm:relIds xmlns:dgm="${NS_DGM}" xmlns:r="${NS_R}" r:dm="${dm}" r:lo="${lo}" r:qs="${qs}" r:cs="${cs}"/>` +
    `</a:graphicData></a:graphic></xdr:graphicFrame>`
  )
}

export { EMU as SMARTART_EMU }
