import type { ChartAnchor, EmbeddedFile, SheetObject, SheetPicture } from "@cells/core"
import { isChart } from "@cells/core"
import { escapeXml } from "@cells/ooxml"
import { packageObject, unpackObject } from "../ole"

/**
 * Insert › Object's embedded files in a worksheet, as Excel writes them: an `oleObject` in the sheet's `oleObjects`
 * (its class, its picture and where it sits), a picture shape in the VML drawing with the same shape id, and the
 * embedding part (a Package compound file, or an Office document as it is).
 */

export const REL_OLE_OBJECT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
export const REL_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
export const CT_OLE_OBJECT = "application/vnd.openxmlformats-officedocument.oleObject"
const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const EMU = 9525

/** Office documents embedded as themselves: extension → [class, content type, part name]. */
const OFFICE_DOCS: Record<string, [string, string, string]> = {
  docx: [
    "Word.Document.12",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "Microsoft_Word_Document",
  ],
  xlsx: [
    "Excel.Sheet.12",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Microsoft_Excel_Worksheet",
  ],
  pptx: [
    "PowerPoint.Show.12",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "Microsoft_PowerPoint_Presentation",
  ],
}

export const isEmbeddedObject = (o: SheetObject): o is SheetPicture & { embed: EmbeddedFile } =>
  !isChart(o) && o.kind === "picture" && o.embed !== undefined

const attr = (tag: string, name: string): string | undefined => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag)
  return m?.[1]
    ?.replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
}

const corner = (xml: string | undefined) => {
  const num = (tag: string) => Number(new RegExp(`<(?:\\w+:)?${tag}>(-?\\d+)<`).exec(xml ?? "")?.[1] ?? 0)
  return { col: num("col"), x: num("colOff") / EMU, row: num("row"), y: num("rowOff") / EMU }
}

export type OleObjectRead = {
  shapeId: number
  progId: string
  /** The embedding's relationship from the sheet. */
  relId: string | undefined
  /** The picture's relationship from the sheet (objectPr), when there is one. */
  imageRelId: string | undefined
  anchor: ChartAnchor | undefined
  placement: SheetPicture["placement"]
}

/** The embedded objects in a sheet's `oleObjects` element (Excel 2010's form preferred to the fallback). */
export function readOleObjects(xml: string): OleObjectRead[] {
  // The Choice and its Fallback describe one object: read the Choice where there is one.
  const plain = xml.replace(/<(\w+:)?Fallback\b[\s\S]*?<\/\1?Fallback>/g, "")
  const out: OleObjectRead[] = []
  for (const m of plain.matchAll(/<oleObject\b([^>]*?)(\/>|>([\s\S]*?)<\/oleObject>)/g)) {
    const tag = m[1]!
    const body = m[3] ?? ""
    const objectPr = /<objectPr\b[^>]*>/.exec(body)?.[0]
    const anchorTag = /<anchor\b[^>]*>/.exec(body)?.[0] ?? ""
    const from = /<from>([\s\S]*?)<\/from>/.exec(body)?.[1]
    const to = /<to>([\s\S]*?)<\/to>/.exec(body)?.[1]
    out.push({
      shapeId: Number(attr(tag, "shapeId") ?? 0),
      progId: attr(tag, "progId") ?? "Package",
      relId: attr(tag, "r:id"),
      imageRelId: objectPr ? attr(objectPr, "r:id") : undefined,
      anchor: from && to ? { from: corner(from), to: corner(to) } : undefined,
      placement:
        attr(anchorTag, "moveWithCells") === "1"
          ? attr(anchorTag, "sizeWithCells") === "1"
            ? undefined
            : "move"
          : anchorTag
            ? "fixed"
            : "move",
    })
  }
  return out
}

/** An object's anchor from its VML shape (old files without `objectPr`). */
export function vmlAnchor(vml: string | undefined, shapeId: number): ChartAnchor | undefined {
  const shape = vml && new RegExp(`<v:shape\\b[^>]*id="_x0000_s${shapeId}"[\\s\\S]*?</v:shape>`).exec(vml)?.[0]
  const a =
    shape &&
    /<x:Anchor>\s*([\d\s,]+)<\/x:Anchor>/
      .exec(shape)?.[1]
      ?.split(",")
      .map((v) => +v.trim())
  if (!a || a.length < 8) return undefined
  return { from: { col: a[0]!, x: a[1]!, row: a[2]!, y: a[3]! }, to: { col: a[4]!, x: a[5]!, row: a[6]!, y: a[7]! } }
}

/** The picture relationship of an object's VML shape (its `v:imagedata`). */
export function vmlImageRel(vml: string | undefined, shapeId: number): string | undefined {
  const shape = vml && new RegExp(`<v:shape\\b[^>]*id="_x0000_s${shapeId}"[\\s\\S]*?</v:shape>`).exec(vml)?.[0]
  const data = shape && /<v:imagedata\b[^>]*>/.exec(shape)?.[0]
  return data ? attr(data, "o:relid") : undefined
}

/** What's embedded: the file in a Package, an Office document as it is, or anything else kept as it came. */
export function embedFrom(
  progId: string,
  part: { name: string; data: Uint8Array; contentType: string } | undefined,
  relType: string,
): EmbeddedFile {
  if (!part) return { name: progId, progId }
  const file = part.name.split("/").pop() ?? "Object"
  const ext = /\.([^.]+)$/.exec(file)?.[1]?.toLowerCase() ?? "bin"
  if (/^Package$|^Packager/i.test(progId)) {
    const inner = unpackObject(part.data)
    if (inner) return { name: inner.name, data: inner.data, progId: "Package" }
  }
  if (OFFICE_DOCS[ext] && relType === REL_PACKAGE) return { name: file, data: part.data, progId }
  return {
    name: file.replace(/\.[^.]+$/, ""),
    progId,
    raw: { bytes: part.data, ext, contentType: part.contentType, relType },
  }
}

export type OleWriter = {
  /** Writes a part and returns its relationship id from the sheet. */
  embed(name: (n: number) => string, contentType: string, relType: string, data: Uint8Array): string
}

/** The embedding part for an object, written through `out`; returns its relationship id. */
export function writeEmbedding(e: EmbeddedFile, out: OleWriter): { relId: string; progId: string } {
  if (e.raw) {
    const ext = e.raw.ext
    return {
      relId: out.embed((n) => `/xl/embeddings/oleObject${n}.${ext}`, e.raw.contentType, e.raw.relType, e.raw.bytes),
      progId: e.progId,
    }
  }
  const ext = /\.([^.]+)$/.exec(e.name)?.[1]?.toLowerCase() ?? ""
  const office = OFFICE_DOCS[ext]
  if (office && e.data && e.progId !== "Package") {
    const [progId, type, base] = office
    return {
      relId: out.embed((n) => `/xl/embeddings/${base}${n === 1 ? "" : n}.${ext}`, type, REL_PACKAGE, e.data),
      progId,
    }
  }
  return {
    relId: out.embed(
      (n) => `/xl/embeddings/oleObject${n}.bin`,
      CT_OLE_OBJECT,
      REL_OLE_OBJECT,
      packageObject(e.name, e.data ?? new Uint8Array()),
    ),
    progId: "Package",
  }
}

const cornerXml = (c: ChartAnchor["from"]) =>
  `<xdr:col>${c.col}</xdr:col><xdr:colOff>${Math.round(c.x * EMU)}</xdr:colOff>` +
  `<xdr:row>${c.row}</xdr:row><xdr:rowOff>${Math.round(c.y * EMU)}</xdr:rowOff>`

/** One `oleObject`, in Excel 2010's form with the fallback older Excel reads. */
export function oleObjectXml(
  o: SheetPicture,
  ids: { shapeId: number; relId: string; imageRelId: string; progId: string },
): string {
  const head = `<oleObject progId="${escapeXml(ids.progId)}" dvAspect="DVASPECT_ICON" shapeId="${ids.shapeId}" r:id="${ids.relId}"`
  const move = o.placement === "fixed" ? "" : ` moveWithCells="1"`
  const size = o.placement === undefined ? ` sizeWithCells="1"` : ""
  return (
    `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:x14="${NS_X14}" Requires="x14">${head}>` +
    `<objectPr defaultSize="0" autoPict="0" r:id="${ids.imageRelId}"><anchor${move}${size}>` +
    `<from>${cornerXml(o.anchor.from)}</from><to>${cornerXml(o.anchor.to)}</to></anchor></objectPr></oleObject>` +
    `</mc:Choice><mc:Fallback>${head}/></mc:Fallback></mc:AlternateContent>`
  )
}

/** The sheet's `oleObjects` element. */
export const oleObjectsXml = (items: string[]) =>
  items.length ? `<oleObjects xmlns:xdr="${NS_XDR}">${items.join("")}</oleObjects>` : ""

/** An object's VML shape: its picture where it sits (Excel draws objects from this). */
export function oleShapeXml(
  o: SheetPicture,
  shapeId: number,
  imageRelId: string,
  box: { x: number; y: number; w: number; h: number },
): string {
  const pt = (px: number) => `${Math.round(px * 0.75 * 100) / 100}pt`
  const a = o.anchor
  const place = o.placement === "fixed" ? "" : o.placement === "move" ? "<x:MoveWithCells/>" : "<x:SizeWithCells/>"
  return (
    `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t75" style="position:absolute;margin-left:${pt(box.x)};margin-top:${pt(box.y)};width:${pt(box.w)};height:${pt(box.h)};z-index:1"` +
    ` filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]" o:insetmode="auto">` +
    `<v:fill color2="window [65]"/><v:imagedata o:relid="${imageRelId}" o:title=""/>` +
    `<x:ClientData ObjectType="Pict">${place}` +
    `<x:Anchor>${a.from.col}, ${Math.round(a.from.x)}, ${a.from.row}, ${Math.round(a.from.y)}, ${a.to.col}, ${Math.round(a.to.x)}, ${a.to.row}, ${Math.round(a.to.y)}</x:Anchor>` +
    `<x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>`
  )
}
