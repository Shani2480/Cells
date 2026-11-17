import type { Model3D, SheetPicture } from "@cells/core"
import { at, kid, raw, type XNode } from "./dom"
import { NS_R } from "./constants"

/**
 * Insert › 3D Models in an xlsx: a graphic frame holding `am3d:model3d` (the .glb by relationship, the view, and a
 * picture of it), in mc:AlternateContent with a plain picture for readers without 3D models. Both sides share
 * the shape id, as Excel writes them.
 */

export const NS_AM3D = "http://schemas.microsoft.com/office/drawing/2017/model3d"
export const REL_MODEL3D = "http://schemas.microsoft.com/office/2017/06/relationships/model3d"
export const CT_GLB = "model/gltf.binary"
const EMU = 9525

/** The `am3d:model3d` of an anchor's Choice, when it holds one. */
export function model3dNode(choice: XNode | undefined): XNode | undefined {
  const frame = choice?.children.find((c) => c.local === "graphicFrame")
  const data = kid(kid(frame, "graphic"), "graphicData")
  return data?.children.find((c) => c.local === "model3d")
}

/** The model of a `am3d:model3d` (its data from its relationship); null without the model's file. */
export function readModel3d(
  doc: Parameters<typeof raw>[0],
  node: XNode,
  model: (rid: string) => Uint8Array | undefined,
): { model: Model3D; rid: string } | null {
  const rid = at(node, "embed", NS_R)
  const data = rid ? model(rid) : undefined
  if (!rid || !data) return null
  const rot = kid(kid(node, "trans"), "rot")
  const angle = (n: string) => Number(at(rot, n) ?? 0) / 60_000
  const viewport = Number(at(kid(node, "objViewport"), "viewportSz") ?? 0) / EMU
  return {
    model: { data, rot: { x: angle("ax"), y: angle("ay"), z: angle("az") }, viewport, xml: raw(doc, node) },
    rid,
  }
}

/** The model's `am3d:model3d` as written: its view, the model's and picture's relationships and its size. */
export function model3dElementXml(m: Model3D, size: { w: number; h: number }, modelRid: string, imageRid: string) {
  const e = (v: number) => Math.round(v * EMU)
  const a = (v: number) => Math.round(v * 60_000)
  const rot =
    (m.rot.x ? ` ax="${a(m.rot.x)}"` : "") +
    (m.rot.y ? ` ay="${a(m.rot.y)}"` : "") +
    (m.rot.z ? ` az="${a(m.rot.z)}"` : "")
  return (m.xml ?? "")
    .replace(/^(<am3d:model3d\b[^>]*?\br:embed=")[^"]*"/, `$1${modelRid}"`)
    .replace(/(<am3d:spPr>\s*<a:xfrm\b[^>]*>[\s\S]*?<a:ext )cx="\d+" cy="\d+"/, `$1cx="${e(size.w)}" cy="${e(size.h)}"`)
    .replace(/<am3d:rot\b[^>]*\/>|<am3d:rot\b[^>]*>[\s\S]*?<\/am3d:rot>/, `<am3d:rot${rot}/>`)
    .replace(/(<am3d:blip\b[^>]*?\br:embed=")[^"]*"/, `$1${imageRid}"`)
    .replace(/(<am3d:objViewport\b[^>]*?\bviewportSz=")\d*"/, `$1${e(m.viewport)}"`)
}

/**
 * A 3D model's anchor content: the graphic frame for Office, and `fallback` (the model's picture, as a `pic`
 * written with the same id) for everyone else.
 */
export function model3dObjectXml(
  o: SheetPicture & { model3d: Model3D },
  size: { w: number; h: number },
  cNvPr: string,
  modelRid: string,
  imageRid: string,
  fallback: string,
) {
  return (
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice xmlns:am3d="${NS_AM3D}" Requires="am3d"><xdr:graphicFrame macro="">` +
    `<xdr:nvGraphicFramePr>${cNvPr}<xdr:cNvGraphicFramePr><a:graphicFrameLocks/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS_AM3D}">${model3dElementXml(o.model3d, size, modelRid, imageRid)}</a:graphicData></a:graphic>` +
    `</xdr:graphicFrame></mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
  )
}
