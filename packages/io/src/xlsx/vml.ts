import { SIGNING_INSTRUCTIONS, type SheetPicture, type SignatureSetup } from "@cells/core"
import { escapeXml } from "@cells/ooxml"

/**
 * Pictures in a sheet's VML drawing (the part the notes use): Insert › Signature Line's lines. Office keeps a line as
 * a picture shape with an `o:signatureline` element naming the signer; its picture is a rel of the VML part.
 */

const attr = (tag: string, name: string): string | undefined => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag)
  return m ? unescapeXml(m[1]!) : undefined
}

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(+d))
    .replace(/&amp;/g, "&")
}

export type VmlSignature = Omit<SheetPicture, "id" | "name" | "src"> & {
  /** The picture's relationship in the VML part. */
  relId: string | undefined
  /** The VML shape's id ("_x0000_s1025"). */
  spid: string
}

/** The signature lines in a VML drawing, with where each sits and its picture's rel. */
export function readVmlSignatures(vml: string): VmlSignature[] {
  const out: VmlSignature[] = []
  for (const m of vml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)) {
    const body = m[2]!
    const line = /<o:signatureline\b[^>]*>/.exec(body)?.[0]
    if (!line || attr(line, "issignatureline") === "f") continue
    const anchor = /<x:Anchor>\s*([\d\s,]+)<\/x:Anchor>/
      .exec(body)?.[1]
      ?.split(",")
      .map((v) => +v.trim())
    if (!anchor || anchor.length < 8) continue
    const [c0, x0, r0, y0, c1, x1, r1, y1] = anchor as [number, number, number, number, number, number, number, number]
    const signature: SignatureSetup = {
      id: attr(line, "id") ?? "",
      provider: attr(line, "provid"),
      signer: attr(line, "o:suggestedsigner") ?? "",
      title: attr(line, "o:suggestedsigner2") ?? "",
      email: attr(line, "o:suggestedsigneremail") ?? "",
      instructions: attr(line, "o:signinginstructions") ?? SIGNING_INSTRUCTIONS,
      allowComments: attr(line, "allowcomments") === "t",
      showDate: attr(line, "showsigndate") !== "f",
    }
    const imagedata = /<v:imagedata\b[^>]*>/.exec(body)?.[0] ?? ""
    out.push({
      kind: "picture",
      anchor: { from: { col: c0, x: x0, row: r0, y: y0 }, to: { col: c1, x: x1, row: r1, y: y1 } },
      signature,
      altText: attr(m[1]!, "alt"),
      relId: attr(imagedata, "o:relid"),
      spid: attr(m[1]!, "id") ?? "",
      placement: /<x:SizeWithCells\s*\/>/.test(body)
        ? undefined
        : /<x:MoveWithCells\s*\/>/.test(body)
          ? "move"
          : "fixed",
    })
  }
  return out
}

/** The picture shape type (Office's 75) the lines use. */
export const PICTURE_SHAPETYPE =
  `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">` +
  `<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/>` +
  `<v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>` +
  `<v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/>` +
  `<v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>` +
  `<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>`

/** A signature line's VML shape. `box` is its place in points (VML's style), `relId` its picture's rel. */
export function signatureShapeXml(
  o: SheetPicture & { signature: SignatureSetup },
  shapeId: number,
  relId: string,
  box: { x: number; y: number; w: number; h: number },
): string {
  const s = o.signature
  const a = o.anchor
  const pt = (px: number) => `${Math.round(px * 0.75 * 100) / 100}pt`
  const q = (v: string) => escapeXml(v)
  const place = o.placement === "fixed" ? "" : o.placement === "move" ? "<x:MoveWithCells/>" : "<x:SizeWithCells/>"
  return (
    `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t75" alt="${q(o.altText ?? "")}"` +
    ` style="position:absolute;margin-left:${pt(box.x)};margin-top:${pt(box.y)};width:${pt(box.w)};height:${pt(box.h)};z-index:1">` +
    `<v:imagedata o:relid="${relId}" o:title=""/>` +
    `<o:lock v:ext="edit" ungrouping="t" rotation="t" cropping="t" verticies="t" text="t" grouping="t"/>` +
    `<o:signatureline v:ext="edit" id="${q(s.id)}" provid="${q(s.provider ?? "{00000000-0000-0000-0000-000000000000}")}"` +
    (s.signer ? ` o:suggestedsigner="${q(s.signer)}"` : "") +
    (s.title ? ` o:suggestedsigner2="${q(s.title)}"` : "") +
    (s.email ? ` o:suggestedsigneremail="${q(s.email)}"` : "") +
    (s.instructions ? ` o:signinginstructions="${q(s.instructions)}" signinginstructionsset="t"` : "") +
    (s.allowComments ? ` allowcomments="t"` : "") +
    (s.showDate ? "" : ` showsigndate="f"`) +
    ` issignatureline="t"/>` +
    `<x:ClientData ObjectType="Pict">${place}` +
    `<x:Anchor>${a.from.col}, ${Math.round(a.from.x)}, ${a.from.row}, ${Math.round(a.from.y)}, ${a.to.col}, ${Math.round(a.to.x)}, ${a.to.row}, ${Math.round(a.to.y)}</x:Anchor>` +
    `<x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>`
  )
}

/** A VML drawing of only pictures (a sheet with signature lines and no notes). */
export function pictureVml(sheetIndex: number, shapes: string): string {
  return (
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheetIndex + 1}"/></o:shapelayout>` +
    PICTURE_SHAPETYPE +
    shapes +
    `</xml>`
  )
}

/** The notes' VML with the pictures added after their shapes. */
export function withPictures(vml: string, shapes: string): string {
  return vml.replace(/(<\/o:shapelayout>)/, `$1${PICTURE_SHAPETYPE}`).replace(/<\/xml>\s*$/, `${shapes}</xml>`)
}

/** The highest VML shape number in a drawing (new shapes go above it). */
export function lastShapeId(vml: string): number {
  return Math.max(0, ...[...vml.matchAll(/_x0000_s(\d+)/g)].map((m) => +m[1]!))
}
