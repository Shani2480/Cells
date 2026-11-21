/**
 * Ink (Draw tab pens) in xlsx: each ink object is a content part (xdr:contentPart with Excel 2010's xdr14 properties
 * inside, as Excel writes it) whose InkML part (a customXml relationship, MS-ODRAWXML 2.1.4) holds the strokes. Trace coordinates are in 1/1000 cm; Office fits the traces'
 * extent to the content part's frame, so Cells reads them as fractions of that extent.
 */
import type { InkStroke } from "@cells/core"
import { at, kid, kids, parseTree, type XNode } from "./dom"

export const REL_INK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
export const CT_INK = "application/inkml+xml"
export const NS_XDR14 = "http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing"
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const NS_INKML = "http://www.w3.org/2003/InkML"

/** Pixels (96 per inch) to InkML units (1/1000 cm), and to EMU. */
const UNITS_PER_PX = 2540 / 96
const EMU = 9525

const esc = (v: string) => v.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;")

/** The InkML part for strokes drawn in a box of `w` × `h` pixels. */
export function inkmlXml(strokes: readonly InkStroke[], w: number, h: number): string {
  const brushes = strokes.map((s, i) => {
    const cm = (s.width / 96) * 2.54
    const size = cm.toFixed(5).replace(/0+$/, "").replace(/\.$/, "")
    return (
      `<inkml:brush xml:id="br${i}">` +
      `<inkml:brushProperty name="width" value="${size}" units="cm"/>` +
      `<inkml:brushProperty name="height" value="${size}" units="cm"/>` +
      `<inkml:brushProperty name="color" value="${esc(s.color.toUpperCase())}"/>` +
      (s.highlighter
        ? `<inkml:brushProperty name="tip" value="rectangle"/><inkml:brushProperty name="rasterOp" value="maskPen"/>`
        : "") +
      `</inkml:brush>`
    )
  })
  const traces = strokes.map((s, i) => {
    const pts: string[] = []
    for (let k = 0; k + 1 < s.points.length; k += 2)
      pts.push(`${Math.round(s.points[k] * w * UNITS_PER_PX)} ${Math.round(s.points[k + 1] * h * UNITS_PER_PX)}`)
    return `<inkml:trace contextRef="#ctx0" brushRef="#br${i}">${pts.join(", ")}</inkml:trace>`
  })
  return (
    `<inkml:ink xmlns:inkml="${NS_INKML}">` +
    `<inkml:definitions><inkml:context xml:id="ctx0"><inkml:inkSource xml:id="inkSrc0"><inkml:traceFormat>` +
    `<inkml:channel name="X" type="integer" max="32767" units="cm"/>` +
    `<inkml:channel name="Y" type="integer" max="32767" units="cm"/>` +
    `</inkml:traceFormat><inkml:channelProperties>` +
    `<inkml:channelProperty channel="X" name="resolution" value="1000" units="1/cm"/>` +
    `<inkml:channelProperty channel="Y" name="resolution" value="1000" units="1/cm"/>` +
    `</inkml:channelProperties></inkml:inkSource></inkml:context>${brushes.join("")}</inkml:definitions>` +
    `${traces.join("")}</inkml:ink>`
  )
}

/** The anchor's content: the content part, for an ink object named `name` with shape id `id`, its frame in pixels. */
export function contentPartXml(
  rid: string,
  id: number,
  name: string,
  box: { x: number; y: number; w: number; h: number },
  extra = "",
): string {
  const e = (v: number) => Math.round(v * EMU)
  return (
    `<xdr:contentPart xmlns:r="${NS_R}" r:id="${rid}">` +
    `<xdr14:nvContentPartPr xmlns:xdr14="${NS_XDR14}"><xdr14:cNvPr id="${id}" name="${esc(name)}"${extra}/>` +
    `<xdr14:cNvContentPartPr/></xdr14:nvContentPartPr><xdr14:nvPr xmlns:xdr14="${NS_XDR14}" macro=""/>` +
    `<xdr14:xfrm xmlns:xdr14="${NS_XDR14}"><a:off x="${e(box.x)}" y="${e(box.y)}"/>` +
    `<a:ext cx="${e(box.w)}" cy="${e(box.h)}"/></xdr14:xfrm></xdr:contentPart>`
  )
}

/** An anchor's content part, if it has one: its relationship id and cNvPr. */
export function contentPartOf(anchor: XNode): { rid: string; nv: XNode | undefined } | null {
  for (const choice of kids(kid(anchor, "AlternateContent"), "Choice")) {
    const part = kid(choice, "contentPart")
    const rid = part ? at(part, "id", NS_R) : undefined
    if (rid) return { rid, nv: kid(kid(part, "nvContentPartPr"), "cNvPr") }
  }
  const part = kid(anchor, "contentPart")
  const rid = part ? at(part, "id", NS_R) : undefined
  return rid ? { rid, nv: kid(kid(part, "nvContentPartPr"), "cNvPr") } : null
}

/**
 * Strokes from an InkML part, as fractions of the traces' extent (null when it has none). Brushes give each trace's
 * colour, width (in pixels) and whether it's a highlighter; traces without X and Y aren't read.
 */
export function readInkml(xml: string): InkStroke[] | null {
  const doc = parseTree(xml)
  const all: XNode[] = []
  const walk = (n: XNode) => {
    all.push(n)
    for (const c of n.children) walk(c)
  }
  walk(doc.root)
  const brushes = new Map<string, { color: string; width: number; highlighter: boolean }>()
  for (const b of all.filter((n) => n.local === "brush")) {
    const props = new Map(kids(b, "brushProperty").map((p) => [at(p, "name") ?? "", p]))
    const width = props.get("width")
    const units = at(width!, "units") ?? "cm"
    const v = Number(at(width!, "value") ?? "0.05")
    const px = (units === "mm" ? v / 10 : units === "in" ? v * 2.54 : units === "himetric" ? v / 1000 : v) * (96 / 2.54)
    const raster = at(props.get("rasterOp")!, "value") ?? ""
    const transparency = Number(at(props.get("transparency")!, "value") ?? "0")
    brushes.set(b.attrs["xml:id"] ?? at(b, "id") ?? "", {
      color: at(props.get("color")!, "value") ?? "#000000",
      width: Number.isFinite(px) && px > 0 ? px : 2,
      highlighter: raster === "maskPen" || transparency > 0,
    })
  }
  const traces: { points: number[]; brush: string }[] = []
  for (const t of all.filter((n) => n.local === "trace")) {
    const points = traceXY(t.text)
    if (points.length >= 2) traces.push({ points, brush: (at(t, "brushRef") ?? "").replace(/^#/, "") })
  }
  if (!traces.length) return null
  const xs = traces.flatMap((t) => t.points.filter((_, i) => i % 2 === 0))
  const ys = traces.flatMap((t) => t.points.filter((_, i) => i % 2 === 1))
  const x0 = Math.min(...xs)
  const y0 = Math.min(...ys)
  const w = Math.max(...xs) - x0 || 1
  const h = Math.max(...ys) - y0 || 1
  return traces.map((t) => {
    const b = brushes.get(t.brush) ?? { color: "#000000", width: 2, highlighter: false }
    return {
      points: t.points.map((v, i) => (i % 2 ? (v - y0) / h : (v - x0) / w)),
      color: b.color.startsWith("#") ? b.color : `#${b.color}`,
      width: b.width,
      ...(b.highlighter ? { highlighter: true as const } : {}),
    }
  })
}

/**
 * A trace's X and Y, decoded: values are explicit (!), first differences (') or second differences ("), the
 * qualifier lasting until the next one; numbers may run together ("0-1587"). Other channels are skipped.
 */
export function traceXY(text: string): number[] {
  const out: number[] = []
  const last = [0, 0]
  const step = [0, 0]
  let mode: "!" | "'" | '"' = "!"
  for (const point of text.split(",")) {
    const values = [...point.matchAll(/([!'"])?\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)/g)]
    if (values.length < 2) continue
    values.forEach((m, c) => {
      if (m[1]) mode = m[1] as typeof mode
      const v = Number(m[2])
      if (c > 1) return
      if (mode === "!") {
        step[c] = v - last[c]
        last[c] = v
      } else if (mode === "'") {
        step[c] = v
        last[c] += v
      } else {
        step[c] += v
        last[c] += step[c]
      }
    })
    out.push(last[0], last[1])
  }
  return out
}
