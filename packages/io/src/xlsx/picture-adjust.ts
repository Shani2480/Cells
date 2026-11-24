/**
 * Picture Format › Corrections, Color and Transparency in DrawingML. Office 2007's adjustments sit in the blip and
 * every reader applies them when drawing (a:grayscl, a:biLevel, a:duotone, a:lum, a:alphaModFix, a:clrChange). Office
 * 2010's corrections (brightness and contrast, sharpen, saturation, colour tone, artistic effects) are baked into the
 * blip's picture; the original sits in an a14:imgLayer (an HD Photo in Excel's files; Excel takes a PNG there too, and
 * re-renders from it when the corrections change), with the effects listed in it.
 */
import type { DuotoneColor, PictureAdjust, PictureRecolor, ThemeSlot } from "@cells/core"
import { NS_R } from "./constants"
import { at, kid, kids, raw, type XDoc, type XNode } from "./dom"
import { NS_A14 } from "./omml"

export const IMG_PROPS_EXT = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"
export const REL_HDPHOTO = "http://schemas.microsoft.com/office/2007/relationships/hdphoto"

const SLOT_ALIASES: Record<string, ThemeSlot> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" }

const num = (v: string | undefined, d = 0) => (v === undefined || v === "" ? d : Number(v))
const pct = (v: string | undefined, d = 0) => num(v, d * 100000) / 100000
/** Fractions written as Office's thousandths of a percent. */
const k = (v: number) => Math.round(v * 100000)

function duotoneColor(node: XNode): DuotoneColor | null {
  const val = at(node, "val") ?? ""
  let c: DuotoneColor
  if (node.local === "srgbClr") c = { rgb: val.toUpperCase() }
  else if (node.local === "schemeClr") c = { theme: (SLOT_ALIASES[val] ?? val) as ThemeSlot }
  else if (node.local === "prstClr" && (val === "black" || val === "white")) c = { preset: val }
  else if (node.local === "sysClr") c = { rgb: (at(node, "lastClr") ?? "000000").toUpperCase() }
  else return null
  for (const m of node.children) {
    if (m.local === "tint" || m.local === "shade" || m.local === "satMod") c[m.local] = pct(at(m, "val"))
  }
  return c
}

function duotoneColorXml(c: DuotoneColor): string {
  const mods =
    (c.tint !== undefined ? `<a:tint val="${k(c.tint)}"/>` : "") +
    (c.shade !== undefined ? `<a:shade val="${k(c.shade)}"/>` : "") +
    (c.satMod !== undefined ? `<a:satMod val="${k(c.satMod)}"/>` : "")
  const [tag, val] = c.preset
    ? ["prstClr", c.preset]
    : c.theme
      ? ["schemeClr", c.theme]
      : ["srgbClr", c.rgb ?? "000000"]
  return mods ? `<a:${tag} val="${val}">${mods}</a:${tag}>` : `<a:${tag} val="${val}"/>`
}

/** A picture's adjustments from its blip, and its original (the a14:imgLayer's relationship id), if any. */
export function readPictureAdjust(
  blip: XNode | undefined,
  doc: XDoc | undefined,
): { adjust?: PictureAdjust; layerRid?: string } {
  if (!blip) return {}
  const a: PictureAdjust = {}
  for (const m of blip.children) {
    if (m.local === "grayscl") a.recolor = { kind: "grayscale" }
    else if (m.local === "biLevel") a.recolor = { kind: "blackWhite", threshold: pct(at(m, "thresh"), 0.5) }
    else if (m.local === "duotone") {
      const colors = m.children.map(duotoneColor).filter((c): c is DuotoneColor => !!c)
      if (colors.length === 2) a.recolor = { kind: "duotone", colors: [colors[0], colors[1]] }
    } else if (m.local === "lum") {
      const bright = pct(at(m, "bright"))
      const contrast = pct(at(m, "contrast"))
      if (bright || contrast) a.lum = { bright, contrast }
    } else if (m.local === "alphaModFix") {
      const t = 1 - pct(at(m, "amt"), 1)
      if (t > 0) a.transparency = Math.round(t * 1000) / 1000
    } else if (m.local === "clrChange") {
      const from = kid(m, "clrFrom")?.children[0]
      const to = kid(m, "clrTo")?.children[0]
      const clear = to?.children.some((c) => c.local === "alpha" && pct(at(c, "val")) === 0)
      if (from?.local === "srgbClr" && clear) a.transparentColor = (at(from, "val") ?? "000000").toUpperCase()
    }
  }
  let layerRid: string | undefined
  for (const ext of kids(kid(blip, "extLst"), "ext")) {
    if (at(ext, "uri")?.toUpperCase() !== IMG_PROPS_EXT) continue
    const layer = kid(kid(ext, "imgProps"), "imgLayer")
    layerRid = at(layer, "embed", NS_R)
    for (const e of kids(layer, "imgEffect")) {
      const fx = e.children[0]
      if (!fx) continue
      if (fx.local === "brightnessContrast") {
        a.bright = pct(at(fx, "bright"))
        a.contrast = pct(at(fx, "contrast"))
        if (!a.bright) delete a.bright
        if (!a.contrast) delete a.contrast
      } else if (fx.local === "sharpenSoften") a.sharpen = pct(at(fx, "amount"))
      else if (fx.local === "saturation") a.saturation = pct(at(fx, "sat"), 1)
      else if (fx.local === "colorTemperature") a.temperature = num(at(fx, "colorTemp"), 6500)
      else if (doc) (a.effects ??= []).push(raw(doc, fx))
    }
  }
  return { ...(Object.keys(a).length ? { adjust: a } : {}), ...(layerRid ? { layerRid } : {}) }
}

/** The blip's adjustment elements (before its extLst). */
export function adjustBlipXml(a: PictureAdjust | undefined): string {
  if (!a) return ""
  let xml = ""
  if (a.transparency) xml += `<a:alphaModFix amt="${k(1 - a.transparency)}"/>`
  if (a.transparentColor) {
    const c = a.transparentColor
    xml += `<a:clrChange><a:clrFrom><a:srgbClr val="${c}"/></a:clrFrom><a:clrTo><a:srgbClr val="${c}"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>`
  }
  xml += recolorXml(a.recolor)
  if (a.lum) {
    const attrs =
      (a.lum.bright ? ` bright="${k(a.lum.bright)}"` : "") + (a.lum.contrast ? ` contrast="${k(a.lum.contrast)}"` : "")
    if (attrs) xml += `<a:lum${attrs}/>`
  }
  return xml
}

function recolorXml(r: PictureRecolor | undefined): string {
  if (!r) return ""
  if (r.kind === "grayscale") return `<a:grayscl/>`
  if (r.kind === "blackWhite") return `<a:biLevel thresh="${k(r.threshold)}"/>`
  return `<a:duotone>${duotoneColorXml(r.colors[0])}${duotoneColorXml(r.colors[1])}</a:duotone>`
}

/** The a14:imgProps extension: the original's relationship and the corrections, in the order Office applies them. */
export function imgPropsExtXml(a: PictureAdjust, layerRid: string): string {
  const effects: string[] = [...(a.effects ?? [])]
  if (a.sharpen) effects.push(`<a14:sharpenSoften amount="${k(a.sharpen)}"/>`)
  if (a.temperature !== undefined && a.temperature !== 6500)
    effects.push(`<a14:colorTemperature colorTemp="${Math.round(a.temperature)}"/>`)
  if (a.saturation !== undefined && a.saturation !== 1) effects.push(`<a14:saturation sat="${k(a.saturation)}"/>`)
  if (a.bright || a.contrast) {
    const attrs = (a.bright ? ` bright="${k(a.bright)}"` : "") + (a.contrast ? ` contrast="${k(a.contrast)}"` : "")
    effects.push(`<a14:brightnessContrast${attrs}/>`)
  }
  const list = effects.map((e) => `<a14:imgEffect>${e}</a14:imgEffect>`).join("")
  return (
    `<a:ext uri="${IMG_PROPS_EXT}"><a14:imgProps xmlns:a14="${NS_A14}">` +
    `<a14:imgLayer r:embed="${layerRid}">${list}</a14:imgLayer></a14:imgProps></a:ext>`
  )
}
