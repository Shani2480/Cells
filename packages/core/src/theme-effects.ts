/**
 * Page Layout › Themes › Effects: a theme's format scheme (DrawingML's a:fmtScheme), kept as its XML so Excel gets it
 * exactly, and the parts of its three effect styles Cells draws (shadows and glows; bevels, reflections and 3-D show
 * in Excel). A shape's theme effect (the effect rows of Shape Styles, Excel's effectRef) picks one of the three.
 */

export type ThemeEffects = { name: string; scheme: string }

/** Office's format scheme, as Excel 365 writes it in a new workbook. */
export const OFFICE_EFFECTS: ThemeEffects = {
  name: "Office",
  scheme:
    '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme>',
}

export type EffectLook = {
  /** A shadow: offset and blur in pixels, opacity 0–1; inner: inside the shape. */
  shadow?: { dx: number; dy: number; blur: number; alpha: number; color: string; inner?: boolean }
  /** A glow around the shape, radius in pixels; colour null: the shape's own. */
  glow?: { radius: number; alpha: number; color: string | null }
  /** Effects Cells doesn't draw (they show in Excel). */
  other: ("reflection" | "bevel" | "soft edges")[]
}

const EMU_PX = 1 / 9525
const round = (v: number) => Math.round(v * 100) / 100

const attr = (xml: string, name: string) => {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(xml)
  return m ? Number(m[1]) : undefined
}

function colorOf(xml: string): { color: string | null; alpha: number } {
  const alpha = (attr(/<a:alpha\b[^>]*>/.exec(xml)?.[0] ?? "", "val") ?? 100000) / 100000
  const rgb = /<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(xml)?.[1]
  if (rgb) return { color: `#${rgb.toUpperCase()}`, alpha }
  if (/<a:prstClr\s+val="black"/.test(xml)) return { color: "#000000", alpha }
  if (/<a:prstClr\s+val="white"/.test(xml)) return { color: "#FFFFFF", alpha }
  return { color: null, alpha }
}

/** The three effect styles of a format scheme (Office's when none is given). */
export function effectLooks(scheme: string = OFFICE_EFFECTS.scheme): EffectLook[] {
  const styles = [...scheme.matchAll(/<a:effectStyle>([\s\S]*?)<\/a:effectStyle>/g)].map((m) => m[1])
  return [0, 1, 2].map((i) => {
    const xml = styles[i] ?? ""
    const look: EffectLook = { other: [] }
    const shadow = /<a:(outerShdw|innerShdw)\b([^>]*)>([\s\S]*?)<\/a:\1>/.exec(xml)
    if (shadow) {
      const dist = (attr(shadow[2], "dist") ?? 0) * EMU_PX
      const dir = ((attr(shadow[2], "dir") ?? 0) / 60000) * (Math.PI / 180)
      const c = colorOf(shadow[3])
      look.shadow = {
        dx: round(Math.cos(dir) * dist),
        dy: round(Math.sin(dir) * dist),
        blur: round((attr(shadow[2], "blurRad") ?? 0) * EMU_PX),
        alpha: c.alpha,
        color: c.color ?? "#000000",
        ...(shadow[1] === "innerShdw" ? { inner: true } : {}),
      }
    }
    const glow = /<a:glow\b([^>]*)>([\s\S]*?)<\/a:glow>/.exec(xml)
    if (glow) {
      const c = colorOf(glow[2])
      look.glow = { radius: round((attr(glow[1], "rad") ?? 0) * EMU_PX), alpha: c.alpha, color: c.color }
    }
    if (/<a:reflection\b/.test(xml)) look.other.push("reflection")
    if (/<a:bevelT\b/.test(xml)) look.other.push("bevel")
    if (/<a:softEdge\b/.test(xml)) look.other.push("soft edges")
    return look
  })
}
