import type { HeaderPicture, HeaderPictureId } from "@cells/core"
import { escapeXml } from "@cells/ooxml"

/**
 * Header and footer pictures (Header & Footer › Picture, Format Picture) in SpreadsheetML: a VML drawing the sheet
 * names in legacyDrawingHF, one v:shape per picture whose id says where it goes ("CH" centre header, "LFFIRST" the
 * first page's left footer…), its size in the style (points) and Format Picture's settings on v:imagedata, the
 * picture itself an image relationship of the drawing. As Excel 365 writes it.
 */

/** A VML fraction: "8738f" is 8738/65536; a plain number is the fraction itself. */
function fraction(v: string | undefined, fallback = 0): number {
  if (v === undefined || v === "") return fallback
  const n = v.endsWith("f") ? Number(v.slice(0, -1)) / 65536 : Number(v)
  return Number.isFinite(n) ? n : fallback
}

/** Points from a VML length ("150pt", "200px", "2in", "1.5cm"; a bare number is pixels). */
function points(v: string | undefined): number | null {
  const m = /^(-?[\d.]+)\s*(pt|px|in|cm|mm)?$/.exec(v?.trim() ?? "")
  if (!m) return null
  const n = Number(m[1])
  const unit = m[2] ?? "px"
  return unit === "pt"
    ? n
    : unit === "px"
      ? n * 0.75
      : unit === "in"
        ? n * 72
        : unit === "cm"
          ? (n * 72) / 2.54
          : (n * 72) / 25.4
}

const attrsOf = (tag: string): Record<string, string> => {
  const out: Record<string, string> = {}
  for (const m of tag.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) out[m[1]] = m[2] ?? m[3] ?? ""
  return out
}

const PLACE = /^[LCR][HF](EVEN|FIRST)?$/

/**
 * The pictures of a header/footer VML drawing, each with the relationship id of its image; null when the drawing
 * holds something else Cells can't keep (it then stays in the file as read).
 */
export function readHeaderPictureVml(
  vml: string,
): { id: HeaderPictureId; relId: string; picture: Omit<HeaderPicture, "src"> }[] | null {
  const out: { id: HeaderPictureId; relId: string; picture: Omit<HeaderPicture, "src"> }[] = []
  for (const m of vml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)) {
    const shape = attrsOf(m[1])
    const id = shape.id ?? ""
    if (!PLACE.test(id)) return null
    const image = /<v:imagedata\b([^>]*?)\/?>/.exec(m[2])
    if (!image) return null
    const a = attrsOf(image[1])
    const relId = a["o:relid"] ?? a["r:id"]
    if (!relId) return null
    const style = Object.fromEntries(
      (shape.style ?? "")
        .split(";")
        .map((d) => d.split(":").map((x) => x.trim()))
        .filter((d) => d.length === 2),
    )
    const lock = /<o:lock\b([^>]*)\/?>/.exec(m[2])
    const gain = fraction(a.gain, 1)
    const black = fraction(a.blacklevel, 0)
    out.push({
      id,
      relId,
      picture: {
        title: a["o:title"] ?? "",
        width: points(style.width) ?? 72,
        height: points(style.height) ?? 72,
        crop: {
          top: fraction(a.croptop),
          bottom: fraction(a.cropbottom),
          left: fraction(a.cropleft),
          right: fraction(a.cropright),
        },
        lockAspect: !(lock && /aspectratio\s*=\s*["']f/.test(lock[1])),
        color: a.bilevel === "t" ? "blackAndWhite" : a.grayscale === "t" ? "grayscale" : "automatic",
        brightness: Math.min(1, Math.max(0, 0.5 + black)),
        // Excel's gain: twice the contrast up to 0.5, then 0.5 / (1 − contrast).
        contrast: Math.min(1, Math.max(0, gain <= 1 ? gain / 2 : 1 - 0.5 / gain)),
      },
    })
  }
  return out
}

const f = (n: number) => `${Math.round(n * 65536)}f`
const pt = (n: number) => `${Number(n.toFixed(2))}pt`

/** The VML drawing of a sheet's header and footer pictures; `relOf` gives each picture's image relationship. */
export function headerPictureVml(
  pictures: Record<HeaderPictureId, HeaderPicture>,
  relOf: (id: HeaderPictureId) => string,
  sheetIndex: number,
): string {
  let shapes = ""
  let n = 0
  for (const [id, p] of Object.entries(pictures)) {
    const image: Record<string, string | undefined> = {
      "o:relid": relOf(id),
      "o:title": p.title,
      croptop: p.crop.top ? f(p.crop.top) : undefined,
      cropbottom: p.crop.bottom ? f(p.crop.bottom) : undefined,
      cropleft: p.crop.left ? f(p.crop.left) : undefined,
      cropright: p.crop.right ? f(p.crop.right) : undefined,
      gain:
        p.contrast !== 0.5 ? f(p.contrast <= 0.5 ? p.contrast * 2 : 0.5 / Math.max(0.01, 1 - p.contrast)) : undefined,
      blacklevel: p.brightness !== 0.5 ? f(p.brightness - 0.5) : undefined,
      grayscale: p.color !== "automatic" ? "t" : undefined,
      bilevel: p.color === "blackAndWhite" ? "t" : undefined,
    }
    const attrs = Object.entries(image)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => ` ${k}="${escapeXml(v!)}"`)
      .join("")
    n++
    shapes +=
      `<v:shape id="${id}" o:spid="_x0000_s${1024 * (sheetIndex + 1) + n}" type="#_x0000_t75" ` +
      `style='position:absolute;margin-left:0;margin-top:0;width:${pt(p.width)};height:${pt(p.height)};z-index:${n}'>` +
      `<v:imagedata${attrs}/><o:lock v:ext="edit" rotation="t"${p.lockAspect ? "" : ' aspectratio="f"'}/></v:shape>`
  }
  return (
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheetIndex + 1}"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" ` +
    `path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas>` +
    `<v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/>` +
    `<v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/>` +
    `<v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>` +
    `<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/>` +
    `</v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>` +
    `<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>${shapes}</xml>`
  )
}
