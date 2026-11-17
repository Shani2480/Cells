import type { EmbeddedFile, Theme } from "@cells/core"

/** Insert › Object's icon size: the page and the file's name under it. */
export const OBJECT_ICON_SIZE = { w: 84, h: 76 }

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
const n = (v: number) => String(Math.round(v * 100) / 100)

type Kind = { color: string; mark: string }

const KINDS: [RegExp, Kind][] = [
  [/^(docx?|docm|dotx?|odt|rtf)$/, { color: "#185ABD", mark: "W" }],
  [/^(xlsx?|xlsm|xlsb|xltx?|ods|csv)$/, { color: "#107C41", mark: "X" }],
  [/^(pptx?|pptm|ppsx?|odp)$/, { color: "#C43E1C", mark: "P" }],
  [/^pdf$/, { color: "#D93025", mark: "PDF" }],
  [/^(png|jpe?g|gif|bmp|svg|webp|tiff?)$/, { color: "#0F7B8A", mark: "IMG" }],
  [/^(zip|7z|rar|gz|tar)$/, { color: "#8A6A1F", mark: "ZIP" }],
]

/** The colour and letters a file's icon shows, from its extension (or the Office class that embedded it). */
export function fileKind(name: string, progId = ""): Kind | null {
  const ext = /\.([^.]+)$/.exec(name)?.[1]?.toLowerCase() ?? ""
  const hit = KINDS.find(([re]) => re.test(ext))?.[1]
  if (hit) return hit
  if (/^Word\./i.test(progId)) return KINDS[0]![1]
  if (/^Excel\./i.test(progId)) return KINDS[1]![1]
  if (/^PowerPoint\./i.test(progId)) return KINDS[2]![1]
  if (/^AcroExch\./i.test(progId)) return KINDS[3]![1]
  return null
}

/** A name cut in the middle with an ellipsis to `chars` (keeping its end, where the extension is). */
function shorten(name: string, chars: number): string {
  if (name.length <= chars) return name
  const keep = chars - 1
  const tail = Math.min(Math.ceil(keep / 2), 10)
  return `${name.slice(0, keep - tail).trimEnd()}…${name.slice(name.length - tail)}`
}

/** The label in at most two lines of `chars`, broken at a space where there is one (as Office wraps it). */
export function labelLines(name: string, chars: number): string[] {
  if (name.length <= chars) return [name]
  const space = name.lastIndexOf(" ", chars)
  const cut = space > chars / 3 ? space : chars
  return [name.slice(0, cut), shorten(name.slice(cut).trimStart(), chars)]
}

/**
 * An embedded file as Office shows one "as an icon": a page with a folded corner in the colour of its kind, and its
 * name under it. Drawn in its box (0,0)–(w,h) in the theme's body font.
 */
export function embedSvg(embed: EmbeddedFile, w: number, h: number, theme: Theme): string {
  const k = Math.min(w / OBJECT_ICON_SIZE.w, h / OBJECT_ICON_SIZE.h)
  const pw = 30 * k
  const ph = 38 * k
  const x = (w - pw) / 2
  const y = 4 * k
  const fold = 9 * k
  const kind = fileKind(embed.name, embed.progId)
  const color = kind?.color ?? "#6B6B6B"
  const page =
    `<path d="M${n(x)} ${n(y)}h${n(pw - fold)}l${n(fold)} ${n(fold)}v${n(ph - fold)}h${n(-pw)}z" fill="#FFFFFF" stroke="${color}" stroke-width="${n(1.25 * k)}"/>` +
    `<path d="M${n(x + pw - fold)} ${n(y)}v${n(fold)}h${n(fold)}" fill="none" stroke="${color}" stroke-width="${n(1.25 * k)}"/>`
  const mark = kind
    ? `<rect x="${n(x - 4 * k)}" y="${n(y + ph * 0.42)}" width="${n(kind.mark.length > 1 ? 24 * k : 16 * k)}" height="${n(14 * k)}" rx="${n(2 * k)}" fill="${color}"/>` +
      `<text x="${n(x - 4 * k + (kind.mark.length > 1 ? 12 : 8) * k)}" y="${n(y + ph * 0.42 + 10.5 * k)}" text-anchor="middle" font-family="Segoe UI, Arial, sans-serif" font-weight="700" font-size="${n((kind.mark.length > 1 ? 8 : 10) * k)}" fill="#FFFFFF">${kind.mark}</text>`
    : [0.4, 0.55, 0.7]
        .map(
          (f) =>
            `<rect x="${n(x + 6 * k)}" y="${n(y + ph * f)}" width="${n(pw - 12 * k)}" height="${n(1.5 * k)}" fill="${color}"/>`,
        )
        .join("")
  const font = `font-family="${esc(theme.fonts.minor)}, Calibri, Carlito, sans-serif"`
  const label = labelLines(embed.name, Math.max(6, Math.floor(w / (5.4 * k))))
    .map(
      (line, i) =>
        `<text x="${n(w / 2)}" y="${n(y + ph + (13 + i * 12) * k)}" text-anchor="middle" ${font} font-size="${n(10 * k)}" fill="#000000">${esc(line)}</text>`,
    )
    .join("")
  return `${page}${mark}${label}`
}

/** A whole SVG document of the icon, for the PNG files carry. */
export function embedDocument(embed: EmbeddedFile, w: number, h: number, theme: Theme): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${n(w)}" height="${n(h)}" viewBox="0 0 ${n(w)} ${n(h)}">${embedSvg(embed, w, h, theme)}</svg>`
}
