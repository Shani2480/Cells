import { glyphlessFont } from "./pdf-font"

/**
 * A PDF of pictures: one page per image, each JPEG drawn over the whole page (File › Export › Create PDF, where each
 * printed page is drawn at print resolution first). Sizes are in points. A page's text can lie over its picture,
 * invisible, where it's drawn, so the PDF can be searched and its text selected and copied (as a scanner's
 * searchable PDF).
 */

/**
 * A piece of text on a page: where its baseline starts, in points from the page's top left; its font size and width
 * in points; and its direction (clockwise from the right, in degrees: 90 runs down the page).
 */
export type PdfText = { text: string; x: number; y: number; size: number; width: number; angle?: number }

export type PdfPage = {
  /** The page in points (1/72 inch). */
  width: number
  height: number
  /** The page's picture: JPEG bytes and its size in pixels. */
  jpeg: Uint8Array
  pixelWidth: number
  pixelHeight: number
  /** The text drawn in the picture, laid invisibly over it. */
  text?: readonly PdfText[]
}

export type PdfInfo = { title?: string; creator?: string; date?: Date }

/** PDF text strings: ASCII as a literal, anything else as UTF-16 in hex. */
function pdfString(text: string): string {
  if (/^[\x20-\x7e]*$/.test(text)) return "(" + text.replace(/[\\()]/g, (c) => "\\" + c) + ")"
  let hex = "FEFF"
  for (let i = 0; i < text.length; i++) hex += text.charCodeAt(i).toString(16).padStart(4, "0").toUpperCase()
  return "<" + hex + ">"
}

function pdfDate(d: Date): string {
  const p = (n: number) => String(n).padStart(2, "0")
  return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`
}

/**
 * The colour space of a JPEG's pixels, from its frame header's channel count: grey (1), RGB (3) or CMYK (4). A reader
 * that trusts the PDF over the picture (Android's) draws nothing when they disagree.
 */
export function jpegColorSpace(jpeg: Uint8Array): "/DeviceGray" | "/DeviceRGB" | "/DeviceCMYK" {
  let i = 2
  while (i + 9 < jpeg.length && jpeg[i] === 0xff) {
    const marker = jpeg[i + 1]
    const length = (jpeg[i + 2] << 8) | jpeg[i + 3]
    // Start of frame (baseline, extended, progressive, lossless…), not DHT (C4), JPG (C8) or DAC (CC).
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      const channels = jpeg[i + 9]
      return channels === 1 ? "/DeviceGray" : channels === 4 ? "/DeviceCMYK" : "/DeviceRGB"
    }
    i += 2 + length
  }
  return "/DeviceRGB"
}

const num = (n: number) => (Math.round(n * 1000) / 1000).toString()
const hex4 = (n: number) => n.toString(16).padStart(4, "0").toUpperCase()

/**
 * The text layer's codes: each character used anywhere in the document gets a code of its own (from 1), shown by the
 * shapeless font's one glyph, and the ToUnicode map gives the character back for search and copy.
 */
class TextCodes {
  readonly codes = new Map<string, number>()
  code(ch: string): number {
    let c = this.codes.get(ch)
    if (c === undefined) {
      c = this.codes.size + 1
      this.codes.set(ch, c)
    }
    return c
  }
  get size() {
    return this.codes.size
  }
  /** The ToUnicode CMap: each code to its character in UTF-16 (bfchar, at most 100 to a block). */
  toUnicode(): string {
    const entries = [...this.codes].map(([ch, c]) => {
      let utf16 = ""
      for (let i = 0; i < ch.length; i++) utf16 += hex4(ch.charCodeAt(i))
      return `<${hex4(c)}> <${utf16}>`
    })
    let blocks = ""
    for (let i = 0; i < entries.length; i += 100) {
      const part = entries.slice(i, i + 100)
      blocks += `${part.length} beginbfchar\n${part.join("\n")}\nendbfchar\n`
    }
    return (
      "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n" +
      "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n" +
      "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n" +
      "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n" +
      blocks +
      "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend"
    )
  }
}

/**
 * A page's text drawn invisibly (render mode 3) in the shapeless font: every glyph is half an em wide, so each piece
 * is stretched to its width, and turned to its direction.
 */
function textLayer(text: readonly PdfText[], pageHeight: number, codes: TextCodes): string {
  let out = ""
  for (const t of text) {
    const chars = [...t.text]
    if (!chars.length || !t.text.trim() || t.size <= 0) continue
    const a = ((t.angle ?? 0) * Math.PI) / 180
    const dx = Math.cos(a)
    const dy = Math.sin(a)
    const across = Math.max(0.01, t.width) / (chars.length * 0.5)
    const matrix = [dx * across, -dy * across, dy * t.size, dx * t.size, t.x, pageHeight - t.y].map(num).join(" ")
    out += `${matrix} Tm <${chars.map((ch) => hex4(codes.code(ch))).join("")}> Tj\n`
  }
  return out ? `BT 3 Tr /F1 1 Tf\n${out}ET` : ""
}

export function pdfFromImages(pages: readonly PdfPage[], info: PdfInfo = {}): Uint8Array {
  const encoder = new TextEncoder()
  const chunks: Uint8Array[] = []
  let length = 0
  const offsets: number[] = []
  const write = (data: string | Uint8Array) => {
    const bytes = typeof data === "string" ? encoder.encode(data) : data
    chunks.push(bytes)
    length += bytes.length
  }
  const object = (id: number, body: string | (() => void)) => {
    offsets[id] = length
    write(`${id} 0 obj\n`)
    if (typeof body === "string") write(body)
    else body()
    write("\nendobj\n")
  }

  // Objects: 1 catalog, 2 page tree, 3 info, then per page: page, contents, image; then the text layer's font.
  const pageId = (i: number) => 4 + i * 3
  const fontId = 4 + pages.length * 3
  const hasText = pages.some((p) => p.text?.some((t) => t.text.trim()))
  const codes = new TextCodes()
  // The second line's high bytes mark the file as binary for tools that guess.
  write("%PDF-1.4\n")
  write(Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a))
  object(1, "<< /Type /Catalog /Pages 2 0 R >>")
  object(2, `<< /Type /Pages /Kids [${pages.map((_, i) => `${pageId(i)} 0 R`).join(" ")}] /Count ${pages.length} >>`)
  const meta = [
    info.title ? `/Title ${pdfString(info.title)}` : "",
    `/Creator ${pdfString(info.creator ?? "Cells")}`,
    `/Producer ${pdfString(info.creator ?? "Cells")}`,
    info.date ? `/CreationDate (${pdfDate(info.date)})` : "",
  ]
  object(3, `<< ${meta.filter(Boolean).join(" ")} >>`)
  pages.forEach((page, i) => {
    const id = pageId(i)
    const w = num(page.width)
    const h = num(page.height)
    object(
      id,
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /XObject << /Im0 ${id + 2} 0 R >>${hasText ? ` /Font << /F1 ${fontId} 0 R >>` : ""} >> /Contents ${id + 1} 0 R >>`,
    )
    const layer = page.text ? textLayer(page.text, page.height, codes) : ""
    const content = `q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q${layer ? "\n" + layer : ""}`
    object(id + 1, `<< /Length ${content.length} >>\nstream\n${content}\nendstream`)
    object(id + 2, () => {
      write(
        `<< /Type /XObject /Subtype /Image /Width ${page.pixelWidth} /Height ${page.pixelHeight} /ColorSpace ${jpegColorSpace(page.jpeg)} /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>\nstream\n`,
      )
      write(page.jpeg)
      write("\nendstream")
    })
  })
  if (hasText) {
    // A Type 0 font over the shapeless TrueType: codes are CIDs, every CID drawn with glyph 1.
    object(
      fontId,
      `<< /Type /Font /Subtype /Type0 /BaseFont /GlyphLessFont /Encoding /Identity-H /DescendantFonts [${fontId + 1} 0 R] /ToUnicode ${fontId + 4} 0 R >>`,
    )
    object(
      fontId + 1,
      `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /GlyphLessFont /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${fontId + 2} 0 R /DW 500 /CIDToGIDMap ${fontId + 5} 0 R >>`,
    )
    object(
      fontId + 2,
      `<< /Type /FontDescriptor /FontName /GlyphLessFont /Flags 5 /FontBBox [0 -200 500 800] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 800 /StemV 80 /FontFile2 ${fontId + 3} 0 R >>`,
    )
    const font = glyphlessFont()
    object(fontId + 3, () => {
      write(`<< /Length ${font.length} /Length1 ${font.length} >>\nstream\n`)
      write(font)
      write("\nendstream")
    })
    const cmap = codes.toUnicode()
    object(fontId + 4, `<< /Length ${cmap.length} >>\nstream\n${cmap}\nendstream`)
    const map = new Uint8Array((codes.size + 1) * 2)
    for (let c = 1; c <= codes.size; c++) map[c * 2 + 1] = 1
    object(fontId + 5, () => {
      write(`<< /Length ${map.length} >>\nstream\n`)
      write(map)
      write("\nendstream")
    })
  }
  const count = fontId + (hasText ? 6 : 0)
  const xref = length
  let table = `xref\n0 ${count}\n0000000000 65535 f \n`
  for (let id = 1; id < count; id++) table += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`
  write(table)
  write(`trailer\n<< /Size ${count} /Root 1 0 R /Info 3 0 R >>\nstartxref\n${xref}\n%%EOF\n`)

  const out = new Uint8Array(length)
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return out
}
