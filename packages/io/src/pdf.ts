/**
 * A PDF of pictures: one page per image, each JPEG drawn over the whole page (File › Export › Create PDF, where each
 * printed page is drawn at print resolution first). Sizes are in points.
 */

export type PdfPage = {
  /** The page in points (1/72 inch). */
  width: number
  height: number
  /** The page's picture: JPEG bytes and its size in pixels. */
  jpeg: Uint8Array
  pixelWidth: number
  pixelHeight: number
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

  // Objects: 1 catalog, 2 page tree, 3 info, then per page: page, contents, image.
  const pageId = (i: number) => 4 + i * 3
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
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /XObject << /Im0 ${id + 2} 0 R >> >> /Contents ${id + 1} 0 R >>`,
    )
    const content = `q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q`
    object(id + 1, `<< /Length ${content.length} >>\nstream\n${content}\nendstream`)
    object(id + 2, () => {
      write(
        `<< /Type /XObject /Subtype /Image /Width ${page.pixelWidth} /Height ${page.pixelHeight} /ColorSpace ${jpegColorSpace(page.jpeg)} /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>\nstream\n`,
      )
      write(page.jpeg)
      write("\nendstream")
    })
  })
  const count = 4 + pages.length * 3
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
