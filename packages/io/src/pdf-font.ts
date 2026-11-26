/**
 * A TrueType font with no shapes (as Tesseract's "GlyphLessFont"): two empty glyphs, each 500 units wide on a 1000-unit
 * em, ascent 800, descent 200. A PDF's invisible text layer is set in it: nothing shows, but readers can select, search
 * and copy the text (its ToUnicode map says which characters the codes are).
 */

const TAGS = ["cmap", "glyf", "head", "hhea", "hmtx", "loca", "maxp", "name", "post"] as const

class Bytes {
  private parts: number[] = []
  u8 = (n: number) => (this.parts.push(n & 0xff), this)
  u16 = (n: number) => this.u8(n >> 8).u8(n)
  i16 = (n: number) => this.u16(n < 0 ? n + 0x10000 : n)
  u32 = (n: number) => this.u16(Math.floor(n / 0x10000) & 0xffff).u16(n & 0xffff)
  bytes = () => Uint8Array.from(this.parts)
}

/** A table's checksum: its bytes as big-endian 32-bit words (zero-padded), added up modulo 2³². */
function checksum(data: Uint8Array): number {
  let sum = 0
  for (let i = 0; i < data.length; i += 4)
    sum = (sum + ((data[i] << 24) | (data[i + 1] << 16) | (data[i + 2] << 8) | data[i + 3])) >>> 0
  return sum
}

function nameTable(): Uint8Array {
  const names: [number, string][] = [
    [1, "GlyphLessFont"],
    [2, "Regular"],
    [4, "GlyphLessFont"],
    [6, "GlyphLessFont"],
  ]
  const t = new Bytes()
    .u16(0)
    .u16(names.length)
    .u16(6 + names.length * 12)
  let offset = 0
  for (const [id, text] of names) {
    t.u16(3)
      .u16(1)
      .u16(0x409)
      .u16(id)
      .u16(text.length * 2)
      .u16(offset)
    offset += text.length * 2
  }
  for (const [, text] of names) for (const ch of text) t.u16(ch.charCodeAt(0))
  return t.bytes()
}

let cached: Uint8Array | null = null

/** The font's bytes (the same every time). */
export function glyphlessFont(): Uint8Array {
  if (cached) return cached
  const tables: Record<(typeof TAGS)[number], Uint8Array> = {
    // A Windows Unicode map of nothing (only format 4's closing segment): PDF readers map codes by CIDToGIDMap.
    cmap: new Bytes()
      .u16(0)
      .u16(1)
      .u16(3)
      .u16(1)
      .u32(12)
      .u16(4)
      .u16(24)
      .u16(0)
      .u16(2)
      .u16(2)
      .u16(0)
      .u16(0)
      .u16(0xffff)
      .u16(0)
      .u16(0xffff)
      .u16(1)
      .u16(0)
      .bytes(),
    glyf: new Uint8Array(0),
    head: new Bytes()
      .u32(0x00010000)
      .u32(0x00010000)
      .u32(0) // checkSumAdjustment, set below
      .u32(0x5f0f3cf5)
      .u16(0x000b)
      .u16(1000)
      .u32(0)
      .u32(3850070400) // created: 2026-01-01, in seconds since 1904
      .u32(0)
      .u32(3850070400) // modified
      .i16(0)
      .i16(-200)
      .i16(500)
      .i16(800)
      .u16(0)
      .u16(8)
      .i16(2)
      .i16(0)
      .i16(0)
      .bytes(),
    hhea: new Bytes()
      .u32(0x00010000)
      .i16(800)
      .i16(-200)
      .i16(0)
      .u16(500)
      .i16(0)
      .i16(0)
      .i16(500)
      .i16(1)
      .i16(0)
      .i16(0)
      .i16(0)
      .i16(0)
      .i16(0)
      .i16(0)
      .i16(0)
      .u16(2)
      .bytes(),
    hmtx: new Bytes().u16(500).i16(0).u16(500).i16(0).bytes(),
    loca: new Bytes().u16(0).u16(0).u16(0).bytes(),
    maxp: new Bytes()
      .u32(0x00010000)
      .u16(2)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(1)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(0)
      .u16(0)
      .bytes(),
    name: nameTable(),
    post: new Bytes().u32(0x00030000).u32(0).i16(-100).i16(50).u32(1).u32(0).u32(0).u32(0).u32(0).bytes(),
  }
  const padded = (n: number) => (n + 3) & ~3
  const headerSize = 12 + TAGS.length * 16
  const size = headerSize + TAGS.reduce((s, tag) => s + padded(tables[tag].length), 0)
  const out = new Uint8Array(size)
  const view = new DataView(out.buffer)
  // Offset table: TrueType, the table count and its binary-search fields.
  const n = TAGS.length
  const pow = 2 ** Math.floor(Math.log2(n))
  view.setUint32(0, 0x00010000)
  view.setUint16(4, n)
  view.setUint16(6, pow * 16)
  view.setUint16(8, Math.log2(pow))
  view.setUint16(10, n * 16 - pow * 16)
  let offset = headerSize
  let headAt = 0
  TAGS.forEach((tag, i) => {
    const data = tables[tag]
    const at = 12 + i * 16
    for (let k = 0; k < 4; k++) out[at + k] = tag.charCodeAt(k)
    view.setUint32(at + 4, checksum(data))
    view.setUint32(at + 8, offset)
    view.setUint32(at + 12, data.length)
    out.set(data, offset)
    if (tag === "head") headAt = offset
    offset += padded(data.length)
  })
  view.setUint32(headAt + 8, (0xb1b0afba - checksum(out)) >>> 0)
  cached = out
  return out
}
