/**
 * BIFF8 records ([MS-XLS] 2.1.4): a 2-byte type, a 2-byte length and up to 8224 bytes of data, with CONTINUE
 * records carrying the rest of a longer one. A record here is its data with the CONTINUE pieces joined, and where
 * each piece starts, since a string broken across pieces restarts with a new flags byte.
 */

export const R = {
  FORMULA: 0x0006,
  EOF: 0x000a,
  CALCMODE: 0x000d,
  PROTECT: 0x0012,
  EXTERNSHEET: 0x0017,
  NAME: 0x0018,
  WINDOW1: 0x003d,
  DATEMODE: 0x0022,
  EXTERNNAME: 0x0023,
  FILEPASS: 0x002f,
  FONT: 0x0031,
  CONTINUE: 0x003c,
  PANE: 0x0041,
  CODEPAGE: 0x0042,
  DEFCOLWIDTH: 0x0055,
  COLINFO: 0x007d,
  BOUNDSHEET: 0x0085,
  PALETTE: 0x0092,
  STANDARDWIDTH: 0x0099,
  MULRK: 0x00bd,
  MULBLANK: 0x00be,
  XF: 0x00e0,
  MERGECELLS: 0x00e5,
  SST: 0x00fc,
  LABELSST: 0x00fd,
  SUPBOOK: 0x01ae,
  DIMENSIONS: 0x0200,
  BLANK: 0x0201,
  NUMBER: 0x0203,
  LABEL: 0x0204,
  BOOLERR: 0x0205,
  STRING: 0x0207,
  ROW: 0x0208,
  ARRAY: 0x0221,
  DEFAULTROWHEIGHT: 0x0225,
  WSBOOL: 0x0081,
  WINDOW2: 0x023e,
  RK: 0x027e,
  SCL: 0x00a0,
  SHRFMLA: 0x04bc,
  FORMAT: 0x041e,
  BOF: 0x0809,
} as const

export type BiffRecord = {
  type: number
  /** Where the record starts in the stream (BOUNDSHEET points at a sheet's BOF this way). */
  offset: number
  data: Uint8Array
  /** Where each CONTINUE piece starts in `data` (after the first piece). */
  breaks: number[]
}

/** The stream's records, CONTINUE pieces joined to the record before them. */
export function readRecords(stream: Uint8Array): BiffRecord[] {
  const out: BiffRecord[] = []
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength)
  let at = 0
  let pieces: Uint8Array[] = []
  let current: { type: number; offset: number } | null = null
  const flush = () => {
    if (!current) return
    const size = pieces.reduce((n, p) => n + p.length, 0)
    const data = new Uint8Array(size)
    const breaks: number[] = []
    let pos = 0
    pieces.forEach((p, i) => {
      if (i) breaks.push(pos)
      data.set(p, pos)
      pos += p.length
    })
    out.push({ ...current, data, breaks })
    current = null
    pieces = []
  }
  while (at + 4 <= stream.length) {
    const type = view.getUint16(at, true)
    const length = view.getUint16(at + 2, true)
    const data = stream.subarray(at + 4, Math.min(at + 4 + length, stream.length))
    // A CONTINUE after a record that takes them. (After a drawing record it belongs to the drawing, which Cells
    // skips anyway, so joining it there does no harm.)
    if (type === R.CONTINUE && current) pieces.push(data)
    else {
      flush()
      if (type === 0 && length === 0) {
        // Padding at the end of the stream.
        at += 4
        continue
      }
      current = { type, offset: at }
      pieces = [data]
    }
    at += 4 + length
  }
  flush()
  return out
}

/** Reads a record's data from left to right. */
export class Cursor {
  private readonly view: DataView
  pos: number

  constructor(
    readonly data: Uint8Array,
    pos = 0,
    private readonly breaks: number[] = [],
  ) {
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength)
    this.pos = pos
  }

  get left() {
    return this.data.length - this.pos
  }

  u8() {
    return this.data[this.pos++] ?? 0
  }

  u16() {
    if (this.pos + 2 > this.data.length) return ((this.pos = this.data.length), 0)
    const v = this.view.getUint16(this.pos, true)
    this.pos += 2
    return v
  }

  i16() {
    const v = this.u16()
    return v >= 0x8000 ? v - 0x10000 : v
  }

  u32() {
    if (this.pos + 4 > this.data.length) return ((this.pos = this.data.length), 0)
    const v = this.view.getUint32(this.pos, true)
    this.pos += 4
    return v
  }

  f64() {
    if (this.pos + 8 > this.data.length) return ((this.pos = this.data.length), 0)
    const v = this.view.getFloat64(this.pos, true)
    this.pos += 8
    return v
  }

  skip(n: number) {
    this.pos = Math.min(this.data.length, this.pos + n)
  }

  bytes(n: number) {
    const out = this.data.subarray(this.pos, Math.min(this.data.length, this.pos + n))
    this.pos += out.length
    return out
  }

  /** The next CONTINUE break after the cursor, or the end. */
  private nextBreak() {
    for (const b of this.breaks) if (b > this.pos) return b
    return this.data.length
  }

  /**
   * `count` characters, one or two bytes each by `high`; where the record breaks into a CONTINUE the characters go on
   * after a new flags byte, which may switch between one and two bytes ([MS-XLS] 2.5.293).
   */
  chars(count: number, high: boolean): string {
    let out = ""
    while (count > 0 && this.pos < this.data.length) {
      const end = this.nextBreak()
      const width = high ? 2 : 1
      const fit = Math.min(count, Math.floor((end - this.pos) / width))
      if (high) for (let i = 0; i < fit; i++) out += String.fromCharCode(this.view.getUint16(this.pos + i * 2, true))
      else for (let i = 0; i < fit; i++) out += String.fromCharCode(this.data[this.pos + i])
      this.pos += fit * width
      count -= fit
      if (count > 0) {
        if (this.pos !== end) break
        high = (this.u8() & 1) === 1
      }
    }
    return out
  }

  /** XLUnicodeRichExtendedString: a 2-byte (or 1-byte) count, flags, runs and phonetic data skipped. */
  string(countBytes: 1 | 2 = 2): string {
    const count = countBytes === 2 ? this.u16() : this.u8()
    const flags = this.u8()
    const runs = flags & 0x08 ? this.u16() : 0
    const ext = flags & 0x04 ? this.u32() : 0
    const text = this.chars(count, (flags & 1) === 1)
    this.skip(runs * 4 + ext)
    return text
  }
}
