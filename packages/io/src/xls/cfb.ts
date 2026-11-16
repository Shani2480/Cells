/**
 * The Compound File Binary format ([MS-CFB]) that holds an Excel 97-2003 workbook: a small file system of sectors,
 * a FAT chaining them into streams, and a directory naming the streams. Only reading, and only what an .xls needs:
 * the streams by name.
 */

export const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

const END_OF_CHAIN = 0xfffffffe
const FREE = 0xffffffff

export class CfbError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "CfbError"
  }
}

export function isCompoundFile(bytes: Uint8Array): boolean {
  return bytes.length >= 512 && CFB_SIGNATURE.every((b, i) => bytes[i] === b)
}

type Entry = { name: string; type: number; start: number; size: number }

/** The streams of a compound file, read by name (case-insensitive, as the format compares them). */
export class CompoundFile {
  private readonly view: DataView
  private readonly sectorSize: number
  private readonly miniSectorSize: number
  private readonly miniCutoff: number
  private readonly fat: Uint32Array
  private readonly miniFat: Uint32Array
  private readonly entries: Entry[]
  private miniStream: Uint8Array | null = null

  constructor(private readonly bytes: Uint8Array) {
    if (!isCompoundFile(bytes)) throw new CfbError("Not a compound file.")
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const v = this.view
    const sectorShift = v.getUint16(0x1e, true)
    const miniShift = v.getUint16(0x20, true)
    if (sectorShift !== 9 && sectorShift !== 12) throw new CfbError(`Unexpected sector size 2^${sectorShift}.`)
    this.sectorSize = 1 << sectorShift
    this.miniSectorSize = 1 << miniShift
    this.miniCutoff = v.getUint32(0x38, true) || 4096

    // The DIFAT: the first 109 FAT sector numbers in the header, the rest in a chain of DIFAT sectors.
    const fatCount = v.getUint32(0x2c, true)
    const fatSectors: number[] = []
    for (let i = 0; i < 109 && fatSectors.length < fatCount; i++) fatSectors.push(v.getUint32(0x4c + i * 4, true))
    let difat = v.getUint32(0x44, true)
    const perDifat = this.sectorSize / 4 - 1
    for (let guard = 0; difat !== END_OF_CHAIN && difat !== FREE && fatSectors.length < fatCount; guard++) {
      if (guard > this.sectorCount) throw new CfbError("The DIFAT chain loops.")
      const at = this.offsetOf(difat)
      for (let i = 0; i < perDifat && fatSectors.length < fatCount; i++) fatSectors.push(v.getUint32(at + i * 4, true))
      difat = v.getUint32(at + perDifat * 4, true)
    }
    const perSector = this.sectorSize / 4
    this.fat = new Uint32Array(fatSectors.length * perSector)
    fatSectors.forEach((sector, i) => {
      const at = this.offsetOf(sector)
      for (let j = 0; j < perSector; j++) this.fat[i * perSector + j] = v.getUint32(at + j * 4, true)
    })

    const miniFatBytes = this.chain(v.getUint32(0x3c, true))
    this.miniFat = new Uint32Array(Math.floor(miniFatBytes.length / 4))
    const mv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength)
    for (let i = 0; i < this.miniFat.length; i++) this.miniFat[i] = mv.getUint32(i * 4, true)

    const dir = this.chain(v.getUint32(0x30, true))
    const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength)
    this.entries = []
    for (let at = 0; at + 128 <= dir.length; at += 128) {
      const nameLength = Math.min(64, dv.getUint16(at + 0x40, true))
      let name = ""
      for (let i = 0; i + 2 < nameLength; i += 2) name += String.fromCharCode(dv.getUint16(at + i, true))
      this.entries.push({
        name,
        type: dir[at + 0x42],
        start: dv.getUint32(at + 0x74, true),
        // Version 3 files keep only the low 32 bits meaningful.
        size: dv.getUint32(at + 0x78, true),
      })
    }
  }

  private get sectorCount() {
    return Math.ceil(this.bytes.length / this.sectorSize)
  }

  private offsetOf(sector: number) {
    const at = (sector + 1) * this.sectorSize
    if (at >= this.bytes.length) throw new CfbError(`Sector ${sector} is past the end.`)
    return at
  }

  /** The bytes of a chain of regular sectors, cut to `size` when given. */
  private chain(start: number, size?: number): Uint8Array {
    const parts: number[] = []
    for (let s = start, guard = 0; s !== END_OF_CHAIN && s !== FREE; guard++) {
      if (guard > this.fat.length || s >= this.fat.length) throw new CfbError("A sector chain is broken.")
      parts.push(s)
      s = this.fat[s]
    }
    const out = new Uint8Array(parts.length * this.sectorSize)
    parts.forEach((s, i) => {
      const at = this.offsetOf(s)
      out.set(this.bytes.subarray(at, Math.min(at + this.sectorSize, this.bytes.length)), i * this.sectorSize)
    })
    return size === undefined ? out : out.subarray(0, Math.min(size, out.length))
  }

  private miniChain(start: number, size: number): Uint8Array {
    if (!this.miniStream) {
      const root = this.entries.find((e) => e.type === 5)
      if (!root) throw new CfbError("The file has no root entry.")
      this.miniStream = this.chain(root.start, root.size)
    }
    const out = new Uint8Array(size)
    let written = 0
    for (let s = start, guard = 0; s !== END_OF_CHAIN && s !== FREE && written < size; guard++) {
      if (guard > this.miniFat.length || s >= this.miniFat.length) throw new CfbError("A mini sector chain is broken.")
      const at = s * this.miniSectorSize
      const piece = this.miniStream.subarray(at, Math.min(at + this.miniSectorSize, at + size - written))
      out.set(piece, written)
      written += piece.length
      s = this.miniFat[s]
    }
    return out.subarray(0, written)
  }

  /** The names of the streams. */
  names(): string[] {
    return this.entries.filter((e) => e.type === 2).map((e) => e.name)
  }

  /** A stream's bytes, or null when there is none of that name. */
  stream(name: string): Uint8Array | null {
    const upper = name.toUpperCase()
    const entry = this.entries.find((e) => e.type === 2 && e.name.toUpperCase() === upper)
    if (!entry) return null
    return entry.size < this.miniCutoff ? this.miniChain(entry.start, entry.size) : this.chain(entry.start, entry.size)
  }
}
