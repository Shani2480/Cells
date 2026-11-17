import { CompoundFile, isCompoundFile } from "./xls/cfb"

/**
 * Insert › Object's embedded files as Office keeps them: a file of any kind goes in an OLE "Package" ([MS-OLEDS]
 * 2.3.6 Ole10Native, in a compound file with its CompObj stream); Office documents are embedded as they are.
 */

const SECTOR = 512
const MINI = 64
const MINI_CUTOFF = 4096
const FREE = 0xffffffff
const END = 0xfffffffe
const FAT_SECT = 0xfffffffd
const NONE = 0xffffffff

/** The Package object's class ({0003000C-0000-0000-C000-000000000046}) as the directory stores it. */
const PACKAGE_CLSID = [0x0c, 0x00, 0x03, 0x00, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46]

type Stream = { name: string; data: Uint8Array }

/** A compound file ([MS-CFB], version 3) holding `streams` under its root (small ones in the mini stream). */
export function writeCompoundFile(streams: Stream[], rootClsid: number[] = []): Uint8Array {
  const small = streams.filter((s) => s.data.length < MINI_CUTOFF)
  const large = streams.filter((s) => s.data.length >= MINI_CUTOFF)
  const sectors = (n: number, size = SECTOR) => Math.ceil(n / size)

  // The mini stream: each small stream in whole 64-byte mini sectors.
  const miniStarts = new Map<Stream, number>()
  let miniCount = 0
  for (const s of small) {
    miniStarts.set(s, s.data.length ? miniCount : END)
    miniCount += sectors(s.data.length, MINI)
  }
  const miniFat = new Uint32Array(Math.max(0, miniCount))
  for (const s of small) {
    const first = miniStarts.get(s)!
    const n = sectors(s.data.length, MINI)
    for (let i = 0; i < n; i++) miniFat[first + i] = i === n - 1 ? END : first + i + 1
  }
  const miniStreamBytes = miniCount * MINI
  const miniFatSectors = sectors(miniFat.length * 4)
  const miniStreamSectors = sectors(miniStreamBytes)
  const dirEntries = streams.length + 1
  const dirSectors = sectors(dirEntries * 128)
  const largeSectors = large.reduce((sum, s) => sum + sectors(s.data.length), 0)

  // The FAT covers every sector, its own included (at most 109 FAT sectors: no DIFAT sectors needed here).
  const content = dirSectors + miniFatSectors + miniStreamSectors + largeSectors
  let fatSectors = 1
  while (fatSectors * (SECTOR / 4) < content + fatSectors) fatSectors++
  if (fatSectors > 109) throw new Error("The embedded file is too large.")
  const total = fatSectors + content
  const fat = new Uint32Array(fatSectors * (SECTOR / 4)).fill(FREE)
  let next = 0
  const chain = (count: number): number => {
    if (!count) return END
    const first = next
    for (let i = 0; i < count; i++) fat[next + i] = i === count - 1 ? END : next + i + 1
    next += count
    return first
  }
  for (let i = 0; i < fatSectors; i++) fat[next++] = FAT_SECT
  const dirStart = chain(dirSectors)
  const miniFatStart = chain(miniFatSectors)
  const miniStreamStart = chain(miniStreamSectors)
  const largeStarts = new Map<Stream, number>()
  for (const s of large) largeStarts.set(s, chain(sectors(s.data.length)))

  const out = new Uint8Array(SECTOR * (1 + total))
  const v = new DataView(out.buffer)
  // Header.
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  v.setUint16(0x18, 0x003e, true)
  v.setUint16(0x1a, 0x0003, true)
  v.setUint16(0x1c, 0xfffe, true)
  v.setUint16(0x1e, 9, true)
  v.setUint16(0x20, 6, true)
  v.setUint32(0x2c, fatSectors, true)
  v.setUint32(0x30, dirStart, true)
  v.setUint32(0x38, MINI_CUTOFF, true)
  v.setUint32(0x3c, miniFatSectors ? miniFatStart : END, true)
  v.setUint32(0x40, miniFatSectors, true)
  v.setUint32(0x44, END, true)
  v.setUint32(0x48, 0, true)
  for (let i = 0; i < 109; i++) v.setUint32(0x4c + i * 4, i < fatSectors ? i : FREE, true)

  const at = (sector: number) => SECTOR * (1 + sector)
  for (let i = 0; i < fat.length; i++) v.setUint32(at(0) + i * 4, fat[i]!, true)

  // Directory: the root, then the streams as a chain of right siblings in the order CFB compares names.
  const order = [...streams].sort((a, b) =>
    a.name.length !== b.name.length
      ? a.name.length - b.name.length
      : a.name.toUpperCase() < b.name.toUpperCase()
        ? -1
        : 1,
  )
  const entry = (
    i: number,
    name: string,
    type: number,
    child: number,
    right: number,
    start: number,
    size: number,
    clsid: number[] = [],
  ) => {
    const o = at(dirStart) + i * 128
    const units = [...name].slice(0, 31)
    units.forEach((ch, k) => v.setUint16(o + k * 2, ch.charCodeAt(0), true))
    v.setUint16(o + 0x40, (units.length + 1) * 2, true)
    out[o + 0x42] = type
    out[o + 0x43] = 1 // black
    v.setUint32(o + 0x44, NONE, true)
    v.setUint32(o + 0x48, right, true)
    v.setUint32(o + 0x4c, child, true)
    out.set(clsid.slice(0, 16), o + 0x50)
    v.setUint32(o + 0x74, start, true)
    v.setUint32(o + 0x78, size, true)
  }
  for (let i = dirEntries; i < dirSectors * 4; i++) {
    const o = at(dirStart) + i * 128
    v.setUint32(o + 0x44, NONE, true)
    v.setUint32(o + 0x48, NONE, true)
    v.setUint32(o + 0x4c, NONE, true)
  }
  entry(
    0,
    "Root Entry",
    5,
    order.length ? 1 : NONE,
    NONE,
    miniCount ? miniStreamStart : END,
    miniStreamBytes,
    rootClsid,
  )
  order.forEach((s, i) => {
    const start = s.data.length < MINI_CUTOFF ? miniStarts.get(s)! : largeStarts.get(s)!
    entry(i + 1, s.name, 2, NONE, i + 1 < order.length ? i + 2 : NONE, start, s.data.length)
  })

  for (let i = 0; i < miniFat.length; i++) v.setUint32(at(miniFatStart) + i * 4, miniFat[i]!, true)
  for (let i = miniFat.length; i < miniFatSectors * (SECTOR / 4); i++) v.setUint32(at(miniFatStart) + i * 4, FREE, true)
  for (const s of small) if (s.data.length) out.set(s.data, at(miniStreamStart) + miniStarts.get(s)! * MINI)
  for (const s of large) out.set(s.data, at(largeStarts.get(s)!))
  return out
}

const ascii = (s: string) => {
  const b = new Uint8Array(s.length + 1)
  for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) < 0x80 ? s.charCodeAt(i) : 0x3f
  return b
}

const utf16 = (s: string) => {
  const b = new Uint8Array(s.length * 2)
  for (let i = 0; i < s.length; i++) {
    b[i * 2] = s.charCodeAt(i) & 0xff
    b[i * 2 + 1] = s.charCodeAt(i) >> 8
  }
  return b
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

const u32 = (n: number) => {
  const b = new Uint8Array(4)
  new DataView(b.buffer).setUint32(0, n, true)
  return b
}

const u16 = (n: number) => {
  const b = new Uint8Array(2)
  new DataView(b.buffer).setUint16(0, n, true)
  return b
}

/** The CompObj stream of a Package object, as Office writes it. */
const PACKAGE_COMPOBJ = concat([
  new Uint8Array([0x01, 0x00, 0xfe, 0xff, 0x03, 0x0a, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff]),
  new Uint8Array(PACKAGE_CLSID),
  u32(12),
  ascii("OLE Package"),
  u32(0),
  u32(8),
  ascii("Package"),
  new Uint8Array([0xf4, 0x39, 0xb2, 0x71]),
  u32(0),
  u32(0),
  u32(0),
])

/** A file wrapped as Office's Package object (an oleObject.bin). */
export function packageObject(name: string, data: Uint8Array): Uint8Array {
  // The file's own name as both its label and paths: no folder of this computer goes into the workbook.
  const native = concat([
    u16(2),
    ascii(name),
    ascii(name),
    u32(0x00030000),
    u32(name.length + 1),
    ascii(name),
    u32(data.length),
    data,
    u32(name.length),
    utf16(name),
    u32(name.length),
    utf16(name),
    u32(name.length),
    utf16(name),
  ])
  const ole10 = concat([u32(native.length), native])
  return writeCompoundFile(
    [
      { name: "\u0001CompObj", data: PACKAGE_COMPOBJ },
      { name: "\u0001Ole10Native", data: ole10 },
    ],
    PACKAGE_CLSID,
  )
}

/** The file inside a Package object: its name and bytes; null when `bin` isn't one. */
export function unpackObject(bin: Uint8Array): { name: string; data: Uint8Array } | null {
  if (!isCompoundFile(bin)) return null
  let native: Uint8Array | null
  try {
    native = new CompoundFile(bin).stream("\u0001Ole10Native")
  } catch {
    return null
  }
  if (!native || native.length < 12) return null
  const v = new DataView(native.buffer, native.byteOffset, native.byteLength)
  let at = 4 + 2
  const cstr = () => {
    const start = at
    while (at < native!.length && native![at] !== 0) at++
    const s = String.fromCharCode(...native!.subarray(start, at))
    at++
    return s
  }
  try {
    const label = cstr()
    cstr() // where it came from
    at += 4
    const tempLength = v.getUint32(at, true)
    at += 4 + tempLength
    const size = v.getUint32(at, true)
    at += 4
    if (at + size > native.length) return null
    const data = native.slice(at, at + size)
    at += size
    // The Unicode label, when there is one, has the name's real spelling.
    let name = label
    const skip = () => {
      const n = v.getUint32(at, true)
      at += 4 + n * 2
      return n
    }
    if (at + 4 <= native.length) {
      skip()
      if (at + 4 <= native.length) {
        const n = v.getUint32(at, true)
        if (n > 0 && at + 4 + n * 2 <= native.length) {
          let s = ""
          for (let i = 0; i < n; i++) s += String.fromCharCode(v.getUint16(at + 4 + i * 2, true))
          name = s
        }
      }
    }
    return { name: name.split(/[\\/]/).pop() || label, data }
  } catch {
    return null
  }
}
