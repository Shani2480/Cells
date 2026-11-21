/**
 * The VBA project of a macro-enabled workbook (vbaProject.bin in .xlsm, _VBA_PROJECT_CUR in .xls), read only to show
 * its modules and macros (View › Macros): the dir stream lists the modules and where their source starts, and the
 * source is stored with MS-OVBA's compression (2.4.1). Cells doesn't run VBA; the project stays in the file as read.
 */
import type { VbaModule } from "@cells/core"
import { CompoundFile } from "./xls/cfb"

/** MS-OVBA 2.4.1: a CompressedContainer back to its bytes. */
export function decompressVba(data: Uint8Array, start = 0): Uint8Array {
  if (data[start] !== 1) throw new Error("Not a compressed VBA container.")
  const out: number[] = []
  let at = start + 1
  while (at + 2 <= data.length) {
    const header = data[at] | (data[at + 1] << 8)
    const size = (header & 0x0fff) + 3
    const compressed = (header & 0x8000) !== 0
    const end = Math.min(data.length, at + size)
    at += 2
    const chunkStart = out.length
    if (!compressed) {
      for (let i = 0; i < 4096 && at < end; i++) out.push(data[at++])
      continue
    }
    while (at < end) {
      const flags = data[at++]
      for (let bit = 0; bit < 8 && at < end; bit++) {
        if ((flags & (1 << bit)) === 0) {
          out.push(data[at++])
          continue
        }
        const token = data[at] | (data[at + 1] << 8)
        at += 2
        const done = out.length - chunkStart
        let bitCount = 4
        while (1 << bitCount < done) bitCount++
        bitCount = Math.max(bitCount, 4)
        const lengthMask = 0xffff >> bitCount
        const offset = (token >> (16 - bitCount)) + 1
        const length = (token & lengthMask) + 3
        for (let i = 0; i < length; i++) out.push(out[out.length - offset])
      }
    }
  }
  return Uint8Array.from(out)
}

/** Code pages VBA projects use, as TextDecoder labels. */
const CODE_PAGES: Record<number, string> = {
  874: "windows-874",
  932: "shift_jis",
  936: "gbk",
  949: "euc-kr",
  950: "big5",
  1250: "windows-1250",
  1251: "windows-1251",
  1252: "windows-1252",
  1253: "windows-1253",
  1254: "windows-1254",
  1255: "windows-1255",
  1256: "windows-1256",
  1257: "windows-1257",
  1258: "windows-1258",
  65001: "utf-8",
}

type ModuleEntry = { name: string; stream: string; offset: number; kind: VbaModule["kind"] }

/** The dir stream (decompressed): the code page and each module's name, stream and source offset. */
function readDir(dir: Uint8Array): { codePage: number; modules: ModuleEntry[] } {
  const v = new DataView(dir.buffer, dir.byteOffset, dir.byteLength)
  let codePage = 1252
  const modules: ModuleEntry[] = []
  let current: Partial<ModuleEntry> & { unicodeName?: string; unicodeStream?: string } = {}
  let at = 0
  const ascii = (from: number, n: number) => new TextDecoder("latin1").decode(dir.subarray(from, from + n))
  const utf16 = (from: number, n: number) => new TextDecoder("utf-16le").decode(dir.subarray(from, from + n))
  while (at + 6 <= dir.length) {
    const id = v.getUint16(at, true)
    const size = v.getUint32(at + 2, true)
    const data = at + 6
    // PROJECTVERSION's size field says 4 but six bytes follow it.
    if (id === 0x0009) {
      at = data + 6
      continue
    }
    if (id === 0x0003) codePage = v.getUint16(data, true)
    else if (id === 0x0019) current = { name: ascii(data, size), kind: "module" }
    else if (id === 0x0047) current.unicodeName = utf16(data, size)
    else if (id === 0x001a) current.stream = ascii(data, size)
    else if (id === 0x0032) current.unicodeStream = utf16(data, size)
    else if (id === 0x0031) current.offset = v.getUint32(data, true)
    else if (id === 0x0021) current.kind = "module"
    else if (id === 0x0022) current.kind = "class"
    else if (id === 0x002b && current.name !== undefined) {
      modules.push({
        name: current.unicodeName || current.name,
        stream: current.unicodeStream || current.stream || current.name,
        offset: current.offset ?? 0,
        kind: current.kind ?? "module",
      })
      current = {}
    }
    at = data + size
  }
  return { codePage, modules }
}

/** The modules of a VBA project in a compound file (vbaProject.bin, or an .xls with its _VBA_PROJECT_CUR storage). */
export function readVbaProject(bytes: Uint8Array): VbaModule[] {
  try {
    const cfb = new CompoundFile(bytes)
    const dirStream = cfb.stream("dir")
    if (!dirStream) return []
    const { codePage, modules } = readDir(decompressVba(dirStream))
    let decoder: TextDecoder
    try {
      decoder = new TextDecoder(CODE_PAGES[codePage] ?? "windows-1252")
    } catch {
      decoder = new TextDecoder("windows-1252")
    }
    const out: VbaModule[] = []
    for (const m of modules) {
      const stream = cfb.stream(m.stream)
      if (!stream || m.offset >= stream.length) continue
      const source = decoder.decode(decompressVba(stream, m.offset))
      const code = source
        .split(/\r\n|\r|\n/)
        .filter((line) => !/^Attribute VB_/i.test(line))
        .join("\n")
        .replace(/^\n+|\n+$/g, "")
      // Document modules (ThisWorkbook, Sheet1) carry VB_Base / predeclared attributes; treat their kind as such.
      const kind = /^Attribute VB_Base = "0\{000208/im.test(source) ? "document" : m.kind
      out.push({ name: m.name, kind, code })
    }
    return out
  } catch {
    return []
  }
}

/**
 * The macros View › Macros lists: Subs without parameters that aren't Private, as "Module.Sub" for document and class
 * modules and plain "Sub" for standard modules, as Excel names them.
 */
export function vbaMacros(modules: readonly VbaModule[]): { name: string; module: string }[] {
  const out: { name: string; module: string }[] = []
  for (const m of modules) {
    if (m.kind === "class") continue
    for (const match of m.code.matchAll(
      /^[ \t]*(?:Public[ \t]+)?(?:Static[ \t]+)?Sub[ \t]+([A-Za-z_]\w*)[ \t]*\([ \t]*\)/gim,
    ))
      out.push({ name: m.kind === "document" ? `${m.name}.${match[1]}` : match[1], module: m.name })
  }
  return out
}
