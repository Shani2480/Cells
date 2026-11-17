import type { Hyperlink } from "@cells/core"
import { Cursor } from "./biff"

/**
 * HLINK (0x01B8) and HLINKTOOLTIP (0x0800) records of an .xls sheet ([MS-XLS] 2.4.140, [MS-OSHARED] 2.3.7.1): a
 * range, then a Hyperlink Object whose flags say which parts follow. Addresses come as a URL moniker (web and mail
 * addresses), a file moniker (paths, relative ones with "..\" counted in cAnti), or plain strings.
 */

const URL_MONIKER = "e0c9ea79f9bace118c8200aa004ba90b"
const FILE_MONIKER = "0303000000000000c000000000000046"

const HAS_MONIKER = 0x01
const HAS_LOCATION = 0x08
const HAS_DISPLAY_NAME = 0x10
const HAS_GUID = 0x20
const HAS_CREATION_TIME = 0x40
const HAS_FRAME_NAME = 0x80
const MONIKER_AS_STRING = 0x100

const hex = (bytes: Uint8Array) => [...bytes].map((b) => b.toString(16).padStart(2, "0")).join("")

/** A UTF-16LE string of `bytes` bytes, up to its first null. */
function utf16(c: Cursor, bytes: number): string {
  const data = c.bytes(bytes)
  let out = ""
  for (let i = 0; i + 1 < data.length; i += 2) {
    const code = data[i] | (data[i + 1] << 8)
    if (code === 0) break
    out += String.fromCharCode(code)
  }
  return out
}

/** HyperlinkString: a character count (with the null) and the characters. */
const linkString = (c: Cursor) => utf16(c, c.u32() * 2)

function readMoniker(c: Cursor): string | null {
  const clsid = hex(c.bytes(16))
  if (clsid === URL_MONIKER) {
    const length = c.u32()
    return utf16(c, length)
  }
  if (clsid === FILE_MONIKER) {
    const up = c.u16()
    const ansiLength = c.u32()
    const ansi = c.bytes(ansiLength)
    c.skip(2 + 2 + 16 + 4)
    const unicodeSize = c.u32()
    let path = ""
    if (unicodeSize > 0) {
      const bytes = c.u32()
      c.skip(2)
      path = utf16(c, bytes)
    } else {
      for (const b of ansi) {
        if (b === 0) break
        path += String.fromCharCode(b)
      }
    }
    return "..\\".repeat(up) + path
  }
  return null
}

/** One HLINK record's link, or null when it can't be read. */
export function readHlink(data: Uint8Array): Hyperlink | null {
  const c = new Cursor(data)
  const r0 = c.u16()
  const r1 = c.u16()
  const c0 = c.u16()
  const c1 = c.u16()
  c.skip(16)
  c.u32()
  const flags = c.u32()
  const link: Hyperlink = { range: { r0, r1, c0, c1 } }
  if (flags & HAS_DISPLAY_NAME) {
    const display = linkString(c)
    if (display) link.display = display
  }
  if (flags & HAS_FRAME_NAME) linkString(c)
  if (flags & HAS_MONIKER) {
    const address = flags & MONIKER_AS_STRING ? linkString(c) : readMoniker(c)
    if (address === null) return null
    if (address) link.address = address
  }
  if (flags & HAS_LOCATION) {
    const location = linkString(c)
    if (location) link.location = location
  }
  if (flags & HAS_GUID) c.skip(16)
  if (flags & HAS_CREATION_TIME) c.skip(8)
  return link.address || link.location ? link : null
}

/** An HLINKTOOLTIP record: the range it belongs to and the ScreenTip. */
export function readHlinkTooltip(data: Uint8Array): { r0: number; c0: number; tooltip: string } | null {
  const c = new Cursor(data)
  // frtRefHeaderNoGrbit: the record type again, then the range.
  c.skip(2)
  const r0 = c.u16()
  c.u16()
  const c0 = c.u16()
  c.u16()
  const tooltip = utf16(c, c.left)
  return tooltip ? { r0, c0, tooltip } : null
}
