/**
 * Text encodings for CSV and text files, with Excel's "File origin" names. Reading looks at the byte-order mark,
 * then tries UTF-8 and falls back to Windows-1252 (Excel itself reads a UTF-8 file without a mark as Windows-1252,
 * which garbles accents; Cells checks the bytes instead).
 */
export type TextEncoding = "utf-8" | "windows-1252" | "utf-16le" | "utf-16be"

export const TEXT_ENCODINGS: { id: TextEncoding; label: string }[] = [
  { id: "utf-8", label: "65001 : Unicode (UTF-8)" },
  { id: "windows-1252", label: "1252 : Western European (Windows)" },
  { id: "utf-16le", label: "1200 : Unicode" },
  { id: "utf-16be", label: "1201 : Unicode (Big-Endian)" },
]

export type DecodedText = { text: string; encoding: TextEncoding; bom: boolean }

/** The encoding of a text file's bytes, and whether it starts with a byte-order mark. */
export function sniffEncoding(bytes: Uint8Array): { encoding: TextEncoding; bom: boolean } {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { encoding: "utf-8", bom: true }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { encoding: "utf-16le", bom: true }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { encoding: "utf-16be", bom: true }
  // UTF-16 without a mark: ASCII text has every other byte zero.
  if (bytes.length >= 4 && bytes[0] !== 0 && bytes[1] === 0 && bytes[2] !== 0 && bytes[3] === 0)
    return { encoding: "utf-16le", bom: false }
  if (bytes.length >= 4 && bytes[0] === 0 && bytes[1] !== 0 && bytes[2] === 0 && bytes[3] !== 0)
    return { encoding: "utf-16be", bom: false }
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes.subarray(0, 1 << 20))
    return { encoding: "utf-8", bom: false }
  } catch {
    // A cut multi-byte character at the 1 MB boundary also lands here: check the whole file then.
    if (bytes.length > 1 << 20) {
      try {
        new TextDecoder("utf-8", { fatal: true }).decode(bytes)
        return { encoding: "utf-8", bom: false }
      } catch {
        // Not UTF-8.
      }
    }
    return { encoding: "windows-1252", bom: false }
  }
}

const BOM_LENGTH: Record<TextEncoding, number> = { "utf-8": 3, "windows-1252": 0, "utf-16le": 2, "utf-16be": 2 }

/** The file's text in `encoding` (detected when not given); a byte-order mark is not part of the text. */
export function decodeText(bytes: Uint8Array, encoding?: TextEncoding): DecodedText {
  const sniffed = sniffEncoding(bytes)
  const chosen = encoding ?? sniffed.encoding
  const bom = sniffed.bom && sniffed.encoding === chosen
  const body = bom ? bytes.subarray(BOM_LENGTH[chosen]) : bytes
  let text: string
  if (chosen === "utf-16be") {
    const swapped = new Uint8Array(body.length - (body.length % 2))
    for (let i = 0; i + 1 < body.length; i += 2) {
      swapped[i] = body[i + 1]
      swapped[i + 1] = body[i]
    }
    text = new TextDecoder("utf-16le").decode(swapped)
  } else if (chosen === "windows-1252") text = decode1252(body)
  else text = new TextDecoder(chosen).decode(body)
  // A decoder may keep a mark the sniffing didn't remove (a file read with another encoding than its own).
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1)
  return { text, encoding: chosen, bom }
}

// Windows-1252's 0x80–0x9F, where it differs from Latin-1 (0 = unused).
const HIGH = [
  0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0, 0x017d, 0, 0,
  0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0, 0x017e, 0x0178,
]
const FROM_UNICODE = new Map(HIGH.map((code, i) => [code, 0x80 + i] as const).filter(([code]) => code !== 0))

function decode1252(bytes: Uint8Array): string {
  const parts: string[] = []
  const CHUNK = 8192
  for (let start = 0; start < bytes.length; start += CHUNK) {
    const codes: number[] = []
    for (let i = start; i < Math.min(bytes.length, start + CHUNK); i++) {
      const b = bytes[i]
      codes.push(b >= 0x80 && b < 0xa0 ? HIGH[b - 0x80] || b : b)
    }
    parts.push(String.fromCharCode(...codes))
  }
  return parts.join("")
}

/** Whether Windows-1252 can hold every character of `text`. */
export function fits1252(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i)
    if (c >= 0x80 && !(c >= 0xa0 && c <= 0xff) && !FROM_UNICODE.has(c)) return false
  }
  return true
}

/** `text` as bytes; a character Windows-1252 lacks becomes "?" (one per UTF-16 unit, as Excel writes them). */
export function encodeText(text: string, encoding: TextEncoding, bom: boolean): Uint8Array {
  if (encoding === "utf-8") {
    const body = new TextEncoder().encode(text)
    if (!bom) return body
    const out = new Uint8Array(body.length + 3)
    out.set([0xef, 0xbb, 0xbf])
    out.set(body, 3)
    return out
  }
  if (encoding === "windows-1252") {
    const out = new Uint8Array(text.length)
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i)
      out[i] = c < 0x80 || (c >= 0xa0 && c <= 0xff) ? c : (FROM_UNICODE.get(c) ?? 0x3f)
    }
    return out
  }
  const lead = bom ? 2 : 0
  const out = new Uint8Array(lead + text.length * 2)
  const little = encoding === "utf-16le"
  if (bom) out.set(little ? [0xff, 0xfe] : [0xfe, 0xff])
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i)
    out[lead + i * 2] = little ? c & 0xff : c >> 8
    out[lead + i * 2 + 1] = little ? c >> 8 : c & 0xff
  }
  return out
}
