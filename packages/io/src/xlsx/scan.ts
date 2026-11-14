/**
 * A lean XML scanner for worksheet parts, which hold nearly all of a big file's bytes. It reports start tags (local
 * name, attributes by local name, offsets), end tags and decoded text, and skips comments, processing instructions and
 * CDATA markers. It does not resolve namespaces: callers match local names, and slice elements whose namespaces matter
 * out of the source to parse properly. Malformed markup (an unclosed tag, a DOCTYPE) is an error.
 */

export type ScanTag = {
  local: string
  /** Attributes as written: qualified name → decoded value. */
  attrs: Map<string, string>
  start: number
  end: number
  selfClosing: boolean
}

export class ScanError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "ScanError"
  }
}

const ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }

export function decodeEntities(s: string): string {
  if (s.indexOf("&") < 0) return s
  return s.replace(/&(#x[0-9A-Fa-f]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e: string) => {
    if (e[0] !== "#") return ENTITIES[e]
    const code = e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
    return Number.isFinite(code) && code <= 0x10ffff ? String.fromCodePoint(code) : m
  })
}

const localOf = (name: string) => {
  const i = name.indexOf(":")
  return i < 0 ? name : name.slice(i + 1)
}

const ATTR = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

/** Where a scan stopped: it can go on from here (`scanXml` with a budget). */
export type ScanState = { pos: number; depth: number }

/**
 * Scans `xml` from `state.pos`. With a `budget` (characters), it stops at the first tag after that many and returns
 * false; call again with the same state to go on. Returns true when the whole part is done.
 */
export function scanXml(
  xml: string,
  handlers: { start: (tag: ScanTag) => void; end: (local: string, end: number) => void; text: (text: string) => void },
  state: ScanState = { pos: 0, depth: 0 },
  budget = Infinity,
): boolean {
  const n = xml.length
  let i = state.pos
  let depth = state.depth
  const stop = i + budget
  while (i < n) {
    if (i > stop) {
      state.pos = i
      state.depth = depth
      return false
    }
    const lt = xml.indexOf("<", i)
    if (lt < 0) {
      if (depth > 0) throw new ScanError("The part ends inside an element.")
      return true
    }
    if (lt > i && depth > 0) handlers.text(decodeEntities(xml.slice(i, lt)))
    const c = xml.charCodeAt(lt + 1)
    if (c === 33 /* ! */) {
      if (xml.startsWith("<!--", lt)) {
        const e = xml.indexOf("-->", lt + 4)
        if (e < 0) throw new ScanError("Unclosed comment.")
        i = e + 3
      } else if (xml.startsWith("<![CDATA[", lt)) {
        const e = xml.indexOf("]]>", lt + 9)
        if (e < 0) throw new ScanError("Unclosed CDATA section.")
        if (depth > 0) handlers.text(xml.slice(lt + 9, e))
        i = e + 3
      } else throw new ScanError("DOCTYPE is not allowed in Office XML.")
      continue
    }
    if (c === 63 /* ? */) {
      const e = xml.indexOf("?>", lt + 2)
      if (e < 0) throw new ScanError("Unclosed processing instruction.")
      i = e + 2
      continue
    }
    const gt = tagEnd(xml, lt)
    if (gt < 0) throw new ScanError("Unclosed tag.")
    if (c === 47 /* / */) {
      handlers.end(localOf(xml.slice(lt + 2, gt).trim()), gt + 1)
      depth--
      i = gt + 1
      continue
    }
    const selfClosing = xml.charCodeAt(gt - 1) === 47
    const bodyEnd = selfClosing ? gt - 1 : gt
    let nameEnd = lt + 1
    while (nameEnd < bodyEnd) {
      const ch = xml.charCodeAt(nameEnd)
      if (ch === 32 || ch === 9 || ch === 10 || ch === 13) break
      nameEnd++
    }
    const local = localOf(xml.slice(lt + 1, nameEnd))
    const attrs = new Map<string, string>()
    if (nameEnd < bodyEnd) {
      const body = xml.slice(nameEnd, bodyEnd)
      ATTR.lastIndex = 0
      let m: RegExpExecArray | null
      while ((m = ATTR.exec(body))) attrs.set(m[1], decodeEntities(m[2] ?? m[3]))
    }
    handlers.start({ local, attrs, start: lt, end: gt + 1, selfClosing })
    if (selfClosing) handlers.end(local, gt + 1)
    else depth++
    i = gt + 1
  }
  if (depth > 0) throw new ScanError("The part ends inside an element.")
  state.pos = n
  state.depth = 0
  return true
}

/** The ">" closing the tag that starts at `lt`, skipping ">" inside quoted attribute values. */
function tagEnd(xml: string, lt: number): number {
  const n = xml.length
  let quote = 0
  for (let i = lt + 1; i < n; i++) {
    const c = xml.charCodeAt(i)
    if (quote) {
      if (c === quote) quote = 0
    } else if (c === 34 || c === 39) quote = c
    else if (c === 62) return i
  }
  return -1
}
