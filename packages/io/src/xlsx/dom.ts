import { decodeXml, escapeXml, parseXmlEvents } from "../xml"

/**
 * A light element tree for the small parts (workbook, styles, theme, comments). Each node remembers where it sits
 * in the source text, so an element Cells doesn't understand can be written back exactly as it was read.
 */
export type XNode = {
  /** Local name, without the prefix. */
  local: string
  /** Qualified name as written (`x14ac:dyDescent`, `c`). */
  name: string
  uri: string
  /** Attributes by qualified name. */
  attrs: Record<string, string>
  /** Attributes by namespace URI and local name (`uri|local`; no namespace: just `local`). */
  nsAttrs: Record<string, string>
  /** Namespace declarations on this element (prefix → uri; "" for the default). */
  declares: Record<string, string>
  children: XNode[]
  text: string
  start: number
  /** Just after the start tag's ">". */
  tagEnd: number
  end: number
}

export type XDoc = { root: XNode; source: string }

export function parseTree(source: string | Uint8Array): XDoc {
  const xml = typeof source === "string" ? source : decodeXml(source)
  const stack: XNode[] = []
  let root: XNode | null = null
  parseXmlEvents(xml, {
    start(tag) {
      const attrs: Record<string, string> = {}
      const nsAttrs: Record<string, string> = {}
      for (const a of tag.attributes) {
        attrs[a.name] = a.value
        if (a.prefix !== "xmlns" && a.name !== "xmlns") nsAttrs[a.uri ? `${a.uri}|${a.local}` : a.local] = a.value
      }
      const node: XNode = {
        local: tag.local,
        name: tag.prefix ? `${tag.prefix}:${tag.local}` : tag.local,
        uri: tag.uri,
        attrs,
        nsAttrs,
        declares: tag.namespaces,
        children: [],
        text: "",
        start: tag.start,
        tagEnd: tag.end,
        end: tag.end,
      }
      const parent = stack[stack.length - 1]
      if (parent) parent.children.push(node)
      else root = node
      stack.push(node)
    },
    end(tag) {
      const node = stack.pop()
      if (node) node.end = tag.end
    },
    text(text) {
      const node = stack[stack.length - 1]
      if (node) node.text += text
    },
  })
  if (!root) throw new Error("The part has no root element.")
  return { root, source: xml }
}

/**
 * Parses one element cut out of a larger part, with the namespace declarations of that part's root (`rootTag`),
 * so prefixes declared there (`r:id`) resolve. The returned source is the wrapped text; offsets refer to it.
 */
export function parseFragment(fragment: string, rootTag: string): XDoc {
  const decls = rootTag.match(/\sxmlns(?::[\w.-]+)?="[^"]*"/g)?.join("") ?? ""
  const doc = parseTree(`<cells-fragment${decls}>${fragment}</cells-fragment>`)
  const root = doc.root.children[0]
  if (!root) throw new Error("The fragment has no element.")
  return { root, source: doc.source }
}

export const kids = (node: XNode | undefined, local: string) => node?.children.filter((c) => c.local === local) ?? []
export const kid = (node: XNode | undefined, local: string) => node?.children.find((c) => c.local === local)
/** An attribute by its plain name (no namespace), or with a namespace URI. */
export const at = (node: XNode | undefined, local: string, uri?: string) =>
  node?.nsAttrs[uri ? `${uri}|${local}` : local]

/** The element's source text, exactly as read. */
export const raw = (doc: XDoc, node: XNode) => doc.source.slice(node.start, node.end)

/** The element's start tag as read (with its namespace declarations), without the closing ">" or "/>". */
export const startTag = (doc: XDoc, node: XNode) => doc.source.slice(node.start, node.tagEnd).replace(/\/?>$/, "")

/** All text below the node (shared strings with runs: the runs' `t` joined). */
export function textOf(node: XNode, local = "t"): string {
  if (node.local === local) return node.text
  let out = ""
  for (const c of node.children) {
    // Phonetic runs (rPh) are reading aids, not part of the text.
    if (c.local === "rPh" || c.local === "phoneticPr") continue
    out += textOf(c, local)
  }
  return out
}

/** Attributes as XML (`name="value"` pairs, leading space each). */
export function attrText(attrs: Record<string, string | number | boolean | undefined | null>): string {
  let out = ""
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue
    out += ` ${k}="${escapeXml(v === true ? "1" : String(v))}"`
  }
  return out
}

/** An element as read, kept to be written back: its local name and source text. */
export type Kept = { local: string; xml: string }

/**
 * Children of a rebuilt element, in the schema's order: elements Cells writes itself (`generated`, by local name; an
 * empty string drops one) replace the source's at their place; the source's other children stay where they were,
 * exactly as read; generated elements the source didn't have go where the schema puts them.
 */
export function assemble(order: readonly string[], source: readonly Kept[], generated: Record<string, string>): string {
  const rank = (local: string) => {
    const i = order.indexOf(local)
    return i < 0 ? -1 : i
  }
  const out: Kept[] = []
  const placed = new Set<string>()
  for (const child of source) {
    if (child.local in generated) {
      if (placed.has(child.local)) continue
      placed.add(child.local)
      if (generated[child.local]) out.push({ local: child.local, xml: generated[child.local] })
    } else out.push(child)
  }
  for (const [local, xml] of Object.entries(generated)) {
    if (placed.has(local) || !xml) continue
    const r = rank(local)
    // Before the first known element that comes later in the schema (unknown elements keep their neighbours).
    let at = out.length
    for (let i = 0; i < out.length; i++) {
      const ri = rank(out[i].local)
      if (ri > r) {
        at = i
        break
      }
    }
    out.splice(at, 0, { local, xml })
  }
  return out.map((k) => k.xml).join("")
}
