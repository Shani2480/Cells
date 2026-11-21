/**
 * FILTERXML: the values an XPath finds in an XML text, one per row. A small XML reader and the common part of XPath 1.0
 * (child and descendant steps, `.`, `..`, `*`, `@attr`, `text()`, predicates with positions, last(), position(),
 * comparisons, and/or, not(), contains(), starts-with()). An expression that isn't a list of nodes (count(…)), XML that
 * isn't well formed and a path that finds nothing are #VALUE!, as in Excel.
 */
import { isError } from "../../values"
import { toText, VALUE } from "../coerce"
import { matrix, type Scalar } from "../types"
import { scalarFn } from "./registry"

type El = { kind: "el"; name: string; attrs: Map<string, string>; children: XNode[]; parent: El | null }
type Txt = { kind: "text"; value: string; parent: El }
type Attr = { kind: "attr"; name: string; value: string; parent: El }
type XNode = El | Txt
type Item = El | Txt | Attr

const ENTITIES: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" }

function decode(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (_m, e: string) => {
    if (e[0] === "#")
      return String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : Number(e.slice(1)))
    const v = ENTITIES[e]
    if (v === undefined) throw new Error("entity")
    return v
  })
}

/** The document's root element; throws when the text isn't well-formed XML. */
function parseXml(src: string): El {
  const doc: El = { kind: "el", name: "#document", attrs: new Map(), children: [], parent: null }
  let cur = doc
  let i = 0
  const name = /[A-Za-z_:][\w:.-]*/y
  while (i < src.length) {
    const lt = src.indexOf("<", i)
    const textEnd = lt < 0 ? src.length : lt
    if (textEnd > i) {
      const t = src.slice(i, textEnd)
      if (cur === doc) {
        if (t.trim()) throw new Error("text outside the root")
      } else cur.children.push({ kind: "text", value: decode(t), parent: cur })
    }
    if (lt < 0) break
    i = lt
    if (src.startsWith("<!--", i)) {
      const e = src.indexOf("-->", i)
      if (e < 0) throw new Error("comment")
      i = e + 3
    } else if (src.startsWith("<![CDATA[", i)) {
      const e = src.indexOf("]]>", i)
      if (e < 0 || cur === doc) throw new Error("cdata")
      cur.children.push({ kind: "text", value: src.slice(i + 9, e), parent: cur })
      i = e + 3
    } else if (src.startsWith("<?", i) || src.startsWith("<!", i)) {
      const e = src.indexOf(">", i)
      if (e < 0) throw new Error("declaration")
      i = e + 1
    } else if (src[i + 1] === "/") {
      name.lastIndex = i + 2
      const m = name.exec(src)
      if (!m || m[0] !== cur.name) throw new Error("mismatched end tag")
      i = name.lastIndex
      while (/\s/.test(src[i] ?? "")) i++
      if (src[i] !== ">") throw new Error("end tag")
      i++
      cur = cur.parent!
    } else {
      name.lastIndex = i + 1
      const m = name.exec(src)
      if (!m) throw new Error("tag")
      if (cur === doc && doc.children.length) throw new Error("two roots")
      const el: El = { kind: "el", name: m[0], attrs: new Map(), children: [], parent: cur }
      cur.children.push(el)
      i = name.lastIndex
      for (;;) {
        while (/\s/.test(src[i] ?? "")) i++
        if (src.startsWith("/>", i)) {
          i += 2
          break
        }
        if (src[i] === ">") {
          i++
          cur = el
          break
        }
        name.lastIndex = i
        const a = name.exec(src)
        if (!a) throw new Error("attribute")
        i = name.lastIndex
        while (/\s/.test(src[i] ?? "")) i++
        if (src[i] !== "=") throw new Error("attribute")
        i++
        while (/\s/.test(src[i] ?? "")) i++
        const q = src[i]
        if (q !== '"' && q !== "'") throw new Error("attribute")
        const e = src.indexOf(q, i + 1)
        if (e < 0) throw new Error("attribute")
        el.attrs.set(a[0], decode(src.slice(i + 1, e)))
        i = e + 1
      }
    }
  }
  if (cur !== doc || doc.children.length !== 1) throw new Error("unclosed")
  return doc
}

function stringValue(n: Item): string {
  if (n.kind !== "el") return n.value
  return n.children.map(stringValue).join("")
}

// XPath: tokens, then a recursive-descent evaluator over node lists.
type Tok = { t: "num"; v: number } | { t: "str"; v: string } | { t: "name"; v: string } | { t: "op"; v: string }

function tokenize(src: string): Tok[] {
  const out: Tok[] = []
  const re =
    /\s*(?:(\d+(?:\.\d*)?|\.\d+)|"([^"]*)"|'([^']*)'|(\/\/|\/|\.\.|\.|@|\*|\[|\]|\(|\)|,|\||!=|<=|>=|=|<|>|\+|-|::)|([A-Za-z_][\w.:-]*?)(?=::|\s|[^\w.:-]|$))/y
  let i = 0
  while (i < src.length) {
    if (/^\s*$/.test(src.slice(i))) break
    re.lastIndex = i
    const m = re.exec(src)
    if (!m) throw new Error("xpath")
    i = re.lastIndex
    if (m[1] !== undefined) out.push({ t: "num", v: Number(m[1]) })
    else if (m[2] !== undefined || m[3] !== undefined) out.push({ t: "str", v: m[2] ?? m[3] })
    else if (m[4] !== undefined) out.push({ t: "op", v: m[4] })
    else out.push({ t: "name", v: m[5] })
  }
  return out
}

type XValue = Item[] | string | number | boolean

class XPath {
  private i = 0
  constructor(private readonly toks: Tok[]) {}

  private peek(v?: string): boolean {
    const t = this.toks[this.i]
    return t !== undefined && (v === undefined || ((t.t === "op" || t.t === "name") && t.v === v))
  }
  private take(v: string) {
    if (!this.peek(v)) throw new Error("xpath")
    this.i++
  }
  get done() {
    return this.i >= this.toks.length
  }

  /** Parses an expression into a function of (context node, position, size). */
  expr(): (n: Item, pos: number, size: number) => XValue {
    return this.or()
  }
  private or() {
    let l = this.and()
    while (this.peek("or")) {
      this.i++
      const a = l
      const b = this.and()
      l = (n, p, s) => bool(a(n, p, s)) || bool(b(n, p, s))
    }
    return l
  }
  private and() {
    let l = this.compare()
    while (this.peek("and")) {
      this.i++
      const a = l
      const b = this.compare()
      l = (n, p, s) => bool(a(n, p, s)) && bool(b(n, p, s))
    }
    return l
  }
  private compare() {
    let l = this.additive()
    for (const op of ["=", "!=", "<=", ">=", "<", ">"]) {
      if (this.peek(op)) {
        this.i++
        const a = l
        const b = this.additive()
        l = (n, p, s) => compareValues(a(n, p, s), b(n, p, s), op)
        break
      }
    }
    return l
  }
  private additive() {
    let l = this.union()
    while (this.peek("+") || this.peek("-")) {
      const op = (this.toks[this.i++] as { v: string }).v
      const a = l
      const b = this.union()
      l = (n, p, s) => (op === "+" ? number(a(n, p, s)) + number(b(n, p, s)) : number(a(n, p, s)) - number(b(n, p, s)))
    }
    return l
  }
  private union() {
    let l = this.primary()
    while (this.peek("|")) {
      this.i++
      const a = l
      const b = this.primary()
      l = (n, p, s) => {
        const x = a(n, p, s)
        const y = b(n, p, s)
        if (!Array.isArray(x) || !Array.isArray(y)) throw new Error("union")
        return documentOrder([...new Set([...x, ...y])])
      }
    }
    return l
  }
  private primary(): (n: Item, pos: number, size: number) => XValue {
    const t = this.toks[this.i]
    if (!t) throw new Error("xpath")
    if (t.t === "num") {
      this.i++
      return () => t.v
    }
    if (t.t === "str") {
      this.i++
      return () => t.v
    }
    if (t.t === "op" && t.v === "(") {
      this.i++
      const e = this.expr()
      this.take(")")
      return e
    }
    if (t.t === "op" && t.v === "-") {
      this.i++
      const e = this.primary()
      return (n, p, s) => -number(e(n, p, s))
    }
    const next = this.toks[this.i + 1]
    if (t.t === "name" && next?.t === "op" && next.v === "(" && !["text", "node"].includes(t.v)) {
      this.i += 2
      const args: ((n: Item, p: number, s: number) => XValue)[] = []
      while (!this.peek(")")) {
        args.push(this.expr())
        if (this.peek(",")) this.i++
      }
      this.take(")")
      return fn(t.v, args)
    }
    return this.path()
  }

  private path(): (n: Item, pos: number, size: number) => Item[] {
    let absolute = false
    const steps: {
      descendant: boolean
      step: (n: Item) => Item[]
      preds: ((n: Item, p: number, s: number) => XValue)[]
    }[] = []
    let descendant = false
    if (this.peek("/") || this.peek("//")) {
      absolute = true
      descendant = this.peek("//")
      this.i++
      if (!descendant && (this.done || !this.stepStarts())) {
        return (n) => [rootOf(n)]
      }
    }
    for (;;) {
      const step = this.step()
      const preds: ((n: Item, p: number, s: number) => XValue)[] = []
      while (this.peek("[")) {
        this.i++
        preds.push(this.expr())
        this.take("]")
      }
      steps.push({ descendant, step, preds })
      if (this.peek("/") || this.peek("//")) {
        descendant = this.peek("//")
        this.i++
      } else break
    }
    return (n) => {
      let nodes: Item[] = [absolute ? rootOf(n) : n]
      for (const s of steps) {
        const next: Item[] = []
        for (const c of nodes) {
          const from = s.descendant ? [c, ...descendants(c)] : [c]
          for (const f of from) {
            let found = s.step(f)
            for (const pred of s.preds) {
              const pool = found
              found = pool.filter((x, k) => {
                const v = pred(x, k + 1, pool.length)
                return typeof v === "number" ? v === k + 1 : bool(v)
              })
            }
            for (const x of found) if (!next.includes(x)) next.push(x)
          }
        }
        nodes = next
      }
      return nodes
    }
  }

  private stepStarts() {
    const t = this.toks[this.i]
    return t.t === "name" || (t.t === "op" && ["@", "*", ".", ".."].includes(t.v))
  }

  private step(): (n: Item) => Item[] {
    if (this.peek(".")) {
      this.i++
      return (n) => [n]
    }
    if (this.peek("..")) {
      this.i++
      return (n) => (n.parent && n.parent.name !== "#document" ? [n.parent] : n.parent ? [n.parent] : [])
    }
    let axis = "child"
    if (this.peek("@")) {
      this.i++
      axis = "attribute"
    } else if (
      this.toks[this.i]?.t === "name" &&
      this.toks[this.i + 1]?.t === "op" &&
      (this.toks[this.i + 1] as { v: string }).v === "::"
    ) {
      axis = (this.toks[this.i] as { v: string }).v
      this.i += 2
    }
    const t = this.toks[this.i++]
    if (!t) throw new Error("xpath")
    let test: string
    if (t.t === "op" && t.v === "*") test = "*"
    else if (t.t === "name") {
      test = t.v
      if ((test === "text" || test === "node") && this.peek("(")) {
        this.i++
        this.take(")")
        test += "()"
      }
    } else throw new Error("xpath")
    const matches = (x: Item) =>
      test === "node()" ||
      (test === "text()" ? x.kind === "text" : x.kind !== "text" && (test === "*" || x.name === test))
    return (n) => {
      if (n.kind !== "el") return []
      let pool: Item[]
      if (axis === "attribute")
        pool = [...n.attrs].map(([name, value]) => ({ kind: "attr", name, value, parent: n }) as Attr)
      else if (axis === "child") pool = n.children
      else if (axis === "descendant") pool = descendants(n)
      else if (axis === "descendant-or-self") pool = [n, ...descendants(n)]
      else if (axis === "self") pool = [n]
      else if (axis === "parent") pool = n.parent ? [n.parent] : []
      else throw new Error("axis")
      return pool.filter(matches)
    }
  }
}

/** Nodes sorted as they come in the document (attributes right after their element). */
function documentOrder(nodes: Item[]): Item[] {
  if (nodes.length < 2) return nodes
  const root = rootOf(nodes[0])
  const order = new Map<XNode, number>([[root, 0]])
  descendants(root).forEach((n, i) => order.set(n, i + 1))
  const key = (n: Item) => (n.kind === "attr" ? (order.get(n.parent) ?? 0) + 0.5 : (order.get(n) ?? 0))
  return nodes.slice().sort((a, b) => key(a) - key(b))
}

function rootOf(n: Item): El {
  let r: El = n.kind === "el" ? n : n.parent
  while (r.parent) r = r.parent
  return r
}

function descendants(n: Item): XNode[] {
  if (n.kind !== "el") return []
  const out: XNode[] = []
  for (const c of n.children) {
    out.push(c)
    out.push(...descendants(c))
  }
  return out
}

function bool(v: XValue): boolean {
  if (Array.isArray(v)) return v.length > 0
  if (typeof v === "string") return v.length > 0
  if (typeof v === "number") return v !== 0 && !Number.isNaN(v)
  return v
}
function str(v: XValue): string {
  if (Array.isArray(v)) return v.length ? stringValue(v[0]) : ""
  return String(v)
}
function number(v: XValue): number {
  if (typeof v === "number") return v
  if (typeof v === "boolean") return v ? 1 : 0
  const s = str(v).trim()
  return s === "" ? NaN : Number(s)
}

function compareValues(a: XValue, b: XValue, op: string): boolean {
  const test = (x: string | number | boolean, y: string | number | boolean): boolean => {
    if (op === "=" || op === "!=") {
      const eq =
        typeof x === "boolean" || typeof y === "boolean"
          ? bool(x) === bool(y)
          : typeof x === "number" || typeof y === "number"
            ? number(x) === number(y)
            : x === y
      return op === "=" ? eq : !eq
    }
    const l = number(x)
    const r = number(y)
    return op === "<" ? l < r : op === ">" ? l > r : op === "<=" ? l <= r : l >= r
  }
  const values = (v: XValue) => (Array.isArray(v) ? v.map(stringValue) : [v])
  return values(a).some((x) => values(b).some((y) => test(x, y)))
}

function fn(name: string, args: ((n: Item, p: number, s: number) => XValue)[]) {
  const arg = (k: number, n: Item, p: number, s: number) => {
    const a = args[k]
    if (!a) throw new Error("argument")
    return a(n, p, s)
  }
  switch (name) {
    case "last":
      return (_n: Item, _p: number, s: number) => s
    case "position":
      return (_n: Item, p: number) => p
    case "not":
      return (n: Item, p: number, s: number) => !bool(arg(0, n, p, s))
    case "true":
      return () => true
    case "false":
      return () => false
    case "contains":
      return (n: Item, p: number, s: number) => str(arg(0, n, p, s)).includes(str(arg(1, n, p, s)))
    case "starts-with":
      return (n: Item, p: number, s: number) => str(arg(0, n, p, s)).startsWith(str(arg(1, n, p, s)))
    case "string":
      return (n: Item, p: number, s: number) => (args.length ? str(arg(0, n, p, s)) : stringValue(n))
    case "number":
      return (n: Item, p: number, s: number) => number(args.length ? arg(0, n, p, s) : stringValue(n))
    case "count":
      return (n: Item, p: number, s: number) => {
        const v = arg(0, n, p, s)
        if (!Array.isArray(v)) throw new Error("count")
        return v.length
      }
    case "name":
    case "local-name":
      return (n: Item) => (n.kind === "text" ? "" : n.name)
    case "normalize-space":
      return (n: Item, p: number, s: number) =>
        (args.length ? str(arg(0, n, p, s)) : stringValue(n)).trim().replace(/\s+/g, " ")
    case "string-length":
      return (n: Item, p: number, s: number) => (args.length ? str(arg(0, n, p, s)) : stringValue(n)).length
  }
  throw new Error("function")
}

const NUMBER = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/

scalarFn("FILTERXML", 2, 2, (a) => {
  const xml = toText(a[0] as Scalar)
  if (isError(xml)) return xml
  const xpath = toText(a[1] as Scalar)
  if (isError(xpath)) return xpath
  let found: XValue
  try {
    const doc = parseXml(xml)
    const p = new XPath(tokenize(xpath))
    const e = p.expr()
    if (!p.done) return VALUE
    found = e(doc, 1, 1)
  } catch {
    return VALUE
  }
  if (!Array.isArray(found) || found.length === 0) return VALUE
  const values: Scalar[] = found.map((n) => {
    const s = stringValue(n)
    return NUMBER.test(s) ? Number(s) : s
  })
  return values.length === 1 ? values[0] : matrix(values.length, 1, values)
})
