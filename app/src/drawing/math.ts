import { mathItalic, type MathList, type MathNode } from "@cells/core"
import type { Measure } from "./render"

/**
 * Equations drawn as Office's Professional format: each part a box (width, ascent and descent from its baseline)
 * holding SVG drawn at its own origin, put together as TeX and Office place them (fractions on the math axis,
 * scripts at 70%, stretching brackets and radicals drawn as paths).
 */

export type MathBox = { w: number; a: number; d: number; svg: string }

export const MATH_FONT = `'Cambria Math', 'STIX Two Math', 'Latin Modern Math', 'DejaVu Math TeX Gyre', 'Noto Sans Math', serif`

type Ctx = { size: number; color: string; measure: Measure; level: number; inFraction: boolean }

const n = (v: number) => Math.round(v * 100) / 100
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
const at = (x: number, y: number, b: MathBox) =>
  b.svg ? (x || y ? `<g transform="translate(${n(x)} ${n(y)})">${b.svg}</g>` : b.svg) : ""
const EMPTY = (ctx: Ctx): MathBox => ({ w: 0, a: ctx.size * 0.5, d: 0, svg: "" })

/** The math axis (where fraction bars and the minus sign sit) above the baseline. */
const axis = (ctx: Ctx) => ctx.size * 0.25
const scaled = (ctx: Ctx, by = 1): Ctx => {
  const level = Math.min(2, ctx.level + by)
  return { ...ctx, level, size: ctx.size * (level === ctx.level ? 1 : level === 1 ? 0.7 : 0.71) }
}

const BIN = new Set([..."+−-±∓×÷⋅∙∗∘∩∪∧∨⊕⊗⊖"])
const REL = new Set([..."=<>≤≥≠≈≡≅∼∝→←↔⇒⇐⇔∈∉∋⊂⊃⊆⊇≪≫⊥∥:≔"])
const PUNCT = new Set([...",;"])

function glyph(text: string, ctx: Ctx, font = MATH_FONT, size = ctx.size): MathBox {
  const css = `${n(size)}px ${font}`
  const w = ctx.measure(text, css)
  return {
    w,
    a: size * 0.72,
    d: size * 0.22,
    svg: `<text y="0" style="font:${css}" fill="${ctx.color}">${esc(text)}</text>`,
  }
}

function row(boxes: MathBox[], gaps: number[] = []): MathBox {
  let x = 0
  let a = 0
  let d = 0
  let svg = ""
  boxes.forEach((b, i) => {
    x += gaps[i] ?? 0
    svg += at(x, 0, b)
    x += b.w
    a = Math.max(a, b.a)
    d = Math.max(d, b.d)
  })
  return { w: x + (gaps[boxes.length] ?? 0), a, d, svg }
}

/** How an atom spaces: `fn` is a function or n-ary operator (a thin space before it, an operand after it). */
type Atom = { box: MathBox; kind: "ord" | "bin" | "rel" | "punct" | "open" | "close" | "fn" }

/** A list laid out with Office's spacing: medium around binary operators, thick around relations. */
export function layoutList(list: MathList, ctx: Ctx): MathBox {
  const atoms: Atom[] = []
  for (const node of list) {
    if (node.t === "r") atoms.push(...runAtoms(node, ctx))
    else atoms.push({ box: layoutNode(node, ctx), kind: node.t === "nary" || node.t === "func" ? "fn" : "ord" })
  }
  if (!atoms.length) return EMPTY(ctx)
  const em = ctx.size
  const gaps: number[] = [0]
  const script = ctx.level > 0
  for (let i = 0; i < atoms.length; i++) {
    const cur = atoms[i]!
    const prev = atoms[i - 1]
    // A binary operator with nothing on its left (−𝑏, (−1)) is a sign: no space.
    if (cur.kind === "bin" && (!prev || (prev.kind !== "ord" && prev.kind !== "close" && prev.kind !== "fn")))
      cur.kind = "ord"
    if (i > 0) {
      let g = 0
      if (!script && (cur.kind === "bin" || prev!.kind === "bin")) g = (em * 4) / 18
      else if (!script && (cur.kind === "rel" || prev!.kind === "rel")) g = (em * 5) / 18
      else if (prev!.kind === "punct") g = (em * 3) / 18
      else if (cur.kind === "fn" && (prev!.kind === "ord" || prev!.kind === "close" || prev!.kind === "fn"))
        g = (em * 3) / 18
      gaps.push(g)
    }
  }
  return row(
    atoms.map((x) => x.box),
    gaps,
  )
}

function runAtoms(node: Extract<MathNode, { t: "r" }>, ctx: Ctx): Atom[] {
  if (node.sty === "p") return [{ box: glyph(node.text, ctx), kind: "ord" }]
  const out: Atom[] = []
  let word = ""
  const flush = () => {
    if (!word) return
    const shown = node.sty ? word : [...word].map(mathItalic).join("")
    const box = glyph(shown, ctx, node.sty === "b" || node.sty === "bi" ? `700 ${MATH_FONT}` : MATH_FONT)
    // Spaces kept as text (after a comma) space nothing themselves: what follows them starts afresh.
    out.push({ box, kind: word.trim() ? "ord" : "open" })
    word = ""
  }
  for (const ch of node.text) {
    if (ch === "□") {
      // A place to type (Office's dotted box).
      flush()
      out.push({ box: placeholder(ctx), kind: "ord" })
      continue
    }
    const kind = BIN.has(ch) ? "bin" : REL.has(ch) ? "rel" : PUNCT.has(ch) ? "punct" : null
    if (!kind) {
      word += ch
      continue
    }
    flush()
    out.push({ box: glyph(ch === "-" ? "−" : ch, ctx), kind })
  }
  flush()
  return out
}

function placeholder(ctx: Ctx): MathBox {
  const s = ctx.size
  const w = s * 0.55
  const h = s * 0.68
  return {
    w: w + s * 0.1,
    a: h,
    d: s * 0.02,
    svg: `<rect x="${n(s * 0.05)}" y="${n(-h)}" width="${n(w)}" height="${n(h)}" fill="none" stroke="${ctx.color}" stroke-width="${n(Math.max(0.5, s * 0.04))}" stroke-dasharray="${n(s * 0.06)} ${n(s * 0.06)}" opacity="0.7"/>`,
  }
}

function layoutNode(node: MathNode, ctx: Ctx): MathBox {
  switch (node.t) {
    case "r":
      return layoutList([node], ctx)
    case "raw":
      return glyph(node.text, ctx)
    case "f":
      return fraction(node, ctx)
    case "sSup":
      return scripts(layoutList(node.e, ctx), null, node.sup, ctx)
    case "sSub":
      return scripts(layoutList(node.e, ctx), node.sub, null, ctx)
    case "sSubSup":
      return scripts(layoutList(node.e, ctx), node.sub, node.sup, ctx)
    case "sPre": {
      const s = scripts(EMPTY(ctx), node.sub, node.sup, ctx)
      return row([s, layoutList(node.e, ctx)])
    }
    case "rad":
      return radical(node, ctx)
    case "nary":
      return nary(node, ctx)
    case "d":
      return delimited(node, ctx)
    case "func": {
      const name = layoutList(node.name, ctx)
      return row([name, layoutList(node.e, ctx)], [0, (ctx.size * 3) / 18])
    }
    case "acc":
      return accent(node.chr, layoutList(node.e, ctx), ctx)
    case "bar": {
      const e = layoutList(node.e, ctx)
      const sw = ctx.size * 0.05
      const y = node.pos === "top" ? -e.a - sw * 2 : e.d + sw * 2
      return {
        w: e.w,
        a: node.pos === "top" ? e.a + sw * 3 : e.a,
        d: node.pos === "top" ? e.d : e.d + sw * 3,
        svg: e.svg + line(0, y, e.w, y, sw, ctx.color),
      }
    }
    case "limLow":
    case "limUpp":
      return limits(
        layoutList(node.e, ctx),
        node.t === "limLow" ? node.lim : null,
        node.t === "limUpp" ? node.lim : null,
        ctx,
      )
    case "groupChr":
      return groupChar(node, ctx)
    case "m":
      return matrix(node.rows, ctx)
    case "eqArr":
      return matrix(
        node.rows.map((r) => [r]),
        ctx,
      )
    case "box":
      return layoutList(node.e, ctx)
    case "borderBox": {
      const e = layoutList(node.e, ctx)
      const p = ctx.size * 0.15
      const sw = ctx.size * 0.05
      return {
        w: e.w + p * 2,
        a: e.a + p,
        d: e.d + p,
        svg:
          at(p, 0, e) +
          `<rect x="${n(sw / 2)}" y="${n(-e.a - p + sw / 2)}" width="${n(e.w + p * 2 - sw)}" height="${n(e.a + e.d + p * 2 - sw)}" fill="none" stroke="${ctx.color}" stroke-width="${n(sw)}"/>`,
      }
    }
  }
}

const line = (x1: number, y1: number, x2: number, y2: number, w: number, color: string) =>
  `<line x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" stroke="${color}" stroke-width="${n(w)}"/>`

function fraction(node: Extract<MathNode, { t: "f" }>, ctx: Ctx): MathBox {
  // In the equation itself parts stay full size; inside a fraction or a script they shrink.
  const inner = ctx.inFraction || ctx.level > 0 ? { ...scaled(ctx), inFraction: true } : { ...ctx, inFraction: true }
  const num = layoutList(node.num, inner)
  const den = layoutList(node.den, inner)
  if (node.type === "lin" || node.type === "skw") {
    const slash = glyph("/", ctx)
    return row([num, slash, den])
  }
  const pad = ctx.size * 0.12
  const sw = node.type === "noBar" ? 0 : Math.max(0.6, ctx.size * 0.055)
  const w = Math.max(num.w, den.w) + pad * 2
  const ax = axis(ctx)
  const gap = ctx.size * 0.14
  const numY = -ax - sw / 2 - gap - num.d
  const denY = -ax + sw / 2 + gap + den.a
  return {
    w,
    a: -numY + num.a,
    d: denY + den.d,
    svg:
      at((w - num.w) / 2, numY, num) +
      at((w - den.w) / 2, denY, den) +
      (sw ? line(pad * 0.5, -ax, w - pad * 0.5, -ax, sw, ctx.color) : ""),
  }
}

function scripts(base: MathBox, sub: MathList | null, sup: MathList | null, ctx: Ctx): MathBox {
  const small = scaled(ctx)
  const supBox = sup ? layoutList(sup, small) : null
  const subBox = sub ? layoutList(sub, small) : null
  const x = base.w + ctx.size * 0.04
  let svg = base.svg
  let a = base.a
  let d = base.d
  let w = base.w
  if (supBox) {
    const y = -Math.max(ctx.size * 0.42, base.a - small.size * 0.42)
    svg += at(x, y, supBox)
    a = Math.max(a, -y + supBox.a)
    w = Math.max(w, x + supBox.w)
  }
  if (subBox) {
    const y = Math.max(ctx.size * 0.2, base.d + small.size * 0.15)
    svg += at(x, y, subBox)
    d = Math.max(d, y + subBox.d)
    w = Math.max(w, x + subBox.w)
  }
  return { w: w + ctx.size * 0.04, a, d, svg }
}

function limits(base: MathBox, below: MathList | null, above: MathList | null, ctx: Ctx): MathBox {
  const small = scaled(ctx)
  const lo = below ? layoutList(below, small) : null
  const hi = above ? layoutList(above, small) : null
  const w = Math.max(base.w, lo?.w ?? 0, hi?.w ?? 0)
  const gap = ctx.size * 0.1
  let svg = at((w - base.w) / 2, 0, base)
  let a = base.a
  let d = base.d
  if (hi) {
    const y = -base.a - gap - hi.d
    svg += at((w - hi.w) / 2, y, hi)
    a = -y + hi.a
  }
  if (lo) {
    const y = base.d + gap + lo.a
    svg += at((w - lo.w) / 2, y, lo)
    d = y + lo.d
  }
  return { w, a, d, svg }
}

function nary(node: Extract<MathNode, { t: "nary" }>, ctx: Ctx): MathBox {
  const integral = /[∫∬∭∮∯∰]/.test(node.chr)
  const big = ctx.size * (ctx.level === 0 ? (integral ? 1.6 : 1.4) : 1.1)
  const sign = glyph(node.chr, ctx, MATH_FONT, big)
  const ax = axis(ctx)
  const ink = { a: big * 0.62, d: big * 0.28 }
  // The glyph placed so the middle of its ink is on the axis.
  const centred: MathBox = {
    w: sign.w,
    a: ax + (ink.a + ink.d) / 2,
    d: (ink.a + ink.d) / 2 - ax,
    svg: at(0, -ax + (ink.a - ink.d) / 2, sign),
  }
  const sub = node.sub.length ? node.sub : null
  const sup = node.sup.length ? node.sup : null
  const op =
    node.limLoc === "subSup" || ctx.level > 0
      ? scriptsBeside(centred, sub, sup, ctx, integral)
      : limits(centred, sub, sup, ctx)
  const body = layoutList(node.e, ctx)
  return row([op, body], [0, ctx.size * 0.12])
}

/** An integral's limits beside it, the lower one tucked under the slant. */
function scriptsBeside(
  sign: MathBox,
  sub: MathList | null,
  sup: MathList | null,
  ctx: Ctx,
  integral: boolean,
): MathBox {
  const small = scaled(ctx)
  const hi = sup ? layoutList(sup, small) : null
  const lo = sub ? layoutList(sub, small) : null
  let svg = sign.svg
  let w = sign.w
  let a = sign.a
  let d = sign.d
  if (hi) {
    const y = -sign.a + hi.a * 0.9
    svg += at(sign.w, y, hi)
    w = Math.max(w, sign.w + hi.w)
    a = Math.max(a, -y + hi.a)
  }
  if (lo) {
    const x = sign.w - (integral ? ctx.size * 0.35 : 0)
    const y = sign.d - lo.d * 0.2
    svg += at(x, y, lo)
    w = Math.max(w, x + lo.w)
    d = Math.max(d, y + lo.d)
  }
  return { w, a, d, svg }
}

function radical(node: Extract<MathNode, { t: "rad" }>, ctx: Ctx): MathBox {
  const e = layoutList(node.e, ctx)
  const s = ctx.size
  const sw = Math.max(0.6, s * 0.055)
  const gap = s * 0.12
  const top = -e.a - gap
  const bottom = e.d + s * 0.04
  const h = bottom - top
  const signW = s * 0.42 + Math.min(h, s * 3) * 0.06
  const deg = node.deg.length ? layoutList(node.deg, scaled(ctx, 2)) : null
  const dx = deg ? Math.max(0, deg.w - signW * 0.55) : 0
  const tickY = bottom - h * 0.42
  const path =
    `M${n(dx)} ${n(tickY + s * 0.04)} L${n(dx + s * 0.1)} ${n(tickY - s * 0.02)} ` +
    `L${n(dx + signW * 0.45)} ${n(bottom)} L${n(dx + signW)} ${n(top)} H${n(dx + signW + e.w + s * 0.1)}`
  let svg =
    `<path d="${path}" fill="none" stroke="${ctx.color}" stroke-width="${n(sw)}" stroke-linejoin="round" stroke-linecap="round"/>` +
    at(dx + signW + s * 0.05, 0, e)
  let a = -top + sw
  if (deg) {
    const y = tickY - s * 0.12 - deg.d
    svg += at(dx + signW * 0.5 - deg.w, y, deg)
    a = Math.max(a, -y + deg.a)
  }
  return { w: dx + signW + e.w + s * 0.16, a, d: bottom + sw, svg }
}

function accent(chr: string, e: MathBox, ctx: Ctx): MathBox {
  const s = ctx.size
  const sw = Math.max(0.6, s * 0.05)
  const y = -e.a - s * 0.08
  const mid = e.w / 2
  let mark: string
  if (chr === "⃗") {
    // A vector arrow.
    const half = Math.max(e.w / 2, s * 0.25)
    mark =
      line(mid - half, y - s * 0.08, mid + half, y - s * 0.08, sw, ctx.color) +
      `<path d="M${n(mid + half - s * 0.12)} ${n(y - s * 0.16)} L${n(mid + half)} ${n(y - s * 0.08)} L${n(mid + half - s * 0.12)} ${n(y)}" fill="none" stroke="${ctx.color}" stroke-width="${n(sw)}" stroke-linejoin="round"/>`
  } else if (chr === "̅" || chr === "̄" || chr === "¯") mark = line(0, y - s * 0.06, e.w, y - s * 0.06, sw, ctx.color)
  else {
    const shown: Record<string, string> = {
      "̂": "ˆ",
      "̃": "˜",
      "̇": "˙",
      "̈": "¨",
      "́": "´",
      "̀": "`",
      "̌": "ˇ",
      "̆": "˘",
    }
    const g = glyph(shown[chr] ?? chr, ctx)
    mark = at(mid - g.w / 2, y + s * 0.55, g)
  }
  return { w: e.w, a: e.a + s * 0.3, d: e.d, svg: e.svg + mark }
}

function groupChar(node: Extract<MathNode, { t: "groupChr" }>, ctx: Ctx): MathBox {
  const e = layoutList(node.e, ctx)
  const s = ctx.size
  const sw = Math.max(0.6, s * 0.05)
  const depth = s * 0.22
  const y0 = node.pos === "top" ? -e.a - s * 0.08 : e.d + s * 0.08
  const dir = node.pos === "top" ? -1 : 1
  const w = e.w
  const q = Math.min(s * 0.2, w / 4)
  const y1 = y0 + (dir * depth) / 2
  const y2 = y0 + dir * depth
  const brace =
    `M0 ${n(y0)} Q0 ${n(y1)} ${n(q)} ${n(y1)} H${n(w / 2 - q)} Q${n(w / 2)} ${n(y1)} ${n(w / 2)} ${n(y2)} ` +
    `Q${n(w / 2)} ${n(y1)} ${n(w / 2 + q)} ${n(y1)} H${n(w - q)} Q${n(w)} ${n(y1)} ${n(w)} ${n(y0)}`
  return {
    w,
    a: node.pos === "top" ? -y2 + sw : e.a,
    d: node.pos === "top" ? e.d : y2 + sw,
    svg:
      e.svg + `<path d="${brace}" fill="none" stroke="${ctx.color}" stroke-width="${n(sw)}" stroke-linecap="round"/>`,
  }
}

function matrix(rows: MathList[][], ctx: Ctx): MathBox {
  const cells = rows.map((r) => r.map((c) => layoutList(c, ctx)))
  const cols = Math.max(1, ...cells.map((r) => r.length))
  const colW = Array.from({ length: cols }, (_, c) => Math.max(0, ...cells.map((r) => r[c]?.w ?? 0)))
  const gapX = ctx.size * 0.8
  const gapY = ctx.size * 0.25
  const heights = cells.map((r) => ({
    a: Math.max(ctx.size * 0.7, ...r.map((b) => b.a)),
    d: Math.max(ctx.size * 0.2, ...r.map((b) => b.d)),
  }))
  const total = heights.reduce((s, h) => s + h.a + h.d, 0) + gapY * (rows.length - 1)
  // The block centred on the math axis.
  let y = -axis(ctx) - total / 2
  let svg = ""
  cells.forEach((r, i) => {
    const base = y + heights[i]!.a
    let x = 0
    r.forEach((b, c) => {
      svg += at(x + (colW[c]! - b.w) / 2, base, b)
      x += colW[c]! + gapX
    })
    y = base + heights[i]!.d + gapY
  })
  const w = colW.reduce((s, v) => s + v, 0) + gapX * (cols - 1)
  return { w, a: axis(ctx) + total / 2, d: total / 2 - axis(ctx), svg }
}

// ── Brackets ─────────────────────────────────────────────────────────────────────────────────────────────

function delimited(node: Extract<MathNode, { t: "d" }>, ctx: Ctx): MathBox {
  const parts = node.items.map((i) => layoutList(i, ctx))
  const sep = node.sep ?? "|"
  const inner: MathBox[] = []
  const gaps: number[] = [0]
  parts.forEach((p, i) => {
    if (i) {
      inner.push(fence(sep, p.a, p.d, ctx))
      gaps.push(ctx.size * 0.1, ctx.size * 0.1)
    }
    inner.push(p)
  })
  const body = row(inner, gaps)
  const open = fence(node.beg, body.a, body.d, ctx)
  const close = fence(node.end, body.a, body.d, ctx)
  const pad = ctx.size * 0.04
  return row([open, body, close], [0, pad, pad, 0])
}

/** A bracket tall enough for content of ascent `a` and descent `d`, symmetric about the math axis. */
function fence(chr: string, a: number, d: number, ctx: Ctx): MathBox {
  if (!chr) return { w: 0, a: 0, d: 0, svg: "" }
  const s = ctx.size
  const ax = axis(ctx)
  const half = Math.max(a - ax, d + ax) + s * 0.08
  if (half * 2 <= s * 1.15) return glyph(chr, ctx)
  const top = -ax - half
  const bottom = -ax + half
  const mid = -ax
  const sw = Math.max(0.6, s * 0.055)
  const w = s * 0.32 + Math.min(half, s * 3) * 0.06
  const stroke = `fill="none" stroke="${ctx.color}" stroke-width="${n(sw)}" stroke-linecap="round" stroke-linejoin="round"`
  const flip = (path: string) => `<g transform="translate(${n(w)} 0) scale(-1 1)">${path}</g>`
  let svg: string
  switch (chr) {
    case "(":
    case ")": {
      const x1 = w - s * 0.06
      const x0 = s * 0.04
      const t = sw * 1.8
      const p = `<path d="M${n(x1)} ${n(top)} Q${n(x0)} ${n(mid)} ${n(x1)} ${n(bottom)} Q${n(x0 + t)} ${n(mid)} ${n(x1)} ${n(top)} Z" fill="${ctx.color}" stroke="${ctx.color}" stroke-width="${n(sw * 0.4)}"/>`
      svg = chr === "(" ? p : flip(p)
      break
    }
    case "[":
    case "]": {
      const p = `<path d="M${n(w - s * 0.06)} ${n(top)} H${n(s * 0.1)} V${n(bottom)} H${n(w - s * 0.06)}" ${stroke}/>`
      svg = chr === "[" ? p : flip(p)
      break
    }
    case "⌊":
    case "⌋":
    case "⌈":
    case "⌉": {
      const floor = chr === "⌊" || chr === "⌋"
      const p = floor
        ? `<path d="M${n(s * 0.1)} ${n(top)} V${n(bottom)} H${n(w - s * 0.06)}" ${stroke}/>`
        : `<path d="M${n(w - s * 0.06)} ${n(top)} H${n(s * 0.1)} V${n(bottom)}" ${stroke}/>`
      svg = chr === "⌊" || chr === "⌈" ? p : flip(p)
      break
    }
    case "{":
    case "}": {
      const r = w - s * 0.06
      const m = w / 2
      const l = s * 0.04
      const q = Math.min(s * 0.25, half / 3)
      const p =
        `<path d="M${n(r)} ${n(top)} Q${n(m)} ${n(top)} ${n(m)} ${n(top + q)} V${n(mid - q)} Q${n(m)} ${n(mid)} ${n(l)} ${n(mid)} ` +
        `Q${n(m)} ${n(mid)} ${n(m)} ${n(mid + q)} V${n(bottom - q)} Q${n(m)} ${n(bottom)} ${n(r)} ${n(bottom)}" ${stroke}/>`
      svg = chr === "{" ? p : flip(p)
      break
    }
    case "⟨":
    case "⟩": {
      const p = `<path d="M${n(w - s * 0.06)} ${n(top)} L${n(s * 0.08)} ${n(mid)} L${n(w - s * 0.06)} ${n(bottom)}" ${stroke}/>`
      svg = chr === "⟨" ? p : flip(p)
      break
    }
    case "|":
    case "│":
    case "∣":
      return { w: s * 0.3, a: -top, d: bottom, svg: line(s * 0.15, top, s * 0.15, bottom, sw, ctx.color) }
    case "‖":
      return {
        w: s * 0.42,
        a: -top,
        d: bottom,
        svg:
          line(s * 0.13, top, s * 0.13, bottom, sw, ctx.color) + line(s * 0.29, top, s * 0.29, bottom, sw, ctx.color),
      }
    default: {
      // Any other character: the glyph stretched to the height.
      const g = glyph(chr, ctx)
      const k = (half * 2) / (g.a + g.d)
      return {
        w: g.w,
        a: -top,
        d: bottom,
        svg: `<g transform="translate(0 ${n(mid)}) scale(1 ${n(k)}) translate(0 ${n((g.a - g.d) / 2)})">${g.svg}</g>`,
      }
    }
  }
  return { w, a: -top, d: bottom, svg }
}

/** An equation's lines as boxes at `size` pixels in `color`. */
export function layoutEquation(lines: MathList[], size: number, color: string, measure: Measure): MathBox[] {
  return lines.map((l) => layoutList(l, { size, color, measure, level: 0, inFraction: false }))
}
