import { FUNCTION_NAMES, NARY, type MathList, type MathNode } from "./model"

/**
 * Office's linear format for equations (UnicodeMath, Unicode Technical Note 28), the subset Office's own equations
 * use: `x=(−b±√(b^2−4ac))/2a`, `∑_(k=0)^n▒〖x^k〗`, `sin⁡〖nπx/L〗`, `■(a&b@c&d)`. Typing in an equation is in this
 * form; it builds up into the tree drawn professionally.
 */

const FUNC_APPLY = "⁡"
const OF = "▒"
const GROUP_OPEN = "〖"
const GROUP_CLOSE = "〗"
/** Brackets and their closers; ├ and ┤ are Office's invisible ones ({■(…)┤ is a brace on the left only). */
const OPENERS: Record<string, string> = {
  "(": ")",
  "[": "]",
  "{": "}",
  "⟨": "⟩",
  "⌊": "⌋",
  "⌈": "⌉",
  "|": "|",
  "‖": "‖",
  "├": "┤",
}
const CLOSERS = new Set([")", "]", "}", "⟩", "⌋", "⌉", "┤", GROUP_CLOSE])
const FRACTIONS: Record<string, "noBar" | "lin" | undefined> = { "/": undefined, "¦": "noBar", "∕": "lin" }
/** Characters that belong to the operand before them (n!, f′). */
const POSTFIX = new Set(["!", "′", "″", "‴", "%"])
/** Ordinary symbols that read as operands. */
const ORDINARY = new Set(["∞", "∂", "∇", "…", "⋯", "ℏ", "∅", "□"])
/** Combining marks Office uses for accents: x̂ x̃ ẋ ẍ x⃗ x̅. */
const ACCENTS = new Set(["̂", "̃", "̇", "̈", "⃗", "̅", "́", "̀", "̌", "̆"])

/** Math AutoCorrect: \alpha → α, \sum → ∑ … (Office's list, the common part). */
export const MATH_AUTOCORRECT: Record<string, string> = {
  alpha: "α",
  beta: "β",
  gamma: "γ",
  delta: "δ",
  epsilon: "ϵ",
  varepsilon: "ε",
  zeta: "ζ",
  eta: "η",
  theta: "θ",
  vartheta: "ϑ",
  iota: "ι",
  kappa: "κ",
  lambda: "λ",
  mu: "μ",
  nu: "ν",
  xi: "ξ",
  pi: "π",
  varpi: "ϖ",
  rho: "ρ",
  sigma: "σ",
  varsigma: "ς",
  tau: "τ",
  upsilon: "υ",
  phi: "ϕ",
  varphi: "φ",
  chi: "χ",
  psi: "ψ",
  omega: "ω",
  Gamma: "Γ",
  Delta: "Δ",
  Theta: "Θ",
  Lambda: "Λ",
  Xi: "Ξ",
  Pi: "Π",
  Sigma: "Σ",
  Upsilon: "Υ",
  Phi: "Φ",
  Psi: "Ψ",
  Omega: "Ω",
  sum: "∑",
  prod: "∏",
  coprod: "∐",
  int: "∫",
  iint: "∬",
  iiint: "∭",
  oint: "∮",
  bigcup: "⋃",
  bigcap: "⋂",
  sqrt: "√",
  cbrt: "∛",
  qdrt: "∜",
  infty: "∞",
  partial: "∂",
  nabla: "∇",
  pm: "±",
  mp: "∓",
  times: "×",
  div: "÷",
  cdot: "⋅",
  bullet: "∙",
  le: "≤",
  leq: "≤",
  ge: "≥",
  geq: "≥",
  ne: "≠",
  neq: "≠",
  approx: "≈",
  equiv: "≡",
  cong: "≅",
  sim: "∼",
  propto: "∝",
  to: "→",
  rightarrow: "→",
  leftarrow: "←",
  Rightarrow: "⇒",
  Leftarrow: "⇐",
  leftrightarrow: "↔",
  Leftrightarrow: "⇔",
  in: "∈",
  notin: "∉",
  ni: "∋",
  subset: "⊂",
  supset: "⊃",
  subseteq: "⊆",
  supseteq: "⊇",
  cup: "∪",
  cap: "∩",
  emptyset: "∅",
  forall: "∀",
  exists: "∃",
  neg: "¬",
  wedge: "∧",
  vee: "∨",
  angle: "∠",
  perp: "⊥",
  parallel: "∥",
  degree: "°",
  hbar: "ℏ",
  ell: "ℓ",
  Re: "ℜ",
  Im: "ℑ",
  aleph: "ℵ",
  ldots: "…",
  cdots: "⋯",
  vdots: "⋮",
  ddots: "⋱",
  of: OF,
  funcapply: FUNC_APPLY,
  matrix: "■",
  eqarray: "█",
  below: "┬",
  above: "┴",
  box: "□",
  rect: "▭",
  hat: "̂",
  tilde: "̃",
  dot: "̇",
  ddot: "̈",
  vec: "⃗",
  bar: "̅",
  atop: "¦",
  langle: "⟨",
  rangle: "⟩",
  lfloor: "⌊",
  rfloor: "⌋",
  lceil: "⌈",
  rceil: "⌉",
  norm: "‖",
}

// ── Math italic letters (𝑥) and plain ones ──────────────────────────────────────────────────────────────

const ITALIC_A = 0x1d434
const ITALIC_a = 0x1d44e
const ITALIC_alpha = 0x1d6fc

/** A letter as Office draws a variable: mathematical italic (𝑥, 𝐴, 𝛼); other characters as they are. */
export function mathItalic(ch: string): string {
  const c = ch.codePointAt(0)!
  if (ch === "h") return "ℎ"
  if (c >= 0x41 && c <= 0x5a) return String.fromCodePoint(ITALIC_A + c - 0x41)
  if (c >= 0x61 && c <= 0x7a) return String.fromCodePoint(ITALIC_a + c - 0x61)
  if (c >= 0x3b1 && c <= 0x3c9) return String.fromCodePoint(ITALIC_alpha + c - 0x3b1)
  return ch
}

/** A math italic letter back to its plain letter (the tree keeps plain letters). */
export function plainLetter(ch: string): string {
  const c = ch.codePointAt(0)!
  if (ch === "ℎ") return "h"
  if (c >= ITALIC_A && c < ITALIC_A + 26) return String.fromCharCode(0x41 + c - ITALIC_A)
  if (c >= ITALIC_a && c < ITALIC_a + 26) return String.fromCharCode(0x61 + c - ITALIC_a)
  if (c >= ITALIC_alpha && c < ITALIC_alpha + 25) return String.fromCharCode(0x3b1 + c - ITALIC_alpha)
  return ch
}

const isLetter = (ch: string) => /\p{L}/u.test(ch)
const isDigit = (ch: string) => /[0-9]/.test(ch)

// ── Parsing ──────────────────────────────────────────────────────────────────────────────────────────────

/** Replaces Math AutoCorrect names (\alpha) with their characters. */
export function autocorrect(text: string): string {
  return text.replace(/\\([A-Za-z]+)/g, (all, name: string) => MATH_AUTOCORRECT[name] ?? all)
}

class Parser {
  private chars: string[]
  private i = 0
  constructor(text: string) {
    this.chars = [...autocorrect(text)].map(plainLetter)
  }
  private peek(k = 0) {
    return this.chars[this.i + k]
  }
  private next() {
    return this.chars[this.i++]
  }
  done() {
    return this.i >= this.chars.length
  }

  /** A sequence up to a closing bracket or one of `stops` (not consumed). */
  expression(stops: string[] = []): MathList {
    const out: MathList[] = []
    while (!this.done()) {
      const ch = this.peek()!
      if (CLOSERS.has(ch) || stops.includes(ch)) break
      if (ch === " ") {
        this.i++
        // After punctuation a space is text (Office keeps "…,  −∞<𝑥<∞" as typed).
        const last = out[out.length - 1]?.[0]
        if (last?.t === "r" && /[,;:] *$/.test(last.text)) last.text += " "
        continue
      }
      if (ch in FRACTIONS) {
        this.i++
        const num = out.pop() ?? []
        const den = this.operand(stops)
        const type = FRACTIONS[ch]
        out.push([{ t: "f", num: unwrap(num), den: unwrap(den), ...(type ? { type } : {}) }])
        continue
      }
      if (startsOperand(ch, this.peek(1))) {
        out.push(this.operand(stops))
        continue
      }
      // An operator, relation or punctuation: a run of its own.
      this.i++
      out.push([{ t: "r", text: ch }])
    }
    return merge(out.flat())
  }

  /** An operand: atoms with nothing between them (`2a`, `n!`, `x^2`, `(a+b)`), up to an operator or space. */
  operand(stops: string[]): MathList {
    const atoms: MathList = []
    while (!this.done()) {
      const ch = this.peek()!
      if (ch === " " || CLOSERS.has(ch) || stops.includes(ch) || ch in FRACTIONS) break
      if (!startsOperand(ch, this.peek(1)) && !POSTFIX.has(ch)) break
      if (POSTFIX.has(ch)) {
        this.i++
        atoms.push({ t: "r", text: ch })
        continue
      }
      atoms.push(...this.scripted(stops))
    }
    return merge(atoms)
  }

  /** An atom with what follows it: scripts, accents, a limit below or above. */
  private scripted(stops: string[]): MathList {
    let base = this.atom(stops)
    for (;;) {
      const ch = this.peek()
      if (ch && ACCENTS.has(ch)) {
        this.i++
        base = [ch === "̅" ? { t: "bar", pos: "top", e: base } : { t: "acc", chr: ch, e: base }]
        continue
      }
      if (ch === "┬" || ch === "┴") {
        this.i++
        const lim = unwrap(this.scriptArg(stops))
        base = [ch === "┬" ? { t: "limLow", e: base, lim } : { t: "limUpp", e: base, lim }]
        continue
      }
      break
    }
    let sub: MathList | null = null
    let sup: MathList | null = null
    while (this.peek() === "_" || this.peek() === "^") {
      const which = this.next()
      const arg = unwrap(this.scriptArg(stops))
      if (which === "_") sub = arg
      else sup = arg
    }
    // A function name's argument comes after its scripts (sin^2⁡x).
    if (sub || sup) {
      const e = base
      base = [sub && sup ? { t: "sSubSup", e, sub, sup } : sub ? { t: "sSub", e, sub } : { t: "sSup", e, sup: sup! }]
    }
    if (this.peek() === FUNC_APPLY) {
      this.i++
      while (this.peek() === " ") this.i++
      const arg = this.peek() === GROUP_OPEN ? this.group() : this.scripted(stops)
      return [{ t: "func", name: base, e: arg }]
    }
    return base
  }

  /** A script's argument: one bracketed group, a number, or one character. */
  private scriptArg(stops: string[]): MathList {
    const ch = this.peek()
    if (!ch) return []
    if (ch === "(" || ch === GROUP_OPEN) {
      this.i++
      const inner = this.expression(stops)
      if (this.peek() === OPENERS[ch] || this.peek() === GROUP_CLOSE) this.i++
      return inner
    }
    if (isDigit(ch)) {
      let text = ""
      while (
        this.peek() !== undefined &&
        (isDigit(this.peek()!) || (this.peek() === "." && isDigit(this.peek(1) ?? "")))
      )
        text += this.next()
      return [{ t: "r", text }]
    }
    return this.atom(stops)
  }

  /** `〖…〗`: a group that only holds its content together. */
  private group(): MathList {
    this.i++
    const inner = this.expression()
    if (this.peek() === GROUP_CLOSE) this.i++
    return inner
  }

  private atom(stops: string[]): MathList {
    const ch = this.peek()!
    if (ch === GROUP_OPEN) return this.group()
    if (ch === "■" || ch === "█") return this.array(ch)
    if (ch === "□" && this.peek(1) === "(") {
      this.i += 2
      const e = this.expression()
      if (this.peek() === ")") this.i++
      return [{ t: "box", e }]
    }
    if (ch === "▭" && this.peek(1) === "(") {
      this.i += 2
      const e = this.expression()
      if (this.peek() === ")") this.i++
      return [{ t: "borderBox", e }]
    }
    if (ch === "√" || ch === "∛" || ch === "∜") {
      this.i++
      const deg: MathList = ch === "∛" ? [{ t: "r", text: "3" }] : ch === "∜" ? [{ t: "r", text: "4" }] : []
      if (this.peek() === "(") {
        this.i++
        const first = this.expression(["&"])
        if (this.peek() === "&") {
          this.i++
          const e = this.expression()
          if (this.peek() === ")") this.i++
          return [{ t: "rad", deg: first, e }]
        }
        if (this.peek() === ")") this.i++
        return [{ t: "rad", deg, e: first }]
      }
      return [{ t: "rad", deg, e: this.scriptArg(stops) }]
    }
    if (NARY[ch]) {
      this.i++
      let sub: MathList = []
      let sup: MathList = []
      while (this.peek() === "_" || this.peek() === "^") {
        const which = this.next()
        const arg = unwrap(this.scriptArg(stops))
        if (which === "_") sub = arg
        else sup = arg
      }
      let e: MathList = []
      while (this.peek() === " ") this.i++
      if (this.peek() === OF) {
        this.i++
        e = this.peek() === GROUP_OPEN ? this.group() : this.operand(stops)
      }
      return [{ t: "nary", chr: ch, sub, sup, e, limLoc: NARY[ch] }]
    }
    if (OPENERS[ch]) {
      this.i++
      const close = OPENERS[ch]!
      const items: MathList[] = [this.expression(["│", "∣", ...(ch === "|" || ch === "‖" ? [ch] : [])])]
      let sep: string | undefined
      while (this.peek() === "│" || this.peek() === "∣") {
        sep = this.next()
        items.push(this.expression(["│", "∣"]))
      }
      // Any closer ends the bracket (Office pairs ( with ], or a brace with the invisible ┤).
      let end = close
      const got = this.peek()
      if (got !== undefined && (got === close || CLOSERS.has(got)) && got !== GROUP_CLOSE) {
        this.i++
        end = got
      }
      return [{ t: "d", beg: ch === "├" ? "" : ch, end: end === "┤" ? "" : end, ...(sep ? { sep } : {}), items }]
    }
    if (isLetter(ch)) {
      // A function name before ⁡ (or a known name followed by a bracket) stays one upright word.
      let j = this.i
      while (j < this.chars.length && /[A-Za-z]/.test(this.chars[j]!)) j++
      const word = this.chars.slice(this.i, j).join("")
      const isName = FUNCTION_NAMES.includes(word) || this.chars[j] === FUNC_APPLY
      if (word.length > 1 && isName) {
        this.i = j
        return [{ t: "r", text: word, sty: "p" }]
      }
      this.i++
      return [{ t: "r", text: ch }]
    }
    if (isDigit(ch)) {
      let text = ""
      while (
        this.peek() !== undefined &&
        (isDigit(this.peek()!) || (this.peek() === "." && isDigit(this.peek(1) ?? "")))
      )
        text += this.next()
      return [{ t: "r", text }]
    }
    this.i++
    return [{ t: "r", text: ch }]
  }

  /** `■(a&b@c&d)` (a matrix) and `█(a@b)` (lines of an equation array). */
  private array(kind: string): MathList {
    this.i++
    if (this.peek() !== "(") return [{ t: "r", text: kind }]
    this.i++
    const rows: MathList[][] = [[]]
    for (;;) {
      rows[rows.length - 1]!.push(this.expression(["&", "@"]))
      const ch = this.peek()
      if (ch === "&") this.i++
      else if (ch === "@") {
        this.i++
        rows.push([])
      } else break
    }
    if (this.peek() === ")") this.i++
    if (kind === "█")
      return [
        { t: "eqArr", rows: rows.map((r) => r.flatMap((c, i) => (i ? [{ t: "r", text: "&" } as MathNode, ...c] : c))) },
      ]
    return [{ t: "m", rows }]
  }
}

function startsOperand(ch: string, after: string | undefined): boolean {
  if (isLetter(ch) || isDigit(ch) || ORDINARY.has(ch)) return true
  if (ch === "." && after !== undefined && isDigit(after)) return true
  if (ch === GROUP_OPEN || ch === "√" || ch === "∛" || ch === "∜" || ch === "■" || ch === "█") return true
  if ((ch === "□" || ch === "▭") && after === "(") return true
  if (NARY[ch]) return true
  return OPENERS[ch] !== undefined
}

/** A fraction's or script's part without the brackets that only held it together: (a+b)/c → a+b over c. */
function unwrap(list: MathList): MathList {
  if (list.length === 1) {
    const only = list[0]!
    if (only.t === "d" && only.beg === "(" && only.end === ")" && only.items.length === 1) return only.items[0]!
  }
  return list
}

/** Neighbouring plain runs as one (Office writes `2a` as one run). */
function merge(list: MathList): MathList {
  const out: MathList = []
  for (const n of list) {
    const last = out[out.length - 1]
    if (
      n.t === "r" &&
      last?.t === "r" &&
      (last.sty ?? "") === (n.sty ?? "") &&
      !n.sty &&
      !isOperator(last.text.slice(-1)) &&
      !isOperator(n.text[0]!)
    )
      last.text += n.text
    else out.push(n.t === "r" ? { ...n } : n)
  }
  return out
}

const isOperator = (ch: string) => !(isLetter(ch) || isDigit(ch) || ch === "." || POSTFIX.has(ch) || ORDINARY.has(ch))

/** An equation typed in the linear format, built up. */
export function parseLinear(text: string): MathList {
  const p = new Parser(text)
  const out: MathList = []
  while (!p.done()) {
    out.push(...p.expression())
    // A stray closing bracket: kept as text.
    if (!p.done()) out.push({ t: "r", text: (p as unknown as { next(): string }).next() })
  }
  return merge(out)
}

// ── Writing ──────────────────────────────────────────────────────────────────────────────────────────────

/** Whether a part reads as one operand without brackets (`2`, `𝑥`, `10`, `𝑛!`). */
function simple(list: MathList): boolean {
  if (!list.length) return false
  return list.every((n) =>
    n.t === "r"
      ? !n.sty && [...n.text].every((c) => !isOperator(c))
      : n.t === "sSup" || n.t === "sSub" || n.t === "sSubSup" || n.t === "d" || n.t === "rad" || n.t === "acc",
  )
}

const arg = (list: MathList) => (simple(list) ? toLinear(list) : `(${toLinear(list)})`)
const scriptArg = (list: MathList) => {
  const one = list.length === 1 && list[0]!.t === "r" && [...(list[0] as { text: string }).text].length === 1
  const number = list.length === 1 && list[0]!.t === "r" && /^[0-9.]+$/.test((list[0] as { text: string }).text)
  return one || number ? toLinear(list) : `(${toLinear(list)})`
}
const body = (list: MathList) =>
  list.length <= 1 && simple(list) ? toLinear(list) : `${GROUP_OPEN}${toLinear(list)}${GROUP_CLOSE}`

/** The tree in the linear format, letters in math italic as Office shows it. */
export function toLinear(list: MathList): string {
  let out = ""
  list.forEach((n, i) => {
    const prev = list[i - 1]
    // A space ends a fraction or script before what follows (1/2 (𝛼+𝛽), 𝑎_𝑛 cos⁡𝑥), and sets off a function name.
    if (prev && startsWithOperand(n) && (ENDS_AT_SPACE.has(prev.t) || (n.t === "func" && !isOperatorNode(prev))))
      out += " "
    out += nodeLinear(n)
  })
  return out
}

const ENDS_AT_SPACE = new Set(["f", "sSup", "sSub", "sSubSup", "limLow", "limUpp", "func", "nary"])
const isOperatorNode = (n: MathNode) => n.t === "r" && !n.sty && [...n.text].every(isOperator)
const startsWithOperand = (n: MathNode) => !(n.t === "r" && !n.sty && isOperator([...n.text][0] ?? " "))

function nodeLinear(n: MathNode): string {
  switch (n.t) {
    case "r":
      return n.sty === "p" ? n.text : [...n.text].map(mathItalic).join("")
    case "raw":
      return n.text
    case "f":
      return n.type === "noBar"
        ? `${toLinear(n.num)}¦${toLinear(n.den)}`
        : `${arg(n.num)}${n.type === "lin" || n.type === "skw" ? "∕" : "/"}${arg(n.den)}`
    case "sSup":
      return `${toLinear(n.e)}^${scriptArg(n.sup)}`
    case "sSub":
      return `${toLinear(n.e)}_${scriptArg(n.sub)}`
    case "sSubSup":
      return `${toLinear(n.e)}_${scriptArg(n.sub)}^${scriptArg(n.sup)}`
    case "sPre":
      return `〖_${scriptArg(n.sub)}^${scriptArg(n.sup)}〗${toLinear(n.e)}`
    case "rad":
      return n.deg.length ? `√(${toLinear(n.deg)}&${toLinear(n.e)})` : `√${arg(n.e)}`
    case "nary": {
      const sub = n.sub.length ? `_${scriptArg(n.sub)}` : ""
      const sup = n.sup.length ? `^${scriptArg(n.sup)}` : ""
      return `${n.chr}${sub}${sup}${OF}${body(n.e)}`
    }
    case "d":
      return `${n.beg || "├"}${n.items.map(toLinear).join(n.sep ?? "│")}${n.end || "┤"}`
    case "func":
      return `${toLinear(n.name)}${FUNC_APPLY}${body(n.e)}`
    case "acc":
      return `${body(n.e)}${n.chr}`
    case "bar":
      return `${body(n.e)}̅`
    case "limLow":
      return `${toLinear(n.e)}┬${scriptArg(n.lim)}`
    case "limUpp":
      return `${toLinear(n.e)}┴${scriptArg(n.lim)}`
    case "groupChr":
      return `${n.chr}${body(n.e)}`
    case "m":
      return `■(${n.rows.map((r) => r.map(toLinear).join("&")).join("@")})`
    case "eqArr":
      return `█(${n.rows.map(toLinear).join("@")})`
    case "box":
      return `□(${toLinear(n.e)})`
    case "borderBox":
      return `▭(${toLinear(n.e)})`
  }
}
