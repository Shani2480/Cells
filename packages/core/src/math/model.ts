/**
 * Insert › Equation: an equation as Office Math keeps it (OMML, ECMA-376 Part 1 §22.1), as a tree. Each kind matches
 * an OMML element; `raw` keeps one Cells doesn't model, written back as it came.
 */

export type MathList = MathNode[]

export type MathNode =
  /** Text: variables, numbers, operators. `sty` "p" is upright (function names, Normal Text); letters are italic. */
  | { t: "r"; text: string; sty?: "p" | "b" | "i" | "bi" }
  | { t: "f"; num: MathList; den: MathList; type?: "bar" | "skw" | "lin" | "noBar" }
  | { t: "sSup"; e: MathList; sup: MathList }
  | { t: "sSub"; e: MathList; sub: MathList }
  | { t: "sSubSup"; e: MathList; sub: MathList; sup: MathList }
  | { t: "sPre"; sub: MathList; sup: MathList; e: MathList }
  | { t: "rad"; deg: MathList; e: MathList }
  /** ∑ ∫ ∏ and the like; limits under and over (`undOvr`) or as scripts (`subSup`). */
  | { t: "nary"; chr: string; sub: MathList; sup: MathList; e: MathList; limLoc?: "undOvr" | "subSup" }
  /** Brackets around items separated by `sep`. */
  | { t: "d"; beg: string; end: string; sep?: string; items: MathList[] }
  | { t: "func"; name: MathList; e: MathList }
  | { t: "acc"; chr: string; e: MathList }
  | { t: "bar"; pos: "top" | "bot"; e: MathList }
  | { t: "limLow"; e: MathList; lim: MathList }
  | { t: "limUpp"; e: MathList; lim: MathList }
  | { t: "groupChr"; chr: string; pos: "top" | "bot"; e: MathList }
  | { t: "m"; rows: MathList[][] }
  | { t: "eqArr"; rows: MathList[] }
  | { t: "box"; e: MathList }
  | { t: "borderBox"; e: MathList }
  | { t: "raw"; xml: string; text: string }

/** An equation: one or more lines of math (Office's math paragraph), drawn centred. */
export type Equation = { lines: MathList[] }

/** Names that read as functions (upright, followed by their argument). */
export const FUNCTION_NAMES = [
  "sin",
  "cos",
  "tan",
  "cot",
  "sec",
  "csc",
  "sinh",
  "cosh",
  "tanh",
  "coth",
  "sech",
  "csch",
  "arcsin",
  "arccos",
  "arctan",
  "log",
  "ln",
  "lg",
  "exp",
  "lim",
  "max",
  "min",
  "det",
  "dim",
  "gcd",
  "arg",
  "deg",
  "sup",
  "inf",
  "Pr",
]

/** N-ary operators: their limits go under and over (∑) or beside (∫), as Office places them by default. */
export const NARY: Record<string, "undOvr" | "subSup"> = {
  "∑": "undOvr",
  "∏": "undOvr",
  "∐": "undOvr",
  "⋃": "undOvr",
  "⋂": "undOvr",
  "⋁": "undOvr",
  "⋀": "undOvr",
  "⨁": "undOvr",
  "⨂": "undOvr",
  "∫": "subSup",
  "∬": "subSup",
  "∭": "subSup",
  "∮": "subSup",
  "∯": "subSup",
  "∰": "subSup",
}

/** The equation as plain characters, for search, screen readers and Excel's fallback. */
export function mathText(list: MathList): string {
  return list.map(nodeText).join("")
}

function nodeText(n: MathNode): string {
  switch (n.t) {
    case "r":
      return n.text
    case "raw":
      return n.text
    case "f":
      return `${mathText(n.num)}/${mathText(n.den)}`
    case "sSup":
      return `${mathText(n.e)}^${mathText(n.sup)}`
    case "sSub":
      return `${mathText(n.e)}_${mathText(n.sub)}`
    case "sSubSup":
      return `${mathText(n.e)}_${mathText(n.sub)}^${mathText(n.sup)}`
    case "sPre":
      return `${mathText(n.sub)}${mathText(n.sup)}${mathText(n.e)}`
    case "rad":
      return `√${mathText(n.e)}`
    case "nary":
      return `${n.chr}${mathText(n.sub)}${mathText(n.sup)}${mathText(n.e)}`
    case "d":
      return `${n.beg}${n.items.map(mathText).join(n.sep ?? "|")}${n.end}`
    case "func":
      return `${mathText(n.name)} ${mathText(n.e)}`
    case "m":
      return n.rows.map((r) => r.map(mathText).join(" ")).join("; ")
    case "eqArr":
      return n.rows.map(mathText).join("; ")
    default:
      return mathText(n.e)
  }
}
