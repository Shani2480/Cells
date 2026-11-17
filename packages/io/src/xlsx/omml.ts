import { mathItalic, NARY, plainLetter, type MathList, type MathNode, type TextRun, toLinear } from "@cells/core"
import { escapeXml } from "@cells/ooxml"
import { at, kid, kids, raw, type XDoc, type XNode } from "./dom"

/**
 * Equations as Office Math Markup (OMML, ECMA-376 Part 1 §22.1) inside DrawingML text, as Excel writes them: an
 * `a14:m` holding an `m:oMathPara`, its runs with DrawingML run properties and letters in math italic (𝑥).
 */

export const NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
export const NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"

const NS_M_ATTR = `${NS_M}|val`
const val = (node: XNode | undefined) => node?.nsAttrs[NS_M_ATTR] ?? at(node, "val")
const prop = (pr: XNode | undefined, name: string) => val(kid(pr, name))
const on = (pr: XNode | undefined, name: string) => {
  const n = kid(pr, name)
  if (!n) return false
  const v = val(n)
  return v === undefined || v === "1" || v === "on" || v === "true"
}

// ── Reading ──────────────────────────────────────────────────────────────────────────────────────────────

/** The lines (`m:oMath`) of a math paragraph, or of one `m:oMath`. */
export function readMath(doc: XDoc, node: XNode): MathList[] {
  if (node.local === "oMath") return [readList(doc, node)]
  return kids(node, "oMath").map((m) => readList(doc, m))
}

/** The look of the equation's text (size, colour, bold) from its first run, as a text run without text. */
export function mathLook(node: XNode, readRun: (rPr: XNode | undefined) => TextRun): TextRun | undefined {
  const first = findFirst(node, "r")
  const rPr = first?.children.find((c) => c.local === "rPr" && c.uri !== NS_M)
  return rPr ? readRun(rPr) : undefined
}

function findFirst(node: XNode, local: string): XNode | undefined {
  for (const c of node.children) {
    if (c.local === local && c.uri === NS_M) return c
    const inner = findFirst(c, local)
    if (inner) return inner
  }
  return undefined
}

function readList(doc: XDoc, node: XNode | undefined): MathList {
  const out: MathList = []
  for (const c of node?.children ?? []) {
    if (c.uri !== NS_M || c.local.endsWith("Pr")) continue
    const n = readNode(doc, c)
    if (n) out.push(n)
  }
  return out
}

function readNode(doc: XDoc, n: XNode): MathNode | null {
  const pr = kid(n, `${n.local}Pr`)
  const e = () => readList(doc, kid(n, "e"))
  switch (n.local) {
    case "r": {
      const text = [
        ...kids(n, "t")
          .map((t) => t.text)
          .join(""),
      ]
        .map(plainLetter)
        .join("")
      const rPr = n.children.find((c) => c.local === "rPr" && c.uri === NS_M)
      const sty = on(rPr, "nor") ? "p" : prop(rPr, "sty")
      const run: MathNode = { t: "r", text }
      if (sty === "p" || sty === "b" || sty === "bi") run.sty = sty
      // An upright run with no letters (Word marks digits and operators so) is just text.
      if (run.sty === "p" && !/\p{L}/u.test(text)) delete run.sty
      return run
    }
    case "f": {
      const type = prop(pr, "type")
      return {
        t: "f",
        num: readList(doc, kid(n, "num")),
        den: readList(doc, kid(n, "den")),
        ...(type === "skw" || type === "lin" || type === "noBar" ? { type } : {}),
      }
    }
    case "sSup":
      return { t: "sSup", e: e(), sup: readList(doc, kid(n, "sup")) }
    case "sSub":
      return { t: "sSub", e: e(), sub: readList(doc, kid(n, "sub")) }
    case "sSubSup":
      return { t: "sSubSup", e: e(), sub: readList(doc, kid(n, "sub")), sup: readList(doc, kid(n, "sup")) }
    case "sPre":
      return { t: "sPre", sub: readList(doc, kid(n, "sub")), sup: readList(doc, kid(n, "sup")), e: e() }
    case "rad":
      return { t: "rad", deg: on(pr, "degHide") ? [] : readList(doc, kid(n, "deg")), e: e() }
    case "nary": {
      const chr = prop(pr, "chr") ?? "∫"
      const loc = prop(pr, "limLoc")
      return {
        t: "nary",
        chr,
        sub: on(pr, "subHide") ? [] : readList(doc, kid(n, "sub")),
        sup: on(pr, "supHide") ? [] : readList(doc, kid(n, "sup")),
        e: e(),
        limLoc: loc === "undOvr" || loc === "subSup" ? loc : (NARY[chr] ?? "undOvr"),
      }
    }
    case "d": {
      const sep = prop(pr, "sepChr")
      return {
        t: "d",
        beg: prop(pr, "begChr") ?? "(",
        end: prop(pr, "endChr") ?? ")",
        ...(sep !== undefined ? { sep } : {}),
        items: kids(n, "e").map((x) => readList(doc, x)),
      }
    }
    case "func":
      return { t: "func", name: readList(doc, kid(n, "fName")), e: e() }
    case "acc":
      return { t: "acc", chr: prop(pr, "chr") ?? "̂", e: e() }
    case "bar":
      return { t: "bar", pos: prop(pr, "pos") === "top" ? "top" : "bot", e: e() }
    case "limLow":
      return { t: "limLow", e: e(), lim: readList(doc, kid(n, "lim")) }
    case "limUpp":
      return { t: "limUpp", e: e(), lim: readList(doc, kid(n, "lim")) }
    case "groupChr":
      return { t: "groupChr", chr: prop(pr, "chr") ?? "⏟", pos: prop(pr, "pos") === "top" ? "top" : "bot", e: e() }
    case "m":
      return { t: "m", rows: kids(n, "mr").map((r) => kids(r, "e").map((x) => readList(doc, x))) }
    case "eqArr":
      return { t: "eqArr", rows: kids(n, "e").map((x) => readList(doc, x)) }
    case "box":
      return { t: "box", e: e() }
    case "borderBox":
      return { t: "borderBox", e: e() }
    default: {
      // Phantoms, group characters Cells doesn't model…: kept as they came, shown as their text.
      const text = collectText(n)
      return { t: "raw", xml: raw(doc, n), text }
    }
  }
}

function collectText(n: XNode): string {
  if (n.local === "t") return [...n.text].map(plainLetter).join("")
  return n.children.map(collectText).join("")
}

// ── Writing ──────────────────────────────────────────────────────────────────────────────────────────────

/** Writes OMML with the given DrawingML run properties (`<a:rPr …>`, with `i="1"` for math italic). */
export function mathParaXml(lines: MathList[], look: TextRun | undefined, runPr: (r: TextRun) => string): string {
  const base: TextRun = { text: "", ...look, font: "Cambria Math" }
  const italic = runPr({ ...base, italic: true })
  const upright = runPr({ ...base, italic: false })
  const ctrl = `<m:ctrlPr>${italic}</m:ctrlPr>`
  const pr = (name: string, inner = "") => `<m:${name}Pr>${inner}${ctrl}</m:${name}Pr>`
  const v = (name: string, value: string) => `<m:${name} m:val="${escapeXml(value)}"/>`
  const list = (l: MathList): string => l.map(node).join("")
  const wrap = (tag: string, l: MathList) => `<m:${tag}>${list(l)}</m:${tag}>`

  function run(text: string, sty?: string): string {
    const shown = sty ? text : [...text.replace(/-/g, "−")].map(mathItalic).join("")
    const mPr = sty ? `<m:rPr>${v("sty", sty)}</m:rPr>` : ""
    return `<m:r>${mPr}${sty === "p" ? upright : italic}<m:t>${escapeXml(shown)}</m:t></m:r>`
  }

  function node(n: MathNode): string {
    switch (n.t) {
      case "r":
        return run(n.text, n.sty)
      case "raw":
        return n.xml
      case "f":
        return `<m:f>${pr("f", n.type && n.type !== "bar" ? v("type", n.type) : "")}${wrap("num", n.num)}${wrap("den", n.den)}</m:f>`
      case "sSup":
        return `<m:sSup>${pr("sSup")}${wrap("e", n.e)}${wrap("sup", n.sup)}</m:sSup>`
      case "sSub":
        return `<m:sSub>${pr("sSub")}${wrap("e", n.e)}${wrap("sub", n.sub)}</m:sSub>`
      case "sSubSup":
        return `<m:sSubSup>${pr("sSubSup")}${wrap("e", n.e)}${wrap("sub", n.sub)}${wrap("sup", n.sup)}</m:sSubSup>`
      case "sPre":
        return `<m:sPre>${pr("sPre")}${wrap("sub", n.sub)}${wrap("sup", n.sup)}${wrap("e", n.e)}</m:sPre>`
      case "rad":
        return `<m:rad>${pr("rad", n.deg.length ? "" : v("degHide", "1"))}${wrap("deg", n.deg)}${wrap("e", n.e)}</m:rad>`
      case "nary": {
        const loc = n.limLoc && n.limLoc !== (NARY[n.chr] ?? "undOvr") ? v("limLoc", n.limLoc) : ""
        const inner =
          (n.chr !== "∫" ? v("chr", n.chr) : "") +
          loc +
          (n.sub.length ? "" : v("subHide", "1")) +
          (n.sup.length ? "" : v("supHide", "1"))
        return `<m:nary>${pr("nary", inner)}${wrap("sub", n.sub)}${wrap("sup", n.sup)}${wrap("e", n.e)}</m:nary>`
      }
      case "d": {
        const inner =
          (n.beg !== "(" ? v("begChr", n.beg) : "") +
          (n.end !== ")" ? v("endChr", n.end) : "") +
          (n.sep !== undefined && n.sep !== "|" ? v("sepChr", n.sep) : "")
        return `<m:d>${pr("d", inner)}${n.items.map((i) => wrap("e", i)).join("")}</m:d>`
      }
      case "func":
        return `<m:func>${pr("func")}${wrap("fName", n.name)}${wrap("e", n.e)}</m:func>`
      case "acc":
        return `<m:acc>${pr("acc", n.chr !== "̂" ? v("chr", n.chr) : "")}${wrap("e", n.e)}</m:acc>`
      case "bar":
        return `<m:bar>${pr("bar", n.pos === "top" ? v("pos", "top") : "")}${wrap("e", n.e)}</m:bar>`
      case "limLow":
        return `<m:limLow>${pr("limLow")}${wrap("e", n.e)}${wrap("lim", n.lim)}</m:limLow>`
      case "limUpp":
        return `<m:limUpp>${pr("limUpp")}${wrap("e", n.e)}${wrap("lim", n.lim)}</m:limUpp>`
      case "groupChr":
        return `<m:groupChr>${pr("groupChr", (n.chr !== "⏟" ? v("chr", n.chr) : "") + (n.pos === "top" ? v("pos", "top") : ""))}${wrap("e", n.e)}</m:groupChr>`
      case "m": {
        const cols = Math.max(1, ...n.rows.map((r) => r.length))
        const mPr = `<m:mPr><m:mcs><m:mc><m:mcPr>${v("count", String(cols))}${v("mcJc", "center")}</m:mcPr></m:mc></m:mcs>${ctrl}</m:mPr>`
        return `<m:m>${mPr}${n.rows.map((r) => `<m:mr>${r.map((c) => wrap("e", c)).join("")}</m:mr>`).join("")}</m:m>`
      }
      case "eqArr":
        return `<m:eqArr>${pr("eqArr")}${n.rows.map((r) => wrap("e", r)).join("")}</m:eqArr>`
      case "box":
        return `<m:box>${pr("box")}${wrap("e", n.e)}</m:box>`
      case "borderBox":
        return `<m:borderBox>${pr("borderBox")}${wrap("e", n.e)}</m:borderBox>`
    }
  }

  const body = lines.map((l) => `<m:oMath>${list(l)}</m:oMath>`).join("")
  return (
    `<a14:m><m:oMathPara xmlns:m="${NS_M}"><m:oMathParaPr>${v("jc", "centerGroup")}</m:oMathParaPr>` +
    `${body}</m:oMathPara></a14:m>`
  )
}

/** The equation as Excel's fallback shows it: its linear text. */
export const mathFallbackText = (lines: MathList[]) => lines.map(toLinear)
