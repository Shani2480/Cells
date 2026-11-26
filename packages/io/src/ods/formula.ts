import { parseStructured, quoteSheet, tokenize, type Table } from "@cells/core"
import { letNames } from "../xlsx/formula"

/**
 * Formulas between OpenDocument's OpenFormula (`of:=SUM([.A1:.B2];[Sheet2.C3])`, [ODF 1.3] part 4) and Excel's
 * (`SUM(A1:B2,Sheet2!C3)`): references in brackets with a dot before the cell, `;` between arguments, `|` between an
 * array's rows, `!` for intersection and `~` for union, and newer Excel functions under COM.MICROSOFT. (with a few
 * renamed), as LibreOffice and Excel write them.
 */

/** Excel's names that OpenFormula spells differently. */
const TO_ODS: Record<string, string> = {
  CHIDIST: "LEGACY.CHIDIST",
  CHIINV: "LEGACY.CHIINV",
  CHITEST: "LEGACY.CHITEST",
  FDIST: "LEGACY.FDIST",
  FINV: "LEGACY.FINV",
  NORMSDIST: "LEGACY.NORMSDIST",
  NORMSINV: "LEGACY.NORMSINV",
  TDIST: "LEGACY.TDIST",
  "F.DIST": "FDIST",
  "F.INV": "FINV",
  FORMULATEXT: "FORMULA",
  "SKEW.P": "SKEWP",
}
const FROM_ODS: Record<string, string> = Object.fromEntries(Object.entries(TO_ODS).map(([a, b]) => [b, a]))
// OpenFormula's own CEILING and FLOOR take a mode; Excel's closest are the .MATH ones.
FROM_ODS.CEILING = "CEILING.MATH"
FROM_ODS.FLOOR = "FLOOR.MATH"

/** Functions LibreOffice writes as COM.MICROSOFT.name (Excel's newer ones, and those with Excel's own rules). */
const MICROSOFT = new Set(
  `AGGREGATE ARRAYTOTEXT BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BYCOL BYROW CEILING CEILING.MATH
CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS CONCAT CONFIDENCE.NORM
CONFIDENCE.T COVARIANCE.P COVARIANCE.S DROP ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST.RT F.INV.RT
F.TEST FILTER FILTERXML FLOOR FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY
FORECAST.ETS.STAT FORECAST.LINEAR GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GROUPBY HSTACK HYPGEOM.DIST IFS IMCOSH IMCOT
IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISO.CEILING ISOMITTED LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS
MINIFS MODE.MULT MODE.SNGL NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PERCENTILE.EXC
PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY
RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST SCAN SEQUENCE SORT SORTBY STDEV.P STDEV.S SWITCH T.DIST
T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRIMRANGE UNIQUE
VALUETOTEXT VAR.P VAR.S VSTACK WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH Z.TEST`.split(/\s+/),
)

const PREFIXES = /^(?:COM\.MICROSOFT\.|_xlfn\.|_xlws\.|ORG\.OPENOFFICE\.|ORG\.LIBREOFFICE\.)/i

function functionFromOds(name: string): string {
  const upper = name.toUpperCase()
  if (FROM_ODS[upper]) return FROM_ODS[upper]
  if (upper.startsWith("LEGACY.")) return upper.slice(7)
  if (PREFIXES.test(name)) return name.replace(PREFIXES, "").toUpperCase()
  return name
}

function functionToOds(name: string): string {
  const upper = name.toUpperCase()
  if (TO_ODS[upper]) return TO_ODS[upper]
  return MICROSOFT.has(upper) ? `COM.MICROSOFT.${upper}` : name
}

/** A quoted sheet name's text, or an unquoted one, from the start of `s`; with how many characters it took. */
function readSheet(s: string): { name: string; length: number } | null {
  let i = 0
  if (s[i] === "$") i++
  if (s[i] === "'") {
    let name = ""
    i++
    while (i < s.length) {
      if (s[i] === "'" && s[i + 1] === "'") {
        name += "'"
        i += 2
      } else if (s[i] === "'") {
        i++
        break
      } else name += s[i++]
    }
    return s[i] === "." ? { name, length: i + 1 } : null
  }
  const dot = s.indexOf(".", i)
  if (dot < 0) return null
  return { name: s.slice(i, dot), length: dot + 1 }
}

/** The inside of an OpenFormula reference ([.A1:.B2], [Sheet1.A1], [Sheet1.A1:Sheet3.B2]) as Excel's, or null. */
function referenceFromOds(inner: string): string | null {
  if (inner.includes("#REF!")) return "#REF!"
  if (/^'[^']*'#|^[a-z]+:\/\//i.test(inner)) return null // another file
  // The parts on each side of the colon (a colon in a quoted sheet name doesn't count).
  const parts: string[] = []
  let quoted = false
  let start = 0
  for (let i = 0; i < inner.length; i++) {
    if (inner[i] === "'") quoted = !quoted
    else if (inner[i] === ":" && !quoted) {
      parts.push(inner.slice(start, i))
      start = i + 1
    }
  }
  parts.push(inner.slice(start))
  if (parts.length > 2) return null
  const read = parts.map((p) => {
    if (p.startsWith(".")) return { sheet: null as string | null, cell: p.slice(1) }
    const sheet = readSheet(p)
    return sheet ? { sheet: sheet.name, cell: p.slice(sheet.length) } : { sheet: null, cell: p }
  })
  if (!read.every((r) => /^\$?[A-Za-z]{0,3}\$?\d*$/.test(r.cell) && r.cell.replace(/\$/g, ""))) return null
  const [a, b] = read
  let prefix = ""
  if (a.sheet !== null) {
    if (b && b.sheet !== null && b.sheet !== a.sheet) {
      const plain = quoteSheet(a.sheet) === a.sheet && quoteSheet(b.sheet) === b.sheet
      prefix = plain ? `${a.sheet}:${b.sheet}!` : `'${a.sheet.replace(/'/g, "''")}:${b.sheet.replace(/'/g, "''")}'!`
    } else prefix = quoteSheet(a.sheet) + "!"
  }
  // One cell written as a range ([$Sheet1.A1:$Sheet3.A1], a 3-D reference) is that cell; H:H and 1:1 stay as they are.
  if (b && a.cell === b.cell && /[A-Za-z]\$?\d/.test(a.cell)) return prefix + a.cell
  if (b) {
    // A whole column or row written out as its cells (H$1:H$1048576, $A5:$XFD5) is Excel's H:H or 5:5.
    const c0 = /^(\$?[A-Za-z]+)\$1$/.exec(a.cell)
    const c1 = /^(\$?[A-Za-z]+)\$1048576$/.exec(b.cell)
    if (c0 && c1) return `${prefix}${c0[1]}:${c1[1]}`
    const r0 = /^\$A(\$?\d+)$/i.exec(a.cell)
    const r1 = /^\$XFD(\$?\d+)$/i.exec(b.cell)
    if (r0 && r1) return `${prefix}${r0[1]}:${r1[1]}`
  }
  return prefix + (b ? `${a.cell}:${b.cell}` : a.cell)
}

/**
 * An OpenDocument formula attribute as Excel's formula text (no "="), or null when it can't be written in Excel's
 * terms (a reference to another file); the cell then keeps its value.
 */
export function formulaFromOds(attr: string): string | null {
  let text = attr
  const ns = /^([a-z]+):/.exec(text)
  if (ns) {
    text = text.slice(ns[0].length)
    // Excel 2007's .ods kept Excel's own syntax.
    if (ns[1] === "msoxl") return text.replace(/^=/, "")
  }
  text = text.replace(/^=/, "")
  let out = ""
  let arrays = 0
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    if (ch === '"') {
      let j = i + 1
      while (j < text.length && !(text[j] === '"' && text[j + 1] !== '"')) j += text[j] === '"' ? 2 : 1
      out += text.slice(i, j + 1)
      i = j + 1
    } else if (ch === "[") {
      let j = i + 1
      let quoted = false
      while (j < text.length && (quoted || text[j] !== "]")) {
        if (text[j] === "'") quoted = !quoted
        j++
      }
      const ref = referenceFromOds(text.slice(i + 1, j))
      if (ref === null) return null
      out += ref
      i = j + 1
    } else if (ch === "{") {
      arrays++
      out += ch
      i++
    } else if (ch === "}") {
      arrays--
      out += ch
      i++
    } else if (ch === ";") {
      out += ","
      i++
    } else if (ch === "|" && arrays > 0) {
      out += ";"
      i++
    } else if (ch === "~") {
      out += ","
      i++
    } else if (ch === "!") {
      out += " "
      i++
    } else if (ch === "$" && text[i + 1] === "$") {
      i += 2
    } else if (/[A-Za-z_\\À-￿]/.test(ch)) {
      let j = i
      while (j < text.length && /[\w.\\À-￿]/.test(text[j])) j++
      const word = text.slice(i, j)
      if (text[j] === "(") {
        const upper = word.toUpperCase()
        if ((upper === "TRUE" || upper === "FALSE") && text[j + 1] === ")") {
          out += upper
          i = j + 2
          continue
        }
        out += /^_xlpm\./i.test(word) ? word.slice(6) : functionFromOds(word)
      } else out += word.replace(/^_xlpm\./i, "")
      i = j
    } else {
      out += ch
      i++
    }
  }
  return out
}

/** A sheet name as an OpenFormula reference writes it: absolute ($Sheet1), as Excel's references to sheets are. */
const odsSheet = (name: string) => "$" + (/^[A-Za-z_][\w]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)

const REF = /^(?:((?:'(?:[^']|'')+'|[^!]+))!)?(.+)$/

function sheetNames(prefix: string): [string, string] {
  const unq = prefix.startsWith("'") ? prefix.slice(1, -1).replace(/''/g, "'") : prefix
  // Sheet1:Sheet3 (a 3-D reference), quoted as one or not.
  const colon = unq.indexOf(":")
  return colon > 0 ? [unq.slice(0, colon), unq.slice(colon + 1)] : [unq, unq]
}

/** An Excel reference token (Sheet1!A1:B2, A:A, 3:5) as OpenFormula's [Sheet1.A1:.B2]. */
function referenceToOds(text: string): string {
  const m = REF.exec(text)!
  const body = m[2]
  const [first, last] = m[1] ? sheetNames(m[1]) : [null, null]
  const [a, b] = body.split(":")
  // Whole columns and rows (A:A, 3:5) as the cells they span.
  const whole = (p: string, end: boolean) => {
    if (/^\$?[A-Za-z]+$/.test(p)) return `${p}${end ? "$1048576" : "$1"}`
    if (/^\$?\d+$/.test(p)) return `${end ? "$XFD" : "$A"}${p}`
    return p
  }
  const left = b === undefined ? a : whole(a, false)
  const right = b === undefined ? undefined : whole(b, true)
  const s1 = first !== null ? odsSheet(first) : ""
  const s2 = last !== null && last !== first ? odsSheet(last) : ""
  // A 3-D reference to one cell (Sheet1:Sheet3!A1) spans that cell on each sheet.
  if (right === undefined) return s2 ? `[${s1}.${left}:${s2}.${left}]` : `[${s1}.${left}]`
  return `[${s1}.${left}:${s2}.${right}]`
}

export type OdsFormulaContext = {
  /** The tables, for structured references (OpenDocument has none: they become the ranges they stand for). */
  tables: { table: Table; sheet: string }[]
  /** Where the formula is, for [@Column]. */
  row: number
  sheet: string
}

function structuredToOds(text: string, ctx: OdsFormulaContext): string | null {
  const spec = parseStructured(text)
  if (!spec) return null
  const found = spec.table
    ? ctx.tables.find((t) => t.table.name.toUpperCase() === spec.table!.toUpperCase())
    : ctx.tables.find((t) => t.sheet === ctx.sheet && ctx.row >= t.table.range.r0 && ctx.row <= t.table.range.r1)
  if (!found) return null
  const { table, sheet } = found
  const r = table.range
  const col = (name: string | null) =>
    name === null ? -1 : table.columns.findIndex((c) => c.name.toUpperCase() === name.toUpperCase())
  const c0 = spec.c0 === null ? r.c0 : r.c0 + col(spec.c0)
  const c1 = spec.c1 === null ? r.c1 : r.c0 + col(spec.c1)
  if (c0 < r.c0 || c1 < r.c0) return null
  const dataTop = r.r0 + (table.header ? 1 : 0)
  const dataBottom = r.r1 - (table.totals ? 1 : 0)
  let r0: number
  let r1: number
  if (spec.thisRow) r0 = r1 = ctx.row
  else {
    const rows: [number, number][] = []
    if (spec.headers && table.header) rows.push([r.r0, r.r0])
    if (spec.data || (!spec.headers && !spec.totals)) rows.push([dataTop, dataBottom])
    if (spec.totals && table.totals) rows.push([r.r1, r.r1])
    if (!rows.length) return null
    r0 = Math.min(...rows.map((x) => x[0]))
    r1 = Math.max(...rows.map((x) => x[1]))
  }
  // This row ([@Qty]) stays relative to the row, as a filled-down formula would be.
  const cell = (row: number, c: number) => `$${columnLetters(c)}${spec.thisRow ? "" : "$"}${row + 1}`
  const s = sheet === ctx.sheet ? "" : odsSheet(sheet)
  return r0 === r1 && c0 === c1 ? `[${s}.${cell(r0, c0)}]` : `[${s}.${cell(r0, c0)}:.${cell(r1, c1)}]`
}

function columnLetters(c: number): string {
  let s = ""
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}

const ENDS = new Set(["ref", "name", "close", "structured", "number", "string", "bool", "error", "arrayClose"])
const STARTS = new Set(["ref", "name", "open", "func", "structured", "arrayOpen"])

/**
 * A Cells formula ("=…") as an OpenDocument formula attribute ("of:=…"), or null when it has something OpenFormula
 * can't say (a spilled range reference, A1#); the cell is then written with its value only.
 */
export function formulaToOds(input: string, ctx: OdsFormulaContext): string | null {
  const tokens = tokenize(input.startsWith("=") ? input.slice(1) : input)
  // LET's names and LAMBDA's parameters are marked _xlpm., as Excel and LibreOffice write them.
  const params = letNames(tokens)
  let out = ""
  const stack: ("func" | "group" | "array")[] = []
  let prev: string | null = null
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k]
    switch (t.kind) {
      case "ref":
        out += referenceToOds(t.text)
        break
      case "structured": {
        const ref = structuredToOds(t.text, ctx)
        if (ref === null) return null
        out += ref
        break
      }
      case "func":
        out += params.has(t.text.toUpperCase()) ? "_xlpm." + t.text : functionToOds(t.text)
        break
      case "name":
        out += params.has(t.text.toUpperCase()) ? "_xlpm." + t.text : t.text
        break
      case "open":
        stack.push(prev === "func" ? "func" : "group")
        out += "("
        break
      case "close":
        stack.pop()
        out += ")"
        break
      case "arrayOpen":
        stack.push("array")
        out += "{"
        break
      case "arrayClose":
        stack.pop()
        out += "}"
        break
      case "comma":
        out += stack[stack.length - 1] === "group" ? "~" : ";"
        break
      case "semicolon":
        out += stack[stack.length - 1] === "array" ? "|" : ";"
        break
      case "bool":
        // Plain TRUE, as Excel writes it: Excel reads TRUE() as a call (and can't in an array); LibreOffice reads both.
        out += t.text.toUpperCase()
        break
      case "spill":
      case "invalid":
        return null
      case "space": {
        // Between two operands a space is Excel's intersection: OpenFormula's "!".
        const next = tokens.slice(k + 1).find((x) => x.kind !== "space")
        if (prev && ENDS.has(prev) && next && STARTS.has(next.kind)) out += "!"
        else out += t.text
        break
      }
      default:
        out += t.text
    }
    if (t.kind !== "space") prev = t.kind
  }
  return "of:=" + out
}
