import { escapeXml } from "@cells/ooxml"
import { at, kids, type XNode } from "../xlsx/dom"

/**
 * Number formats between OpenDocument's data styles (number:number-style, date-style… [ODF 1.3] 16.29) and Excel's
 * format codes. Both directions cover what spreadsheets use: numbers with decimals, grouping and scaling,
 * percentages, currencies, scientific, fractions, text, dates and times (elapsed hours too), colours, and the
 * positive;negative;zero sections (ODF's style:map conditions).
 */

export const NS_NUMBER = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
const NS_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
const NS_FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"

const num = (n: XNode, name: string) => at(n, name, NS_NUMBER)

const COLORS: Record<string, string> = {
  "#000000": "Black",
  "#FFFFFF": "White",
  "#FF0000": "Red",
  "#00FF00": "Green",
  "#0000FF": "Blue",
  "#FFFF00": "Yellow",
  "#FF00FF": "Magenta",
  "#00FFFF": "Cyan",
}

/** A literal as a format code writes it: a few characters stand for themselves, the rest go in quotes. */
function literal(text: string): string {
  if (!text) return ""
  if (/^[-+()/:!^&'~{}<>= $€£¥%]+$/.test(text)) return text.replace(/ /g, " ")
  return `"${text.replace(/"/g, '\\"')}"`
}

function integerPart(min: number, grouping: boolean): string {
  if (!grouping) return min > 0 ? "0".repeat(min) : "#"
  const digits = "0".repeat(min).padStart(4, "#")
  return `${digits.slice(0, -3)},${digits.slice(-3)}`
}

function decimals(n: XNode): string {
  const places = Number(num(n, "decimal-places") ?? 0)
  const min = Number(num(n, "min-decimal-places") ?? places)
  if (!places) return ""
  return "." + "0".repeat(Math.min(min, places)) + "#".repeat(Math.max(0, places - min))
}

/** One data style as a format code section (no conditions). */
function sectionOf(style: XNode, elapsed: boolean): string {
  let out = ""
  let color = ""
  const props = style.children.find((c) => c.local === "text-properties")
  const fc = props && at(props, "color", NS_FO)?.toUpperCase()
  if (fc) color = COLORS[fc] ? `[${COLORS[fc]}]` : ""
  const onlyNumber = style.children.filter((c) => c.local !== "text-properties" && c.local !== "map").length === 1
  for (const c of style.children) {
    switch (c.local) {
      case "number": {
        const hasPlaces = num(c, "decimal-places") !== undefined
        const grouping = num(c, "grouping") === "true"
        const min = Number(num(c, "min-integer-digits") ?? (hasPlaces || grouping ? 1 : 0))
        if (onlyNumber && !hasPlaces && !grouping && style.local === "number-style" && min <= 1) {
          out += "General"
          break
        }
        out += integerPart(min, grouping) + decimals(c)
        const factor = Number(num(c, "display-factor") ?? 1)
        if (factor > 1) out += ",".repeat(Math.round(Math.log10(factor) / 3))
        break
      }
      case "scientific-number": {
        const min = Number(num(c, "min-integer-digits") ?? 1)
        out += integerPart(min, num(c, "grouping") === "true") + decimals(c)
        out += "E+" + "0".repeat(Number(num(c, "min-exponent-digits") ?? 2))
        break
      }
      case "fraction": {
        const min = Number(num(c, "min-integer-digits") ?? 0)
        const numer = Number(num(c, "min-numerator-digits") ?? 1)
        const denom = num(c, "denominator-value")
        const denomDigits = Number(num(c, "min-denominator-digits") ?? 1)
        out += (min > 0 ? "0".repeat(min) + " " : "# ") + "?".repeat(numer) + "/" + (denom ?? "?".repeat(denomDigits))
        break
      }
      case "currency-symbol": {
        const sym = c.text
        out += sym === "$" || sym === "€" || sym === "£" || sym === "¥" ? sym : `[$${sym}]`
        break
      }
      case "text":
        out += literal(c.text)
        break
      case "text-content":
        out += "@"
        break
      case "day":
        out += num(c, "style") === "long" ? "dd" : "d"
        break
      case "month": {
        const long = num(c, "style") === "long"
        out += num(c, "textual") === "true" ? (long ? "mmmm" : "mmm") : long ? "mm" : "m"
        break
      }
      case "year":
        out += num(c, "style") === "long" ? "yyyy" : "yy"
        break
      case "day-of-week":
        out += num(c, "style") === "long" ? "dddd" : "ddd"
        break
      case "hours": {
        const h = num(c, "style") === "long" ? "hh" : "h"
        out += elapsed ? `[${h}]` : h
        break
      }
      case "minutes":
        out += num(c, "style") === "long" ? "mm" : "m"
        break
      case "seconds": {
        out += num(c, "style") === "long" ? "ss" : "s"
        const places = Number(num(c, "decimal-places") ?? 0)
        if (places) out += "." + "0".repeat(places)
        break
      }
      case "am-pm":
        out += "AM/PM"
        break
      case "boolean":
        out += "General"
        break
    }
  }
  return color + out
}

/** The condition of a style:map ("value()>=0") as a format code's ("[>=0]"), or null. */
const condition = (text: string) => {
  const m = /^value\(\)\s*(<=|>=|<>|!=|=|<|>)\s*(-?[\d.]+)$/.exec(text.trim())
  return m ? { op: m[1] === "!=" ? "<>" : m[1], value: Number(m[2]) } : null
}

/** A data style (and the styles its maps point at, by name) as an Excel format code. */
export function formatFromOds(style: XNode, byName: (name: string) => XNode | undefined): string {
  const elapsed = (s: XNode) => s.local === "time-style" && at(s, "truncate-on-overflow", NS_NUMBER) === "false"
  const main = sectionOf(style, elapsed(style))
  const maps = kids(style, "map")
    .map((m) => ({
      cond: condition(at(m, "condition", NS_STYLE) ?? ""),
      target: byName(at(m, "apply-style-name", NS_STYLE) ?? ""),
    }))
    .filter((m): m is { cond: { op: string; value: number }; target: XNode } => !!m.cond && !!m.target)
  if (!maps.length) return main
  const text = (m: (typeof maps)[number]) => sectionOf(m.target, elapsed(m.target))
  const is = (m: (typeof maps)[number] | undefined, op: string, value = 0) =>
    m?.cond.op === op && m.cond.value === value
  // The usual shapes: positive (or not negative); negative; zero.
  if (maps.length === 1 && is(maps[0], ">=")) return `${text(maps[0])};${main}`
  if (maps.length === 2 && is(maps[0], ">") && is(maps[1], "<")) return `${text(maps[0])};${text(maps[1])};${main}`
  if (maps.length === 1 && is(maps[0], ">")) return `${text(maps[0])};${main};${main}`
  return [...maps.map((m) => `[${m.cond.op}${m.cond.value}]${text(m)}`), main].slice(0, 3).join(";")
}

// ——— Excel format code → ODF data style ———

type Part =
  | { kind: "lit"; text: string }
  | { kind: "date"; code: string }
  | { kind: "num"; code: string }
  | { kind: "text" }
  | { kind: "currency"; text: string }

type Section = { parts: Part[]; color: string | null; cond: { op: string; value: number } | null }

const COLOR_HEX: Record<string, string> = Object.fromEntries(
  Object.entries(COLORS).map(([hex, name]) => [name.toUpperCase(), hex]),
)

function sections(code: string): Section[] {
  const out: Section[] = []
  let cur: Section = { parts: [], color: null, cond: null }
  const lit = (text: string) => {
    const last = cur.parts[cur.parts.length - 1]
    if (last?.kind === "lit") last.text += text
    else cur.parts.push({ kind: "lit", text })
  }
  let i = 0
  while (i < code.length) {
    const ch = code[i]
    if (ch === ";") {
      out.push(cur)
      cur = { parts: [], color: null, cond: null }
      i++
    } else if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      lit(code.slice(i + 1, end < 0 ? code.length : end))
      i = end < 0 ? code.length : end + 1
    } else if (ch === "\\") {
      lit(code[i + 1] ?? "")
      i += 2
    } else if (ch === "_") {
      lit(" ")
      i += 2
    } else if (ch === "*") {
      i += 2
    } else if (ch === "[") {
      const end = code.indexOf("]", i)
      const inner = code.slice(i + 1, end < 0 ? code.length : end)
      i = end < 0 ? code.length : end + 1
      const cond = /^(<=|>=|<>|=|<|>)(-?[\d.]+)$/.exec(inner)
      if (cond) cur.cond = { op: cond[1], value: Number(cond[2]) }
      else if (/^\$/.test(inner)) {
        const sym = inner.slice(1).split("-")[0]
        if (sym) cur.parts.push({ kind: "currency", text: sym })
      } else if (/^h+$|^m+$|^s+$/i.test(inner)) cur.parts.push({ kind: "date", code: `[${inner.toLowerCase()}]` })
      else if (COLOR_HEX[inner.toUpperCase()]) cur.color = COLOR_HEX[inner.toUpperCase()]
    } else if (/^(AM\/PM|A\/P)/i.test(code.slice(i))) {
      const m = /^(AM\/PM|A\/P)/i.exec(code.slice(i))![0]
      cur.parts.push({ kind: "date", code: "AM/PM" })
      i += m.length
    } else if (/[dmyhse]/i.test(ch) && !/[eE][+-]/.test(code.slice(i, i + 2))) {
      let j = i
      while (j < code.length && code[j].toLowerCase() === ch.toLowerCase()) j++
      cur.parts.push({ kind: "date", code: code.slice(i, j).toLowerCase() })
      i = j
    } else if (/[0#?.,%Ee/]/.test(ch) || (/[+-]/.test(ch) && /[Ee]/.test(code[i - 1] ?? ""))) {
      let j = i
      while (
        j < code.length &&
        (/[0#?.,%/ ]/.test(code[j]) ||
          /[Ee][+-]/.test(code.slice(j, j + 2)) ||
          (/[+-]/.test(code[j]) && /[Ee]/.test(code[j - 1])))
      ) {
        if (code[j] === " " && !/[0#?]/.test(code[j + 1] ?? "")) break
        j++
      }
      const last = cur.parts[cur.parts.length - 1]
      if (last?.kind === "num") last.code += code.slice(i, j)
      else cur.parts.push({ kind: "num", code: code.slice(i, j) })
      i = j
    } else if (ch === "@") {
      cur.parts.push({ kind: "text" })
      i++
    } else {
      lit(ch)
      i++
    }
  }
  out.push(cur)
  return out
}

const isDate = (s: Section) => s.parts.some((p) => p.kind === "date")

function numberElement(code: string): string {
  const pct = code.includes("%")
  const c = code.replace(/%/g, "")
  const sci = /[Ee][+-]/.exec(c)
  if (sci) {
    const mantissa = c.slice(0, sci.index)
    const [intPart, dec = ""] = mantissa.split(".")
    return `<number:scientific-number number:decimal-places="${dec.length}" number:min-decimal-places="${dec.replace(/#/g, "").length}" number:min-integer-digits="${(intPart.match(/0/g) ?? []).length}" number:min-exponent-digits="${c.slice(sci.index + 2).length}"${intPart.includes(",") ? ` number:grouping="true"` : ""}/>`
  }
  if (c.includes("/")) {
    const m = /^([#0]*)\s*([?#0]+)\/([?#0]+|\d+)$/.exec(c.trim())
    if (m) {
      const denom =
        /^\d+$/.test(m[3]) && !/^0+$/.test(m[3])
          ? ` number:denominator-value="${m[3]}"`
          : ` number:min-denominator-digits="${m[3].length}"`
      return `<number:fraction number:min-integer-digits="${(m[1].match(/0/g) ?? []).length}" number:min-numerator-digits="${m[2].length}"${denom}/>`
    }
  }
  const trailing = /,+$/.exec(c)?.[0].length ?? 0
  const body = trailing ? c.slice(0, -trailing) : c
  const [intPart, dec = ""] = body.split(".")
  const grouping = intPart.includes(",")
  const minInt = (intPart.match(/0/g) ?? []).length
  const factor = trailing ? ` number:display-factor="${10 ** (3 * trailing)}"` : ""
  void pct
  return `<number:number number:decimal-places="${dec.length}" number:min-decimal-places="${(dec.match(/0/g) ?? []).length}" number:min-integer-digits="${minInt}"${grouping ? ` number:grouping="true"` : ""}${factor}/>`
}

function dateElements(s: Section): { xml: string; time: boolean; elapsed: boolean } {
  let xml = ""
  let time = true
  let elapsed = false
  const parts = s.parts
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i]
    if (p.kind === "lit") {
      xml += `<number:text>${escapeXml(p.text)}</number:text>`
      continue
    }
    if (p.kind === "num") {
      // Fractions of a second (ss.00): the dot and zeros after seconds.
      if (/^\.0+$/.test(p.code) && /<number:seconds[^>]*\/>$/.test(xml)) {
        xml = xml.replace(
          /<number:seconds([^>]*)\/>$/,
          `<number:seconds$1 number:decimal-places="${p.code.length - 1}"/>`,
        )
      } else xml += `<number:text>${escapeXml(p.code)}</number:text>`
      continue
    }
    if (p.kind !== "date") continue
    const code = p.code
    const long = (n: number) => (code.replace(/[[\]]/g, "").length >= n ? ` number:style="long"` : "")
    // m is minutes after an hour or before a second, months otherwise.
    const prev = parts
      .slice(0, i)
      .reverse()
      .find((q) => q.kind === "date") as { code: string } | undefined
    const next = parts.slice(i + 1).find((q) => q.kind === "date") as { code: string } | undefined
    if (code.startsWith("[")) {
      elapsed = true
      const k = code[1]
      xml +=
        k === "h"
          ? `<number:hours${long(2)}/>`
          : k === "m"
            ? `<number:minutes${long(2)}/>`
            : `<number:seconds${long(2)}/>`
    } else if (code[0] === "d") {
      if (code.length >= 3) xml += `<number:day-of-week${code.length >= 4 ? ` number:style="long"` : ""}/>`
      else xml += `<number:day${long(2)}/>`
      time = false
    } else if (code[0] === "y") {
      xml += `<number:year${code.length >= 3 ? ` number:style="long"` : ""}/>`
      time = false
    } else if (code[0] === "m") {
      const minutes = code.length <= 2 && (/^\[?h/.test(prev?.code ?? "") || /^\[?s/.test(next?.code ?? ""))
      if (minutes) xml += `<number:minutes${long(2)}/>`
      else {
        time = false
        if (code.length >= 3)
          xml += `<number:month number:textual="true"${code.length >= 4 ? ` number:style="long"` : ""}/>`
        else xml += `<number:month${long(2)}/>`
      }
    } else if (code[0] === "h") xml += `<number:hours${long(2)}/>`
    else if (code[0] === "s") xml += `<number:seconds${long(2)}/>`
    else if (code === "AM/PM") xml += `<number:am-pm/>`
    else if (code[0] === "e") {
      xml += `<number:year number:style="long"/>`
      time = false
    }
  }
  return { xml, time, elapsed }
}

/** One section as an ODF data style element named `name` (and its family: number, date, time…). */
function sectionStyle(name: string, s: Section, maps = ""): string {
  const color = s.color ? `<style:text-properties fo:color="${s.color}"/>` : ""
  if (isDate(s)) {
    const d = dateElements(s)
    const tag = d.time ? "time-style" : "date-style"
    const trunc = d.elapsed ? ` number:truncate-on-overflow="false"` : ""
    return `<number:${tag} style:name="${name}"${trunc}>${color}${d.xml}${maps}</number:${tag}>`
  }
  const text = s.parts.some((p) => p.kind === "text")
  const pct = s.parts.some((p) => p.kind === "num" && p.code.includes("%"))
  const cur = s.parts.some((p) => p.kind === "currency")
  let xml = ""
  for (const p of s.parts) {
    if (p.kind === "lit") xml += `<number:text>${escapeXml(p.text)}</number:text>`
    else if (p.kind === "text") xml += `<number:text-content/>`
    else if (p.kind === "currency") xml += `<number:currency-symbol>${escapeXml(p.text)}</number:currency-symbol>`
    else if (p.kind === "num") {
      xml += numberElement(p.code)
      if (p.code.includes("%")) xml += `<number:text>%</number:text>`
    }
  }
  const tag = text ? "text-style" : pct ? "percentage-style" : cur ? "currency-style" : "number-style"
  return `<number:${tag} style:name="${name}">${color}${xml}${maps}</number:${tag}>`
}

/**
 * An Excel format code as ODF data styles (the main one named `name`, its sections `nameP0`…), and the value type
 * cells in it are written with. General is a plain number style.
 */
export function formatToOds(
  code: string,
  name: string,
): { xml: string; type: "float" | "percentage" | "date" | "time" | "currency" | "string" } {
  if (!code || /^general$/i.test(code))
    return {
      xml: `<number:number-style style:name="${name}"><number:number number:min-integer-digits="1"/></number:number-style>`,
      type: "float",
    }
  const list = sections(code).filter((s) => s.parts.length || s.color)
  if (!list.length) return formatToOds("General", name)
  const first = list[0]
  const type = isDate(first)
    ? dateElements(first).time
      ? "time"
      : "date"
    : first.parts.some((p) => p.kind === "num" && p.code.includes("%"))
      ? "percentage"
      : first.parts.some((p) => p.kind === "currency")
        ? "currency"
        : first.parts.some((p) => p.kind === "text") && !first.parts.some((p) => p.kind === "num")
          ? "string"
          : "float"
  if (list.length === 1 && !first.cond) return { xml: sectionStyle(name, first), type }
  // Sections: the last plain one is the style itself; the others apply by their conditions.
  const conds: { op: string; value: number }[] = []
  if (list.some((s) => s.cond)) list.forEach((s) => conds.push(s.cond ?? { op: "", value: 0 }))
  else if (list.length === 2) conds.push({ op: ">=", value: 0 })
  else conds.push({ op: ">", value: 0 }, { op: "<", value: 0 })
  const mapped = list.slice(0, Math.min(list.length - 1, conds.length))
  const mainSection =
    list.length === 2 && !list.some((s) => s.cond) ? list[1] : list[list.length >= 3 ? 2 : list.length - 1]
  let xml = ""
  let maps = ""
  mapped.forEach((s, i) => {
    const c = conds[i]
    if (!c?.op) return
    const sub = `${name}P${i}`
    xml += sectionStyle(sub, s)
    maps += `<style:map style:condition="value()${escapeXml(c.op)}${c.value}" style:apply-style-name="${sub}"/>`
  })
  // Excel's negative section has no sign of its own: in ODF the negative style's number shows none either.
  return { xml: xml + sectionStyle(name, mainSection, maps), type }
}
