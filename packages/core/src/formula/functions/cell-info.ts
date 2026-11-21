/** CELL (a cell's address, value, format and layout) and INFO (about the workbook and the app). */
import { columnName } from "../../address"
import { isError, type CellError } from "../../values"
import { NA, toText, VALUE } from "../coerce"
import { toScalar } from "../evaluate"
import { quoteSheet } from "../print"
import { isRef, matrix, type Scalar } from "../types"
import { define } from "./registry"

/** The code CELL("format") gives for a number format: G, F2, ,0, C2, P0, S2, D1–D9, with "-" and "()" added. */
export function formatCode(code: string | undefined): string {
  if (!code || /^general$/i.test(code)) return "G"
  const sections = code.split(";")
  const first = sections[0].replace(/"[^"]*"|\\./g, "")
  const negative = (sections[1] ?? "").replace(/"[^"]*"|\\./g, "")
  const plain = first.replace(/\[[^\]]*\]|_.|\*./g, "")
  const suffix =
    (/\[(red|black|blue|cyan|green|magenta|white|yellow|color\s*\d+)\]/i.test(negative) ? "-" : "") +
    (/[()]/.test(plain) ? "()" : "")
  if (/[dmyhs]/i.test(plain) && !/[0#?]/.test(plain)) {
    const p = plain.toLowerCase().replace(/\s+/g, " ").trim()
    const table: Record<string, string> = {
      "d-mmm-yy": "D1",
      "dd-mmm-yy": "D1",
      "d-mmm": "D2",
      "dd-mmm": "D2",
      "mmm-yy": "D3",
      "mm/dd": "D5",
      "h:mm:ss am/pm": "D6",
      "h:mm am/pm": "D7",
      "h:mm:ss": "D8",
      "h:mm": "D9",
    }
    return (table[p] ?? "D4") + suffix
  }
  const decimals = (plain.split(/E[+-]/i)[0].split(".")[1] ?? "").match(/[0#?]/g)?.length ?? 0
  if (/%/.test(plain)) return "P" + decimals + suffix
  if (/E[+-]/i.test(plain)) return "S" + decimals + suffix
  if (/\?\/|\/\?/.test(plain)) return "G"
  if (/[$€£¥]|\[\$/.test(sections[0])) return "C" + decimals + suffix
  if (/[0#?]/.test(plain)) return (/,/.test(plain) ? "," : "F") + decimals + suffix
  return "G"
}

const INFO_TYPES = new Set([
  "address",
  "col",
  "color",
  "contents",
  "filename",
  "format",
  "parentheses",
  "prefix",
  "protect",
  "row",
  "type",
  "width",
])

define("CELL", {
  min: 1,
  max: 2,
  volatile: true,
  call: (a, ctx) => {
    const kind = toText(toScalar(a[0], ctx))
    if (isError(kind)) return kind
    const info = kind.toLowerCase()
    if (!INFO_TYPES.has(info)) return VALUE
    let sheet = ctx.sheet
    let row = ctx.row
    let col = ctx.col
    if (a.length > 1) {
      const ref = a[1]
      if (!isRef(ref)) return VALUE
      ;({ sheet, r0: row, c0: col } = ref.areas[0])
    }
    const value = (): Scalar => ctx.cell(sheet, row, col)
    const look = () => ctx.cellInfo?.(sheet, row, col) ?? { style: {}, width: 64, hidden: false, custom: false }
    switch (info) {
      case "address": {
        const own = sheet === ctx.sheet ? "" : quoteSheet(ctx.sheetName(sheet) ?? "") + "!"
        return `${own}$${columnName(col)}$${row + 1}`
      }
      case "row":
        return row + 1
      case "col":
        return col + 1
      case "contents":
        return value()
      case "type": {
        const v = value()
        return v === null ? "b" : typeof v === "string" ? "l" : "v"
      }
      case "filename":
        return ""
      case "format":
        return formatCode(look().style.numberFormat)
      case "color": {
        const negative = (look().style.numberFormat ?? "").split(";")[1] ?? ""
        return /\[(red|black|blue|cyan|green|magenta|white|yellow|color\s*\d+)\]/i.test(negative) ? 1 : 0
      }
      case "parentheses": {
        const first = (look().style.numberFormat ?? "").split(";")[0].replace(/"[^"]*"|\\.|_.|\[[^\]]*\]/g, "")
        return /[()]/.test(first) ? 1 : 0
      }
      case "prefix": {
        if (typeof value() !== "string") return ""
        const h = look().style.hAlign
        return h === "right" ? '"' : h === "center" ? "^" : h === "fill" ? "\\" : "'"
      }
      case "protect":
        return look().style.locked === false ? 0 : 1
      case "width": {
        // Characters of the default font's digit (7 px, with 5 px of padding), rounded; and whether it's the default.
        const l = look()
        return matrix(1, 2, [Math.round(Math.max(0, l.width - 5) / 7), !l.custom])
      }
    }
    return VALUE
  },
})

function platform(): string {
  const ua = (globalThis as { navigator?: { userAgent?: string } }).navigator?.userAgent ?? ""
  return /Mac/.test(ua) ? "mac" : "pcdos"
}

define("INFO", {
  min: 1,
  max: 1,
  volatile: true,
  params: ["scalar"],
  call: (a, ctx): Scalar | CellError => {
    const t = toText(a[0] as Scalar)
    if (isError(t)) return t
    switch (t.toLowerCase()) {
      case "directory":
        return ""
      case "numfile":
        return ctx.sheetCount()
      case "origin":
        return "$A:$A$1"
      case "osversion": {
        const ua = (globalThis as { navigator?: { userAgent?: string } }).navigator?.userAgent ?? ""
        return /Windows/.test(ua) ? "Windows" : /Mac/.test(ua) ? "macOS" : /Android/.test(ua) ? "Android" : "Linux"
      }
      case "recalc":
        return "Automatic"
      case "release":
        return "16.0"
      case "system":
        return platform()
      case "memavail":
      case "memused":
      case "totmem":
        return NA
    }
    return VALUE
  },
})
