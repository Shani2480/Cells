import {
  bookNameOf,
  errorValue,
  linkKey,
  MAX_COLS,
  parseCell,
  quoteSheet,
  splitBookSheet,
  tokenize,
  type CellValue,
  type ExternalBook,
  type LinkedSheet,
} from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import type { OpcRelationship } from "../opc"
import { NS_MAIN, NS_R } from "./constants"
import { at, attrText, kid, kids, parseTree } from "./dom"

/**
 * Links to other workbooks in .xlsx files. Each linked workbook is an external link part (its sheet names, the names
 * formulas use, and the values last read from it) with an external relationship to the file; workbook.xml lists them
 * in <externalReferences>, and formulas name them by position: [1]Prices!B2, [1]!Rate. Cells shows the file's name
 * instead ('[Price List.xlsx]Prices'!B2) and turns it back into the position when saving.
 */

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
export const REL_EXTERNAL_LINK = REL + "externalLink"
export const REL_EXTERNAL_LINK_PATH = REL + "externalLinkPath"
export const CT_EXTERNAL_LINK = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
const NS_XXL21 = "http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021"
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

const unescapeTarget = (t: string) => {
  try {
    return decodeURIComponent(t.replace(/^file:\/\/\/?/i, ""))
  } catch {
    return t
  }
}
const escapeTarget = (t: string) => t.replace(/%/g, "%25").replace(/ /g, "%20").replace(/#/g, "%23")

/** A linked workbook from its external link part; null for links of other kinds (DDE, OLE). */
export function readExternalLink(xml: string | null, rels: readonly OpcRelationship[]): ExternalBook | null {
  if (!xml) return null
  const root = parseTree(xml).root
  const book = kid(root, "externalBook")
  if (!book) return null
  const target = rels.find((r) => r.id === at(book, "id", NS_R))?.target
  if (target === undefined) return null
  const absoluteId = at(kid(kid(book, "alternateUrls"), "absoluteUrl"), "id", NS_R)
  const absolute = absoluteId ? rels.find((r) => r.id === absoluteId)?.target : undefined
  const sheets: LinkedSheet[] = kids(kid(book, "sheetNames"), "sheetName").map((s) => ({
    name: at(s, "val") ?? "",
    cells: new Map(),
  }))
  for (const data of kids(kid(book, "sheetDataSet"), "sheetData")) {
    const sheet = sheets[Number(at(data, "sheetId"))]
    if (!sheet) continue
    for (const row of kids(data, "row"))
      for (const cell of kids(row, "cell")) {
        const ref = parseCell(at(cell, "r") ?? "")
        const v = kid(cell, "v")?.text
        if (!ref || v === undefined) continue
        const t = at(cell, "t") ?? "n"
        const value: CellValue =
          t === "str" || t === "s" || t === "inlineStr"
            ? v
            : t === "b"
              ? v === "1" || v === "true"
              : t === "e"
                ? errorValue(v as Parameters<typeof errorValue>[0])
                : Number(v)
        if (value == null || (typeof value === "number" && !Number.isFinite(value))) continue
        sheet.cells.set(linkKey(ref.row, ref.col), value)
      }
  }
  const names = kids(kid(book, "definedNames"), "definedName").map((n) => {
    const sheetId = at(n, "sheetId")
    return {
      name: at(n, "name") ?? "",
      refersTo: at(n, "refersTo") ?? "",
      ...(sheetId !== undefined ? { sheet: Number(sheetId) } : {}),
    }
  })
  const name = bookNameOf(unescapeTarget(target))
  return {
    name,
    target: unescapeTarget(target),
    ...(absolute ? { absolute: unescapeTarget(absolute) } : {}),
    sheets,
    names,
  }
}

/** The text of a ref or name token's workbook prefix, split from what follows the "!". */
function prefixOf(text: string): { prefix: string; rest: string } | null {
  if (text.startsWith("'")) {
    let i = 1
    for (; i < text.length; i++) {
      if (text[i] !== "'") continue
      if (text[i + 1] === "'") i++
      else break
    }
    if (text[i + 1] !== "!") return null
    return { prefix: text.slice(1, i).replace(/''/g, "'"), rest: text.slice(i + 2) }
  }
  const bang = text.indexOf("!")
  return bang < 0 ? null : { prefix: text.slice(0, bang), rest: text.slice(bang + 1) }
}

/** "[Book]Sheet" as a formula writes it: bare when both names are plain, quoted otherwise. */
function bookSheet(book: string, sheet: string): string {
  const plain = (s: string) => quoteSheet(s) === s
  return plain(sheet) && /^[\w.]+$/.test(book) ? `[${book}]${sheet}` : `'[${book}]${sheet.replace(/'/g, "''")}'`
}

/** A file formula ("=[1]Prices!B2") with each link written by its file's name ("='[Price List.xlsx]Prices'!B2"). */
export function linksFromFile(formula: string, books: readonly ExternalBook[]): string {
  if (!formula.includes("[") || !books.length) return formula
  let changed = false
  const out = tokenize(formula).map((t) => {
    if (t.kind !== "ref" && t.kind !== "name") return t.text
    const p = prefixOf(t.text)
    const m = p ? /^\[(\d+)\](.*)$/.exec(p.prefix) : null
    const book = m ? books[Number(m[1]) - 1] : undefined
    if (!p || !m || !book) return t.text
    changed = true
    return (m[2] ? bookSheet(book.name, m[2]) : quoteSheet(book.name)) + "!" + p.rest
  })
  return changed ? out.join("") : formula
}

/**
 * The links a save writes: the workbook's own, then any workbook a formula names that has none yet (its values
 * unknown until Workbook Links updates them). `toFile` writes a formula's links by position.
 */
export class LinkTable {
  readonly books: ExternalBook[]
  private readonly sheetNames: Set<string>

  constructor(books: readonly ExternalBook[], sheetNames: readonly string[]) {
    this.books = [...books]
    this.sheetNames = new Set(sheetNames.map((s) => s.toUpperCase()))
  }

  private bookIndex(name: string, sheet: string | null): number {
    const upper = name.toUpperCase()
    let i = this.books.findIndex((b) => b.name.toUpperCase() === upper)
    if (i < 0) {
      this.books.push({ name, target: name, sheets: [], names: [] })
      i = this.books.length - 1
    }
    const book = this.books[i]
    if (sheet && !book.sheets.some((s) => s.name.toUpperCase() === sheet.toUpperCase()))
      this.books[i] = { ...book, sheets: [...book.sheets, { name: sheet, cells: new Map() }] }
    return i
  }

  /** Every workbook the formula names, registered (for the first pass over a workbook's formulas). */
  toFile(formula: string): string {
    if (!formula.includes("!") || (!formula.includes("[") && !this.books.length)) return formula
    let changed = false
    const out = tokenize(formula).map((t) => {
      if (t.kind !== "ref" && t.kind !== "name") return t.text
      const p = prefixOf(t.text)
      if (!p) return t.text
      const parts = splitBookSheet(p.prefix)
      if (parts) {
        changed = true
        const i = this.bookIndex(parts.book, parts.sheet) + 1
        return (parts.sheet ? bookSheet(String(i), parts.sheet) : `[${i}]`) + "!" + p.rest
      }
      // 'Price List.xlsx'!Rate: a linked workbook's name.
      const upper = p.prefix.toUpperCase()
      if (t.kind === "name" && !this.sheetNames.has(upper)) {
        const i = this.books.findIndex((b) => b.name.toUpperCase() === upper)
        if (i >= 0) {
          changed = true
          return `[${i + 1}]!` + p.rest
        }
      }
      return t.text
    })
    return changed ? out.join("") : formula
  }
}

/** A linked workbook's external link part, given its relationships' ids. */
export function externalLinkXml(book: ExternalBook, ids: { target: string; absolute?: string }): string {
  const sheetNames = book.sheets.length
    ? `<sheetNames>${book.sheets.map((s) => `<sheetName${attrText({ val: s.name })}/>`).join("")}</sheetNames>`
    : ""
  const names = book.names.length
    ? `<definedNames>${book.names
        .map((n) => `<definedName${attrText({ name: n.name, refersTo: n.refersTo, sheetId: n.sheet })}/>`)
        .join("")}</definedNames>`
    : ""
  const data = book.sheets
    .map((s, i) => {
      if (!s.cells.size) return ""
      const rows = new Map<number, [number, CellValue][]>()
      for (const [k, v] of s.cells) {
        const row = Math.floor(k / MAX_COLS)
        let list = rows.get(row)
        if (!list) rows.set(row, (list = []))
        list.push([k % MAX_COLS, v])
      }
      const body = [...rows.entries()]
        .sort((a, b) => a[0] - b[0])
        .map(([row, cells]) => {
          const xml = cells
            .sort((a, b) => a[0] - b[0])
            .map(([col, v]) => {
              const r = cellName(row, col)
              if (typeof v === "number") return `<cell r="${r}"><v>${v}</v></cell>`
              if (typeof v === "boolean") return `<cell r="${r}" t="b"><v>${v ? 1 : 0}</v></cell>`
              if (typeof v === "string") return `<cell r="${r}" t="str"><v>${escapeXmlText(v)}</v></cell>`
              return `<cell r="${r}" t="e"><v>${escapeXmlText(v.error)}</v></cell>`
            })
            .join("")
          return `<row r="${row + 1}">${xml}</row>`
        })
        .join("")
      return `<sheetData sheetId="${i}">${body}</sheetData>`
    })
    .join("")
  const dataSet = data ? `<sheetDataSet>${data}</sheetDataSet>` : ""
  const alternate = ids.absolute
    ? `<xxl21:alternateUrls><xxl21:absoluteUrl r:id="${ids.absolute}"/></xxl21:alternateUrls>`
    : ""
  const ns = ids.absolute ? ` xmlns:mc="${NS_MC}" mc:Ignorable="xxl21" xmlns:xxl21="${NS_XXL21}"` : ""
  return (
    `<externalLink xmlns="${NS_MAIN}"${ns}><externalBook xmlns:r="${NS_R}" r:id="${ids.target}">` +
    alternate +
    sheetNames +
    names +
    dataSet +
    `</externalBook></externalLink>`
  )
}

function cellName(row: number, col: number): string {
  let name = ""
  for (let c = col + 1; c > 0; c = Math.floor((c - 1) / 26)) name = String.fromCharCode(65 + ((c - 1) % 26)) + name
  return name + (row + 1)
}

/** The relationship targets of a link: the path as given, and the absolute one as a file URL. */
export function linkTargets(book: ExternalBook): { target: string; absolute?: string } {
  return {
    target: escapeTarget(book.target),
    ...(book.absolute ? { absolute: "file:///" + escapeTarget(book.absolute).replace(/^\/+/, "") } : {}),
  }
}
