import {
  columnName,
  copyTextImport,
  quoteSheet,
  type ColumnFormat,
  type QueryTable,
  type Range,
  type TextEncoding,
  type TextImport,
  type WorkbookQuery,
} from "@cells/core"
import { mashupConnectionXml } from "./mashup"
import { escapeXmlText } from "@cells/ooxml"
import { at, attrText, kid, kids, parseTree, raw } from "./dom"
import { NS_MAIN } from "./constants"

/**
 * Data linked to text files: xl/connections.xml holds each connection (a text connection keeps the file and the Text
 * Import Wizard's choices in textPr), xl/queryTables/queryTableN.xml ties a sheet's range to its connection, and a
 * sheet-level defined name of the same name says where the range is. Connections of other kinds (databases, the
 * web, Power Query) are kept as they were read.
 */

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
export const REL_CONNECTIONS = REL + "connections"
export const REL_QUERY_TABLE = REL + "queryTable"
const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml."
export const CT_CONNECTIONS = CT + "connections+xml"
export const CT_QUERY_TABLE = CT + "queryTable+xml"

const CODE_PAGES: Record<TextEncoding, number> = {
  "utf-8": 65001,
  "windows-1252": 1252,
  "utf-16le": 1200,
  "utf-16be": 1201,
}
const FORMATS = new Set<ColumnFormat>(["general", "text", "skip", "MDY", "DMY", "YMD", "MYD", "DYM", "YDM"])

/** A connection as read: its id, and the text link's parts when Cells can model it. */
export type ConnectionIn = {
  id: number
  xml: string
  text?: { name: string; description: string; refreshOnLoad: boolean; source: string; options: TextImport }
}

const flag = (v: string | undefined, fallback: boolean) => (v === undefined ? fallback : v === "1" || v === "true")

export function readConnections(xml: string | null): Map<number, ConnectionIn> {
  const out = new Map<number, ConnectionIn>()
  if (!xml) return out
  const doc = parseTree(xml)
  for (const node of kids(doc.root, "connection")) {
    const id = Number(at(node, "id"))
    if (!Number.isInteger(id)) continue
    const entry: ConnectionIn = { id, xml: raw(doc, node) }
    const text = kid(node, "textPr")
    if (at(node, "type") === "6" && text && !kid(node, "parameters")) {
      const fields = kids(kid(text, "textFields"), "textField")
      const delimited = flag(at(text, "delimited"), true)
      const codePage = Number(at(text, "codePage") ?? 1252)
      const encoding = (Object.keys(CODE_PAGES) as TextEncoding[]).find((e) => CODE_PAGES[e] === codePage)
      const delimiters = [
        ...(flag(at(text, "tab"), true) ? ["\t"] : []),
        ...(flag(at(text, "semicolon"), false) ? [";"] : []),
        ...(flag(at(text, "comma"), false) ? [","] : []),
        ...(flag(at(text, "space"), false) ? [" "] : []),
        ...(at(text, "delimiter") ? [at(text, "delimiter")!] : []),
      ]
      const qualifier = at(text, "qualifier") ?? "doubleQuote"
      const options: TextImport = {
        encoding,
        kind: delimited ? "delimited" : "fixed",
        delimiters,
        qualifier: qualifier === "none" ? null : qualifier === "singleQuote" ? "'" : '"',
        consecutive: flag(at(text, "consecutive"), false),
        breaks: delimited
          ? []
          : fields
              .map((f) => Number(at(f, "position") ?? 0))
              .filter((p) => p > 0)
              .sort((a, b) => a - b),
        startRow: Math.max(1, Number(at(text, "firstRow") ?? 1)),
        columns: fields.map((f) => {
          const type = (at(f, "type") ?? "general") as ColumnFormat
          return FORMATS.has(type) ? type : "general"
        }),
        decimal: at(text, "decimal") ?? ".",
        thousands: at(text, "thousands") ?? ",",
      }
      // Trailing General columns say nothing (they are General anyway).
      while (options.columns.at(-1) === "general") options.columns.pop()
      if (encoding !== undefined || at(text, "codePage") === undefined)
        entry.text = {
          name: at(node, "name") ?? `Connection${id}`,
          description: at(node, "description") ?? "",
          refreshOnLoad: flag(at(node, "refreshOnLoad"), false),
          source: at(text, "sourceFile") ?? "",
          options,
        }
    }
    out.set(id, entry)
  }
  return out
}

/** A query table part: its name and connection. */
export function readQueryTable(xml: string | null): { name: string; connectionId: number } | null {
  if (!xml) return null
  const root = parseTree(xml).root
  const name = at(root, "name")
  const connectionId = Number(at(root, "connectionId"))
  return name && Number.isInteger(connectionId) ? { name, connectionId } : null
}

/** A text connection's element. */
function textConnectionXml(id: number, q: QueryTable): string {
  const o = q.options
  const delimited = o.kind === "delimited"
  const has = (d: string) => delimited && o.delimiters.includes(d)
  const custom = delimited ? o.delimiters.find((d) => !["\t", ";", ",", " "].includes(d)) : undefined
  const fields = delimited
    ? o.columns.map((c) => `<textField${attrText({ type: c === "general" ? undefined : c })}/>`)
    : [0, ...o.breaks].map(
        (position, i) =>
          `<textField${attrText({ type: (o.columns[i] ?? "general") === "general" ? undefined : o.columns[i], position: position || undefined })}/>`,
      )
  const textFields = fields.length ? `<textFields count="${fields.length}">${fields.join("")}</textFields>` : ""
  const textPr =
    `<textPr${attrText({
      prompt: "0",
      codePage: CODE_PAGES[o.encoding ?? "utf-8"],
      sourceFile: q.source,
      firstRow: o.startRow > 1 ? o.startRow : undefined,
      delimited: delimited ? undefined : "0",
      decimal: o.decimal !== "." ? o.decimal : undefined,
      thousands: o.thousands !== "," ? o.thousands : undefined,
      tab: has("\t") ? undefined : "0",
      semicolon: has(";") || undefined,
      comma: has(",") || undefined,
      space: has(" ") || undefined,
      consecutive: (delimited && o.consecutive) || undefined,
      qualifier: o.qualifier === "'" ? "singleQuote" : o.qualifier === null ? "none" : undefined,
      delimiter: custom,
    })}>` +
    textFields +
    `</textPr>`
  return (
    `<connection${attrText({
      id,
      name: q.connection,
      description: q.description || undefined,
      type: 6,
      refreshedVersion: 8,
      background: "1",
      refreshOnLoad: q.refreshOnLoad || undefined,
      saveData: "1",
    })}>` +
    textPr +
    `</connection>`
  )
}

/** Namespaces Excel's connections use; the revision ones are ignorable (as Excel marks them). */
const KNOWN_NS: Record<string, string> = {
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
  x15: "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
  xr16: "http://schemas.microsoft.com/office/spreadsheetml/2017/revision16",
}
const IGNORABLE = ["x14", "x15", "xr16"]

/**
 * The root's declarations for prefixes the kept connections use (Excel declares them on the root, e.g. xr16:uid);
 * without them the part isn't well-formed and Excel won't open the file.
 */
function rootNamespaces(items: readonly string[]): string {
  const used = new Set<string>()
  for (const xml of items)
    for (const m of xml.matchAll(/[<\s/]([A-Za-z][\w.-]*):[A-Za-z][\w.-]*/g)) if (m[1] !== "xmlns") used.add(m[1])
  const prefixes = Object.keys(KNOWN_NS).filter((p) => used.has(p))
  if (!prefixes.length) return ""
  const ignorable = prefixes.filter((p) => IGNORABLE.includes(p))
  const all = ignorable.length && !prefixes.includes("mc") ? ["mc", ...prefixes] : prefixes
  return (
    all.map((p) => ` xmlns:${p}="${KNOWN_NS[p]}"`).join("") +
    (ignorable.length ? ` mc:Ignorable="${ignorable.join(" ")}"` : "")
  )
}

/** xl/connections.xml: the kept connections, then Cells' text links (ids after the kept ones); null when none. */
export function connectionsXml(
  kept: readonly ConnectionIn[],
  links: readonly QueryTable[],
  queries: readonly WorkbookQuery[] = [],
): { xml: string; ids: Map<QueryTable, number>; queryIds: Map<string, number> } | null {
  const ids = new Map<QueryTable, number>()
  const queryIds = new Map<string, number>()
  if (!kept.length && !links.length && !queries.length) return null
  let next = Math.max(0, ...kept.map((k) => k.id)) + 1
  const items = kept.map((k) => k.xml)
  for (const q of links) {
    ids.set(q, next)
    items.push(textConnectionXml(next++, q))
  }
  // Loaded queries (Get Data): Excel's Power Query connection for each.
  for (const q of queries) {
    queryIds.set(q.name, next)
    items.push(mashupConnectionXml(next++, q))
  }
  return {
    xml: `<connections xmlns="${NS_MAIN}"${rootNamespaces(items)}>${items.join("")}</connections>`,
    ids,
    queryIds,
  }
}

/** A query table part for a text link. Cells refreshes by overwriting and clearing (growShrinkType). */
export function queryTableXml(q: QueryTable, connectionId: number): string {
  return `<queryTable xmlns="${NS_MAIN}"${attrText({
    name: q.name,
    refreshOnLoad: q.refreshOnLoad || undefined,
    growShrinkType: "overwriteClear",
    connectionId,
    autoFormatId: 16,
    applyNumberFormats: "0",
    applyBorderFormats: "0",
    applyFontFormats: "0",
    applyPatternFormats: "0",
    applyAlignmentFormats: "0",
    applyWidthHeightFormats: "0",
  })}/>`
}

/** The sheet-level defined name that places a query table. */
export function queryTableNameXml(sheetName: string, index: number, q: QueryTable): string {
  const abs = (row: number, col: number) => "$" + columnName(col) + "$" + (row + 1)
  const r = q.range
  const ref = quoteSheet(sheetName) + "!" + abs(r.r0, r.c0) + ":" + abs(r.r1, r.c1)
  return `<definedName name="${escapeXmlText(q.name)}" localSheetId="${index}">${escapeXmlText(ref)}</definedName>`
}

/** A query table from its part, connection and range. */
export function queryTableFrom(name: string, connection: ConnectionIn, range: Range): QueryTable | null {
  const t = connection.text
  if (!t) return null
  return {
    name,
    connection: t.name,
    description: t.description,
    range,
    source: t.source,
    options: copyTextImport(t.options),
    refreshOnLoad: t.refreshOnLoad,
  }
}
