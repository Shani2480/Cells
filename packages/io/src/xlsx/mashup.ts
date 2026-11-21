import { dataUrl, dataUrlBytes, type WorkbookQuery } from "@cells/core"
import { escapeXmlText, readZip, storedZip } from "@cells/ooxml"

/**
 * Power Query in xlsx: the queries live in a customXml item holding a DataMashup element, whose text is base64 of
 * MS-QDEFF's binary: a zip with Formulas/Section1.m (every query as a "shared" member of one section), the
 * permissions, a metadata list (per query: loaded to a table or a connection only…) and permission bindings (signed
 * by the machine that wrote them; Cells writes none, and Excel then takes its defaults). A loaded query reaches its
 * table through a connection (Microsoft.Mashup.OleDb, Location = the query's name) and a query table part.
 */

export const REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
export const REL_CUSTOM_XML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
export const CT_CUSTOM_XML_PROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
export const MASHUP_SCHEMA = "http://schemas.microsoft.com/DataMashup"

/** The parts of a DataMashup Cells keeps to write it back: each query's metadata entries, by name. */
export type MashupIn = {
  queries: { name: string; description: string; formula: string }[]
  /** Each query's StableEntries element as read, to keep what Cells doesn't change. */
  entries: Map<string, string>
  /** The item's text as read, written back as is while the queries stay the same. */
  raw: Uint8Array
}

/** The text of a customXml item (UTF-16 with its byte order mark, or UTF-8). */
export function itemText(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.subarray(2))
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.subarray(2))
  return new TextDecoder().decode(bytes)
}

export const isMashup = (text: string) => /<DataMashup\b[^>]*DataMashup/.test(text.slice(0, 400))

// Base64 through core's data URL helpers (io has no DOM).
const fromBase64 = (b64: string): Uint8Array =>
  dataUrlBytes(`data:application/octet-stream;base64,${b64.replace(/\s+/g, "")}`)?.bytes ?? new Uint8Array()
const toBase64 = (bytes: Uint8Array): string => dataUrl(bytes, "application/octet-stream").replace(/^data:[^,]*,/, "")

const decodeName = (n: string) => (n.startsWith("#") ? n.slice(2, -1).replace(/""/g, '"') : n)

/** The members of Section1.m: name, description and formula (up to its closing semicolon). */
export function readSection(m: string): { name: string; description: string; formula: string }[] {
  const out: { name: string; description: string; formula: string }[] = []
  const head = /(?:\[\s*Description\s*=\s*"((?:[^"]|"")*)"\s*\]\s*)?shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][\w.]*)\s*=\s*/y
  let at = m.search(/;/) + 1 // after "section Section1;"
  while (at > 0 && at < m.length) {
    while (at < m.length && /\s/.test(m[at])) at++
    head.lastIndex = at
    const h = head.exec(m)
    if (!h) break
    let i = head.lastIndex
    // The formula runs to the first semicolon outside strings, quoted names and comments.
    for (; i < m.length; i++) {
      const c = m[i]
      if (c === '"') {
        for (i++; i < m.length; i++)
          if (m[i] === '"')
            if (m[i + 1] === '"') i++
            else break
      } else if (c === "/" && m[i + 1] === "/") i = m.indexOf("\n", i) < 0 ? m.length : m.indexOf("\n", i)
      else if (c === "/" && m[i + 1] === "*") i = m.indexOf("*/", i) < 0 ? m.length : m.indexOf("*/", i) + 1
      else if (c === ";") break
    }
    out.push({
      name: decodeName(h[2]),
      description: (h[1] ?? "").replace(/""/g, '"'),
      formula: m.slice(head.lastIndex, i).trim(),
    })
    at = i + 1
  }
  return out
}

/** The queries of a DataMashup item, or null when it isn't one Cells can read. */
export async function readMashup(bytes: Uint8Array): Promise<MashupIn | null> {
  const text = itemText(bytes)
  if (!isMashup(text)) return null
  const b64 = />([^<]*)<\/DataMashup>/.exec(text)?.[1]
  if (!b64) return null
  try {
    const d = fromBase64(b64)
    const v = new DataView(d.buffer, d.byteOffset, d.byteLength)
    let at = 4
    const block = () => {
      const len = v.getUint32(at, true)
      const part = d.subarray(at + 4, at + 4 + len)
      at += 4 + len
      return part
    }
    const pkg = block()
    block() // permissions
    const meta = block()
    const files = await readZip(pkg)
    const section = files.find((f) => /^Formulas\/Section1\.m$/i.test(f.normalizedName))
    if (!section) return null
    const queries = readSection(new TextDecoder().decode(section.data).replace(/^\uFEFF/, ""))
    const entries = new Map<string, string>()
    if (meta.length >= 8) {
      const mv = new DataView(meta.buffer, meta.byteOffset, meta.byteLength)
      const xml = new TextDecoder().decode(meta.subarray(8, 8 + mv.getUint32(4, true))).replace(/^\uFEFF/, "")
      for (const item of xml.matchAll(/<Item>([\s\S]*?)<\/Item>/g)) {
        const path = /<ItemType>Formula<\/ItemType><ItemPath>Section1\/([^</]*)<\/ItemPath>/.exec(item[1])?.[1]
        const stable = /<StableEntries\s*\/>|<StableEntries>[\s\S]*?<\/StableEntries>/.exec(item[1])?.[0]
        if (path !== undefined && stable) entries.set(decodeURIComponent(path), stable)
      }
    }
    return { queries, entries, raw: bytes }
  } catch {
    return null
  }
}

const mName = (n: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n) ? n : `#"${n.replace(/"/g, '""')}"`)

/** Section1.m for the queries, as Excel writes it. */
export function sectionM(queries: readonly WorkbookQuery[]): string {
  const members = queries.map(
    (q) =>
      (q.description ? `[ Description = "${q.description.replace(/"/g, '""')}" ]\r\n` : "") +
      `shared ${mName(q.name)} = ${q.formula.trim()};`,
  )
  return `section Section1;\r\n\r\n${members.join("\r\n\r\n")}`
}

function guid(): string {
  const hex = Array.from({ length: 32 }, () => Math.floor(Math.random() * 16).toString(16)).join("")
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-a${hex.slice(17, 20)}-${hex.slice(20, 32)}`
}

/** A query's metadata entries: loaded to its table or a connection only (kept entries otherwise as read). */
function stableEntries(q: WorkbookQuery, kept: string | undefined): string {
  const e = (type: string, value: string) => `<Entry Type="${type}" Value="${escapeXmlText(value)}" />`
  const old = new Map<string, string>()
  for (const m of (kept ?? "").matchAll(/<Entry Type="([^"]*)" Value="([^"]*)"\s*\/>/g)) old.set(m[1], m[0])
  const own = new Map<string, string>([
    ["IsPrivate", e("IsPrivate", "l0")],
    ["FillEnabled", e("FillEnabled", q.table ? "l1" : "l0")],
    ["FillObjectType", e("FillObjectType", q.table ? "sTable" : "sConnectionOnly")],
    ["FillToDataModelEnabled", e("FillToDataModelEnabled", "l0")],
    ["QueryID", old.get("QueryID") ?? e("QueryID", `s${guid()}`)],
  ])
  if (q.table) {
    own.set("FillTarget", e("FillTarget", `s${q.table}`))
    own.set("ResultType", e("ResultType", "sTable"))
    own.set("BufferNextRefresh", e("BufferNextRefresh", "l1"))
  }
  // What Excel recorded about the last fill (status, columns, counts) stays while the query still fills a table.
  const keep = q.table ? [...old].filter(([k]) => !own.has(k) && k !== "FillTarget") : []
  return `<StableEntries>${[...own.values(), ...keep.map(([, v]) => v)].join("")}</StableEntries>`
}

const u32 = (n: number) => {
  const b = new Uint8Array(4)
  new DataView(b.buffer).setUint32(0, n, true)
  return b
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

const utf8 = (s: string) => new TextEncoder().encode(s)

/** The DataMashup item for the queries (UTF-16 with a byte order mark, as Excel writes it). */
export function mashupItem(queries: readonly WorkbookQuery[], read: MashupIn | null): Uint8Array {
  const bom = "\uFEFF"
  const pkg = storedZip([
    {
      name: "[Content_Types].xml",
      data: utf8(
        `${bom}<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml" /><Default Extension="m" ContentType="application/x-ms-m" /></Types>`,
      ),
    },
    {
      name: "Config/Package.xml",
      data: utf8(
        `${bom}<?xml version="1.0" encoding="utf-8"?><Package xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Version>2.132.15052.0</Version><MinVersion>2.21.0.0</MinVersion><Culture>en-US</Culture></Package>`,
      ),
    },
    { name: "Formulas/Section1.m", data: utf8(sectionM(queries)) },
  ])
  const permissions = utf8(
    `${bom}<?xml version="1.0" encoding="utf-8"?><PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><CanEvaluateFuturePackages>false</CanEvaluateFuturePackages><FirewallEnabled>true</FirewallEnabled></PermissionList>`,
  )
  const items = [
    `<Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries /></Item>`,
    ...queries.map(
      (q) =>
        `<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/${escapeXmlText(encodeURIComponent(q.name))}</ItemPath></ItemLocation>` +
        `${stableEntries(q, read?.entries.get(q.name))}</Item>`,
    ),
  ]
  const metaXml = utf8(
    `${bom}<?xml version="1.0" encoding="utf-8"?><LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Items>${items.join("")}</Items></LocalPackageMetadataFile>`,
  )
  const emptyZip = storedZip([])
  const meta = concat([u32(0), u32(metaXml.length), metaXml, u32(emptyZip.length), emptyZip])
  const binary = concat([
    u32(0),
    u32(pkg.length),
    pkg,
    u32(permissions.length),
    permissions,
    u32(meta.length),
    meta,
    u32(0),
  ])
  const xml = `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="${MASHUP_SCHEMA}">${toBase64(binary)}</DataMashup>`
  const body = new Uint8Array(2 + xml.length * 2)
  body[0] = 0xff
  body[1] = 0xfe
  for (let i = 0; i < xml.length; i++) {
    body[2 + i * 2] = xml.charCodeAt(i) & 0xff
    body[3 + i * 2] = xml.charCodeAt(i) >> 8
  }
  return body
}

/** The item's properties part (its schema reference), for a new item. */
export function mashupPropsXml(): string {
  const id = guid().toUpperCase()
  return (
    `<ds:datastoreItem ds:itemID="{${id}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${MASHUP_SCHEMA}"/></ds:schemaRefs></ds:datastoreItem>`
  )
}

/** Whether the queries are as read (then the item is written back untouched, keeping Excel's bindings). */
export function sameQueries(
  queries: readonly WorkbookQuery[],
  read: MashupIn | null,
  loads: Map<string, string | null>,
) {
  if (!read || read.queries.length !== queries.length) return false
  return queries.every((q, i) => {
    const r = read.queries[i]
    return (
      r.name === q.name &&
      r.description === q.description &&
      r.formula === q.formula.trim() &&
      loads.get(q.name) === q.table
    )
  })
}

/** The connection Excel makes for a loaded query. */
export function mashupConnectionXml(id: number, q: WorkbookQuery): string {
  const location = /^[A-Za-z0-9_]+$/.test(q.name)
    ? q.name
    : `&quot;${escapeXmlText(q.name).replace(/"/g, "&quot;")}&quot;`
  return (
    `<connection id="${id}" keepAlive="1" name="${escapeXmlText(`Query - ${q.name}`)}"` +
    `${q.description ? ` description="${escapeXmlText(q.description)}"` : ""} type="5" refreshedVersion="8" background="1" saveData="1">` +
    `<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=${location};Extended Properties=&quot;&quot;" ` +
    `command="SELECT * FROM [${escapeXmlText(q.name.replace(/]/g, "]]"))}]"/></connection>`
  )
}

/** The query a connection reads, when it's a Power Query connection. */
export function mashupLocation(connectionXml: string): string | null {
  if (!/Microsoft\.Mashup\.OleDb/i.test(connectionXml)) return null
  const conn = /\bconnection="([^"]*)"/.exec(connectionXml)?.[1]
  if (!conn) return null
  const decoded = conn
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
  const m = /Location=(?:"((?:[^"]|"")*)"|([^;]*))/.exec(decoded)
  return m ? (m[1]?.replace(/""/g, '"') ?? m[2]) : null
}

/** The query table part of a loaded query's table: its fields, one per column, ids matching the table's columns. */
export function mashupQueryTableXml(
  name: string,
  connectionId: number,
  columns: { id: number; name: string }[],
): string {
  const fields = columns.map(
    (c, i) => `<queryTableField id="${i + 1}" name="${escapeXmlText(c.name)}" tableColumnId="${c.id}"/>`,
  )
  return (
    `<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${escapeXmlText(name)}" ` +
    `backgroundRefresh="0" connectionId="${connectionId}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" ` +
    `applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0">` +
    `<queryTableRefresh nextId="${columns.length + 1}"><queryTableFields count="${columns.length}">${fields.join("")}</queryTableFields>` +
    `</queryTableRefresh></queryTable>`
  )
}
