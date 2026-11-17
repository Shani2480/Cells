import { columnName, quoteSheet, type Range, type Sheet, type Workbook } from "@cells/core"
import {
  allowsMacros,
  mainContentTypeFor,
  partKey,
  writePackage,
  type OpcPackage,
  type OpcPart,
  type OpcRelationship,
  type RelationshipGroup,
} from "../opc"
import { escapeXmlText, xmlBytes } from "@cells/ooxml"
import { personsXml, writeComments } from "./comments"
import {
  CT_APP,
  CT_COMMENTS,
  CT_CORE,
  CT_METADATA,
  CT_PERSON,
  CT_SHARED_STRINGS,
  CT_STYLES,
  CT_TABLE,
  CT_THREADED,
  CT_VML,
  CT_WORKSHEET,
  NS_MAIN,
  NS_R,
  REL_APP,
  REL_CALC_CHAIN,
  REL_COMMENTS,
  REL_CORE,
  REL_METADATA,
  REL_OFFICE_DOCUMENT,
  REL_PERSON,
  REL_SHARED_STRINGS,
  REL_STYLES,
  REL_TABLE,
  REL_THREADED,
  REL_VML,
  REL_WORKSHEET,
} from "./constants"
import { assemble, attrText, kid, raw, type Kept } from "./dom"
import { writeDrawing } from "./drawing"
import {
  CT_EXTERNAL_LINK,
  externalLinkXml,
  LinkTable,
  linkTargets,
  REL_EXTERNAL_LINK,
  REL_EXTERNAL_LINK_PATH,
} from "./links"
import {
  connectionsXml,
  CT_CONNECTIONS,
  CT_QUERY_TABLE,
  queryTableNameXml,
  queryTableXml,
  REL_CONNECTIONS,
  REL_QUERY_TABLE,
} from "./connections"
import { printNamesXml } from "./page-setup"
import { formulaToFile } from "./formula"
import type { XlsxSource } from "./read"
import { DEFAULT_VIEW, type SheetView } from "./sheet-read"
import { rangeRef, SharedStrings, writeSheet } from "./sheet-write"
import { structuredToFile, tableXml } from "./table"
import { StylesOut } from "./styles"

export type SaveOptions = {
  /** The file the workbook was opened from: what Cells didn't read is written back from it. */
  source?: XlsxSource | null
  /** Decides the package type (.xlsx, .xlsm, .xltx, .xltm). */
  fileName?: string
  views?: readonly SheetView[]
  activeSheet?: number
  /** Author for a new file's properties. */
  author?: string
  /** For a new file's properties (and tests that need the same bytes). */
  now?: Date
}

const WORKBOOK_ORDER = [
  "fileVersion", "fileSharing", "workbookPr", "AlternateContent", "revisionPtr", "absPath", "workbookProtection",
  "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr", "oleSize",
  "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr",
  "webPublishObjects", "extLst",
] // prettier-ignore

/** Excel's metadata for dynamic arrays: cell metadata 1 marks a formula as spilling (`cm="1"`). */
const METADATA_XML =
  `<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
  `<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>` +
  `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
  `<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`

/** Workbook relationships Cells writes itself (or drops: calcChain, and rich values whose metadata it regenerates). */
const OWN_WORKBOOK_RELS = new Set([
  REL_WORKSHEET,
  REL_STYLES,
  REL_SHARED_STRINGS,
  REL_CALC_CHAIN,
  REL_METADATA,
  REL_PERSON,
  REL_CONNECTIONS,
])
const RICH_VALUE_REL = /richvalue|richstyles|rdarray|rdsupportingpropertybag/i
const VBA_REL = /\/vbaProject$/i
const CHARTSHEET_REL = /\/(chartsheet|dialogsheet|xlMacrosheet|xlIntlMacrosheet)$/

/** A relationship target from one part to another, relative to the source's folder. */
export function relativeTarget(from: string, to: string): string {
  const a = from.split("/").slice(1, -1)
  const b = to.split("/").slice(1)
  let i = 0
  while (i < a.length && i < b.length - 1 && a[i].toLowerCase() === b[i].toLowerCase()) i++
  return [...a.slice(i).map(() => ".."), ...b.slice(i)].join("/")
}

/** Saves the workbook as an Excel package. */
export async function saveXlsx(wb: Workbook, options: SaveOptions = {}): Promise<Uint8Array> {
  const source = options.source ?? null
  const fileName = options.fileName ?? source?.pkg.fileName ?? "Book1.xlsx"
  const parts = new Map<string, OpcPart>()
  const relGroups = new Map<string, OpcRelationship[]>()
  const src = source?.pkg
  if (src) {
    for (const p of src.parts) parts.set(partKey(p.name), { ...p })
    for (const g of src.relationships)
      relGroups.set(
        g.source === "" ? "" : partKey(g.source),
        g.items.map((r) => ({ ...r })),
      )
  }
  const main = src?.mainPartName ?? "/xl/workbook.xml"
  const put = (name: string, contentType: string, xml: string) =>
    parts.set(partKey(name), { name, contentType, data: xmlBytes(xml) })
  const remove = (name: string | null) => {
    if (!name) return
    parts.delete(partKey(name))
    relGroups.delete(partKey(name))
  }
  /** A part and what only it uses: a chart's style, colours, shapes and embedded workbook go with it. */
  const removeTree = (name: string) => {
    const key = partKey(name)
    const children = relGroups.get(key) ?? []
    remove(name)
    for (const r of children)
      if (r.targetMode !== "External" && r.partName && /^\/xl\/(charts|drawings|embeddings)\//i.test(r.partName))
        removeTree(r.partName)
  }
  const rels = (part: string) => {
    const key = part === "" ? "" : partKey(part)
    let list = relGroups.get(key)
    if (!list) relGroups.set(key, (list = []))
    return list
  }
  const newId = (list: OpcRelationship[]) => {
    const used = new Set(list.map((r) => r.id))
    let n = 1
    while (used.has(`rId${n}`)) n++
    return `rId${n}`
  }
  const addRel = (from: string, type: string, to: string): string => {
    const list = rels(from)
    const id = newId(list)
    list.push({ id, type, target: relativeTarget(from, to), targetMode: "Internal", partName: to })
    return id
  }
  /** A free part name: `/xl/worksheets/sheet{n}.xml`. */
  const taken = new Set<string>()
  const freeName = (pattern: (n: number) => string) => {
    let n = 1
    while (parts.has(partKey(pattern(n))) || taken.has(partKey(pattern(n)))) n++
    const name = pattern(n)
    taken.add(partKey(name))
    return name
  }

  // ── Workbook relationships: keep the ones Cells doesn't write ──
  const wbRels = rels(main)
  const dropped = wbRels.filter(
    (r) =>
      OWN_WORKBOOK_RELS.has(r.type) ||
      RICH_VALUE_REL.test(r.type) ||
      CHARTSHEET_REL.test(r.type) ||
      (VBA_REL.test(r.type) && !allowsMacros(fileName)),
  )
  // Links to other workbooks: Cells writes them anew, unless the file had links of other kinds (kept as read).
  const ownLinks = source?.linksModelled !== false
  if (ownLinks) for (const r of wbRels) if (r.type === REL_EXTERNAL_LINK && !dropped.includes(r)) dropped.push(r)
  const links = ownLinks
    ? new LinkTable(
        wb.links.list,
        wb.sheets.map((s) => s.name),
      )
    : null
  const keepWorksheet = new Set<string>()
  for (const sheet of wb.sheets) {
    const s = source?.sheets.get(sheet)
    if (s) keepWorksheet.add(partKey(s.partName))
  }
  for (const r of dropped) {
    // A worksheet that still belongs to a sheet is rewritten in place; the rest go.
    if (r.type === REL_WORKSHEET && r.partName && keepWorksheet.has(partKey(r.partName))) continue
    if (r.type === REL_STYLES || r.type === REL_SHARED_STRINGS) continue
    remove(r.partName)
  }
  const stylesName = wbRels.find((r) => r.type === REL_STYLES)?.partName ?? "/xl/styles.xml"
  const sstName = wbRels.find((r) => r.type === REL_SHARED_STRINGS)?.partName ?? "/xl/sharedStrings.xml"
  taken.add(partKey(stylesName))
  taken.add(partKey(sstName))
  for (const name of keepWorksheet) taken.add(name)
  relGroups.set(
    partKey(main),
    wbRels.filter((r) => !dropped.includes(r)),
  )

  const styles = new StylesOut(source?.styles ?? null, source?.reuse ?? new Map())
  const strings = new SharedStrings(source?.rich)
  const spills = new Map<number, Map<string, Range>>()
  for (const a of wb.calc.spillAreas()) {
    let m = spills.get(a.sheet)
    if (!m) spills.set(a.sheet, (m = new Map()))
    m.set(`${a.row},${a.col}`, a.range)
  }
  const views = options.views ?? []
  const active = Math.min(Math.max(0, options.activeSheet ?? 0), wb.sheets.length - 1)
  const authors = new Set<string>()
  let dynamic = false
  const sheetEntries: string[] = []
  const usedTables = new Set<string>()
  // Connections: the kept ones and one per text link; each link's query table goes with its sheet.
  const textLinks = wb.sheets.flatMap((s) => s.queryTables)
  const conns = connectionsXml(source?.keptConnections ?? [], textLinks)
  const usedQueryTables = new Set<string>()

  wb.sheets.forEach((sheet, i) => {
    const s = source?.sheets.get(sheet)
    const partName = s?.partName ?? freeName((n) => `/xl/worksheets/sheet${n}.xml`)
    // Rebuild the sheet's relationships: kept ones, then the notes' parts.
    const own = rels(partName)
    for (const r of [...own]) {
      const keep = s?.rels.some((k) => k.id === r.id)
      if (!keep) {
        own.splice(own.indexOf(r), 1)
        if (r.type === REL_COMMENTS || r.type === REL_THREADED || r.type === REL_VML) remove(r.partName)
      }
    }
    const notes = writeComments(
      sheet,
      i,
      (c) => sheet.cols.size(c),
      (r) => sheet.rows.size(r),
    )
    let legacyDrawing: string | null = null
    if (notes) {
      const comments = freeName((n) => `/xl/comments${n}.xml`)
      put(comments, CT_COMMENTS, notes.legacy)
      addRel(partName, REL_COMMENTS, comments)
      const vml = freeName((n) => `/xl/drawings/vmlDrawing${n}.vml`)
      parts.set(partKey(vml), { name: vml, contentType: CT_VML, data: new TextEncoder().encode(notes.vml) })
      legacyDrawing = addRel(partName, REL_VML, vml)
      if (notes.threaded) {
        const threaded = freeName((n) => `/xl/threadedComments/threadedComment${n}.xml`)
        put(threaded, CT_THREADED, notes.threaded)
        addRel(partName, REL_THREADED, threaded)
        for (const a of notes.persons) authors.add(a)
      }
    }
    // Tables: each back in its own part (a new one gets a new part).
    const tableParts: string[] = []
    for (const t of sheet.tables) {
      const part = s?.tables.find((x) => x.id === t.id)?.partName ?? freeName((n) => `/xl/tables/table${n}.xml`)
      usedTables.add(partKey(part))
      taken.add(partKey(part))
      const toFile = (f: string) => formulaToFile(structuredToFile("=" + f, t.name))
      put(part, CT_TABLE, tableXml(t, styles, rangeRef, toFile))
      tableParts.push(addRel(partName, REL_TABLE, part))
    }
    for (const r of own) if (r.type === REL_QUERY_TABLE && r.partName) usedQueryTables.add(partKey(r.partName))
    for (const q of sheet.queryTables) {
      const part = freeName((n) => `/xl/queryTables/queryTable${n}.xml`)
      usedQueryTables.add(partKey(part))
      put(part, CT_QUERY_TABLE, queryTableXml(q, conns!.ids.get(q)!))
      addRel(partName, REL_QUERY_TABLE, part)
    }
    const drawing = writeDrawing(wb, sheet, partName, s?.drawing, {
      put,
      removeTree,
      addRel,
      rels,
      freeName,
    })
    const view = { ...(views[i] ?? DEFAULT_VIEW), selected: i === active }
    const written = writeSheet({
      sheet,
      index: i,
      view,
      styleTable: wb.styles,
      styles,
      strings,
      source: s,
      spills: spills.get(i) ?? new Map(),
      legacyDrawing,
      drawing,
      tableParts,
      links,
    })
    dynamic ||= written.dynamic
    put(partName, CT_WORKSHEET, written.xml)
    const rid = addRel(main, REL_WORKSHEET, partName)
    sheetEntries.push(
      `<sheet${attrText({ name: sheet.name, sheetId: i + 1, state: sheet.hidden ? (s?.state === "veryHidden" ? "veryHidden" : "hidden") : undefined })} r:id="${rid}"/>`,
    )
  })

  // Tables that are gone (converted to a range, or their sheet deleted) leave with their parts.
  for (const p of [...parts.values()])
    if (p.contentType === CT_TABLE && !usedTables.has(partKey(p.name))) remove(p.name)
  // So do query tables that Cells wrote anew or whose sheet is gone.
  for (const p of [...parts.values()])
    if (p.contentType === CT_QUERY_TABLE && !usedQueryTables.has(partKey(p.name))) remove(p.name)
  if (conns) {
    const name = freeName((n) => (n === 1 ? "/xl/connections.xml" : `/xl/connections${n}.xml`))
    put(name, CT_CONNECTIONS, conns.xml)
    addRel(main, REL_CONNECTIONS, name)
  }

  put(stylesName, CT_STYLES, styles.xml())
  addRel(main, REL_STYLES, stylesName)
  if (strings.size) {
    put(sstName, CT_SHARED_STRINGS, strings.xml())
    addRel(main, REL_SHARED_STRINGS, sstName)
  } else remove(sstName)
  if (dynamic) {
    const meta = freeName((n) => (n === 1 ? "/xl/metadata.xml" : `/xl/metadata${n}.xml`))
    put(meta, CT_METADATA, METADATA_XML)
    addRel(main, REL_METADATA, meta)
  }
  if (authors.size) {
    const persons = freeName((n) => (n === 1 ? "/xl/persons/person.xml" : `/xl/persons/person${n}.xml`))
    put(persons, CT_PERSON, personsXml(authors))
    addRel(main, REL_PERSON, persons)
  }

  // ── workbook.xml ──
  const namesXml = definedNamesXml(wb, source, links)
  const references: string[] = []
  if (links)
    links.books.forEach((book) => {
      const part = freeName((n) => `/xl/externalLinks/externalLink${n}.xml`)
      const targets = linkTargets(book)
      const own = rels(part)
      const target = newId(own)
      own.push({
        id: target,
        type: REL_EXTERNAL_LINK_PATH,
        target: targets.target,
        targetMode: "External",
        partName: null,
      })
      let absolute: string | undefined
      if (targets.absolute) {
        absolute = newId(own)
        own.push({
          id: absolute,
          type: REL_EXTERNAL_LINK_PATH,
          target: targets.absolute,
          targetMode: "External",
          partName: null,
        })
      }
      put(part, CT_EXTERNAL_LINK, externalLinkXml(book, { target, absolute }))
      references.push(`<externalReference r:id="${addRel(main, REL_EXTERNAL_LINK, part)}"/>`)
    })
  const generated: Record<string, string> = {
    sheets: `<sheets>${sheetEntries.join("")}</sheets>`,
    ...(links
      ? {
          externalReferences: references.length
            ? `<externalReferences>${references.join("")}</externalReferences>`
            : "",
        }
      : {}),
    definedNames: namesXml,
    workbookProtection: protectionXml(wb),
    bookViews: bookViewsXml(source, active),
    calcPr: `<calcPr${attrText({ calcId: 191029, calcMode: wb.calc.mode === "manual" ? "manual" : undefined })}/>`,
  }
  let wbXml: string
  if (source) {
    const doc = source.workbook
    const kept: Kept[] = doc.root.children
      .filter((c) => !(c.local in generated))
      .map((c) => ({ local: c.local, xml: raw(doc, c) }))
    const open = doc.source.slice(doc.root.start, doc.root.tagEnd).replace(/\/?>$/, "")
    wbXml = `${open}>${assemble(WORKBOOK_ORDER, kept, generated)}</${doc.root.name}>`
  } else {
    wbXml =
      `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
      assemble(
        WORKBOOK_ORDER,
        [{ local: "workbookPr", xml: `<workbookPr defaultThemeVersion="202300"/>` }],
        generated,
      ) +
      `</workbook>`
  }
  put(main, mainContentTypeFor(fileName), wbXml)

  if (!src) {
    rels("").push({
      id: "rId1",
      type: REL_OFFICE_DOCUMENT,
      target: "xl/workbook.xml",
      targetMode: "Internal",
      partName: main,
    })
    const now = (options.now ?? new Date()).toISOString().replace(/\.\d+Z$/, "Z")
    const author = escapeXmlText(options.author ?? "")
    put(
      "/docProps/core.xml",
      CT_CORE,
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
        `<dc:creator>${author}</dc:creator><cp:lastModifiedBy>${author}</cp:lastModifiedBy>` +
        `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
    )
    rels("").push({
      id: "rId2",
      type: REL_CORE,
      target: "docProps/core.xml",
      targetMode: "Internal",
      partName: "/docProps/core.xml",
    })
    put(
      "/docProps/app.xml",
      CT_APP,
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application></Properties>`,
    )
    rels("").push({
      id: "rId3",
      type: REL_APP,
      target: "docProps/app.xml",
      targetMode: "Internal",
      partName: "/docProps/app.xml",
    })
  }

  // ── Drop parts nothing points at any more (old sheets, calcChain, replaced notes) ──
  const reachable = new Set<string>()
  const visit = (key: string) => {
    for (const r of relGroups.get(key) ?? []) {
      if (r.targetMode !== "Internal" || !r.partName) continue
      const k = partKey(r.partName)
      if (reachable.has(k)) continue
      reachable.add(k)
      visit(k)
    }
  }
  visit("")
  for (const [key] of parts) if (!reachable.has(key)) parts.delete(key)
  for (const [key, list] of relGroups) {
    if (key !== "" && !parts.has(key)) relGroups.delete(key)
    else
      relGroups.set(
        key,
        list.filter((r) => r.targetMode === "External" || (r.partName && parts.has(partKey(r.partName)))),
      )
  }

  const relationships: RelationshipGroup[] = []
  for (const [key, items] of relGroups)
    relationships.push({ source: key === "" ? "" : (parts.get(key)?.name ?? "/" + key), items })
  const pkg: OpcPackage = {
    fileName,
    mainPartName: main,
    parts: [...parts.values()],
    defaults: src?.defaults ?? [],
    overrides: [],
    relationships,
    warnings: [],
    zipEntries: [],
  }
  return writePackage(pkg)
}

function definedNamesXml(wb: Workbook, source: XlsxSource | null, links: LinkTable | null): string {
  const items: string[] = []
  for (const n of wb.names.all()) {
    const ref = links ? links.toFile("=" + n.ref) : "=" + n.ref
    items.push(
      `<definedName${attrText({ name: n.name, comment: n.comment })}>${escapeXmlText(formulaToFile(ref))}</definedName>`,
    )
  }
  // Excel's hidden name for each sheet's AutoFilter range.
  wb.sheets.forEach((sheet, index) => {
    if (!sheet.autoFilter) return
    const r = sheet.autoFilter.range
    const abs = (row: number, col: number) => "$" + columnName(col) + "$" + (row + 1)
    const ref = quoteSheet(sheet.name) + "!" + abs(r.r0, r.c0) + ":" + abs(r.r1, r.c1)
    items.push(
      `<definedName name="_xlnm._FilterDatabase" localSheetId="${index}" hidden="1">${escapeXmlText(ref)}</definedName>`,
    )
  })
  // Each sheet's print area and titles, and the names that place its text links.
  wb.sheets.forEach((sheet, index) => items.push(...printNamesXml(sheet.name, index, sheet.pageSetup)))
  wb.sheets.forEach((sheet, index) => {
    for (const q of sheet.queryTables) items.push(queryTableNameXml(sheet.name, index, q))
  })
  for (const k of source?.keptNames ?? []) {
    if (!k.sheet) {
      items.push(k.xml)
      continue
    }
    const index = wb.sheets.indexOf(k.sheet as Sheet)
    if (index < 0) continue
    items.push(k.xml.replace(/\blocalSheetId="\d+"/, `localSheetId="${index}"`))
  }
  return items.length ? `<definedNames>${items.join("")}</definedNames>` : ""
}

function protectionXml(wb: Workbook): string {
  const p = wb.structure.protection
  if (!p) return ""
  const attrs: Record<string, string | undefined> = {}
  if (p.hash) {
    attrs.workbookAlgorithmName = p.hash.algorithm
    attrs.workbookHashValue = p.hash.hash
    attrs.workbookSaltValue = p.hash.salt
    attrs.workbookSpinCount = String(p.hash.spinCount)
  } else if (p.password) attrs.workbookPassword = p.password
  attrs.lockStructure = "1"
  return `<workbookProtection${attrText(attrs)}/>`
}

function bookViewsXml(source: XlsxSource | null, active: number): string {
  const doc = source?.workbook
  const views = doc ? kid(doc.root, "bookViews") : undefined
  const first = views ? kid(views, "workbookView") : undefined
  if (!doc || !first) return `<bookViews><workbookView${attrText({ activeTab: active || undefined })}/></bookViews>`
  let tag = raw(doc, first)
    .replace(/\s(activeTab|firstSheet)="\d+"/g, "")
    .replace(/^<((?:\w+:)?workbookView)\b/, (m) => (active ? `${m} activeTab="${active}"` : m))
  // Other views (rare: New Window) are kept after the first.
  const rest = views!.children.filter((c) => c !== first).map((c) => raw(doc, c))
  tag += rest.join("")
  return `${doc.source.slice(views!.start, views!.tagEnd)}${tag}</${views!.name}>`
}
