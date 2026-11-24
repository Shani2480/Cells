import { withExt } from "./validation"
import { isEmbeddedObject, oleObjectsXml, oleObjectXml, oleShapeXml, writeEmbedding } from "./ole-objects"
import { lastShapeId, pictureVml, signatureShapeXml, withPictures } from "./vml"
import { customSheetViewsXml, customViewNamesXml, customWorkbookViewsXml, viewNamePrefix } from "./custom-views"
import { hyperlinkTarget } from "./hyperlinks"
import {
  columnName,
  isChart,
  quoteSheet,
  slicerTarget,
  pivotSlicerTarget,
  pivotSlicerItems,
  pivotSlicerOrder,
  type PivotSlicerTarget,
  type PivotCache,
  type SheetTimeline,
  timelineBounds,
  timelineSelection,
  type Range,
  type Sheet,
  type SheetSlicer,
  type Workbook,
} from "@cells/core"
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
  CT_THEME,
  CT_THREADED,
  CT_VML,
  CT_WORKSHEET,
  CT_CHARTSHEET,
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
  REL_HYPERLINK,
  REL_THEME,
  REL_THREADED,
  REL_VML,
  REL_WORKSHEET,
  REL_CHARTSHEET,
} from "./constants"
import { assemble, attrText, kid, raw, type Kept } from "./dom"
import { boxOf, isSignatureLine, putPicture, REL_IMAGE, writeDrawing } from "./drawing"
import { patchThemeXml, themeChanged, themeXml } from "./theme"
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
import {
  CT_CUSTOM_XML_PROPS,
  mashupItem,
  mashupPropsXml,
  mashupQueryTableXml,
  REL_CUSTOM_XML,
  REL_CUSTOM_XML_PROPS,
  sameQueries,
} from "./mashup"
import { printNamesXml } from "./page-setup"
import {
  cacheForWriting,
  CT_PIVOT_CACHE,
  CT_PIVOT_RECORDS,
  CT_PIVOT_TABLE,
  keptPivotChildren,
  movedPivotXml,
  pivotCacheKey,
  pivotCacheXml,
  pivotPlaceKey,
  pivotRecordsXml,
  pivotTableKey,
  pivotTableXml,
  REL_PIVOT_CACHE,
  REL_PIVOT_RECORDS,
  REL_PIVOT_TABLE,
} from "./pivot"
import { formulaToFile } from "./formula"
import type { XlsxSource } from "./read"
import { DEFAULT_VIEW, type SheetView } from "./sheet-read"
import {
  CT_SLICER,
  CT_SLICER_CACHE,
  REL_SLICER,
  REL_SLICER_CACHE,
  slicerCacheName,
  slicerCacheXml,
  slicerPartXml,
  URI_WB_PIVOT_SLICER_CACHES,
  URI_WB_SLICER_CACHES,
  withoutExts,
  workbookPivotSlicersExt,
  workbookSlicersExt,
  pivotSlicerCacheXml,
} from "./slicers"
import {
  CT_TIMELINE,
  CT_TIMELINE_CACHE,
  REL_TIMELINE,
  REL_TIMELINE_CACHE,
  timelineCacheName,
  timelineCacheXml,
  timelinePartXml,
  URI_WB_TIMELINE_CACHES,
  workbookTimelinesExt,
} from "./timelines"
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
  // An unshared legacy shared workbook: its revision headers, logs and user names go.
  if (!wb.legacyShared)
    for (const r of wbRels)
      if (/\/(revisionHeaders|usernames)$/.test(r.type) && !dropped.includes(r)) {
        dropped.push(r)
        for (const log of r.partName ? (relGroups.get(partKey(r.partName)) ?? []) : []) remove(log.partName)
      }
  // Slicers: their caches are written anew (a file with slicers Cells can't read keeps them all as read).
  const ownSlicers = source?.slicersModelled !== false
  const ownTimelines = source?.timelinesModelled !== false
  if (ownTimelines) for (const r of wbRels) if (r.type === REL_TIMELINE_CACHE && !dropped.includes(r)) dropped.push(r)
  if (ownSlicers) for (const r of wbRels) if (r.type === REL_SLICER_CACHE && !dropped.includes(r)) dropped.push(r)
  // Links to other workbooks: Cells writes them anew, unless the file had links of other kinds (kept as read).
  const ownLinks = source?.linksModelled !== false
  if (ownLinks) for (const r of wbRels) if (r.type === REL_EXTERNAL_LINK && !dropped.includes(r)) dropped.push(r)
  const links = ownLinks
    ? new LinkTable(
        wb.links.list,
        wb.sheets.map((s) => s.name),
      )
    : null
  // Power Query: the DataMashup item goes (with its properties) once the workbook has no queries left.
  const mashupPart = source?.mashup?.partName
  if (mashupPart && !wb.queries.length)
    for (const r of wbRels)
      if (
        r.type === REL_CUSTOM_XML &&
        r.partName &&
        partKey(r.partName) === partKey(mashupPart) &&
        !dropped.includes(r)
      ) {
        dropped.push(r)
        for (const p of relGroups.get(partKey(mashupPart)) ?? []) remove(p.partName)
      }
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
  // PivotTable caches Cells models: their relationships are written again below (their parts stay for now, to be
  // written back as read when they didn't change).
  const modelledCaches = new Set([...(source?.pivots?.caches.values() ?? [])].map((c) => partKey(c.partName)))
  relGroups.set(
    partKey(main),
    wbRels.filter(
      (r) =>
        !dropped.includes(r) && !(r.type === REL_PIVOT_CACHE && r.partName && modelledCaches.has(partKey(r.partName))),
    ),
  )

  // Power Query: the queries' DataMashup item, as read while they are the same (Excel's signed permissions stay).
  if (wb.queries.length) {
    const read = source?.mashup
    const data =
      read && sameQueries(wb.queries, read.read, read.loads)
        ? read.read.raw
        : mashupItem(wb.queries, read?.read ?? null)
    if (read) parts.set(partKey(read.partName), { name: read.partName, contentType: "application/xml", data })
    else {
      const item = freeName((n) => `/customXml/item${n}.xml`)
      const n = /item(\d+)\.xml$/.exec(item)![1]
      parts.set(partKey(item), { name: item, contentType: "application/xml", data })
      const props = `/customXml/itemProps${n}.xml`
      taken.add(partKey(props))
      put(props, CT_CUSTOM_XML_PROPS, mashupPropsXml())
      addRel(item, REL_CUSTOM_XML_PROPS, props)
      addRel(main, REL_CUSTOM_XML, item)
    }
  }

  // Page Layout › Themes: the file's theme part with the workbook's colours and fonts, or a new one; the theme's
  // fonts go into the font records that use them.
  const newTheme = themeChanged(wb.theme, source?.theme)
  const styles = new StylesOut(
    source?.styles ?? null,
    source?.reuse ?? new Map(),
    newTheme ? wb.theme.fonts : undefined,
  )
  if (newTheme) {
    const rel = rels(main).find((r) => r.type === REL_THEME && r.partName)
    const old = rel?.partName ? parts.get(partKey(rel.partName)) : undefined
    if (rel?.partName && old) put(rel.partName, CT_THEME, patchThemeXml(new TextDecoder().decode(old.data), wb.theme))
    else {
      const name = freeName((n) => `/xl/theme/theme${n}.xml`)
      put(name, CT_THEME, themeXml(wb.theme))
      addRel(main, REL_THEME, name)
    }
  }
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
  // Loaded queries: each with the table it fills (on some sheet), written with its connection and query table.
  const loadedQueries = wb.queries.filter(
    (q) => q.table && wb.sheets.some((s) => s.tables.some((t) => t.name === q.table && t.query === q.name)),
  )
  const conns = connectionsXml(source?.keptConnections ?? [], textLinks, loadedQueries)
  const queryTableNames: { sheet: number; name: string; range: Range }[] = []
  const usedQueryTables = new Set<string>()

  // One slicer cache per table column with slicers, and per PivotTable field and set of PivotTables; its name is the
  // file's when it can stay.
  type Cache = { name: string; slicer: SheetSlicer } & (
    { kind: "table"; tableId: number; column: string; columnId: number } | { kind: "pivot"; target: PivotSlicerTarget }
  )
  const caches = new Map<string, Cache>()
  const cacheOf = new Map<SheetSlicer, Cache>()
  /** The x14 ids of the PivotTable caches that slicers name (an id of their own for caches that have none yet). */
  const slicerIds = new Map<number, number>()
  const usedIds = new Set(wb.pivotCaches.flatMap((c) => (c.slicerId !== undefined ? [c.slicerId] : [])))
  const needId = (cache: PivotCache) => {
    if (slicerIds.has(cache.id)) return
    let id = cache.slicerId
    if (id === undefined) {
      id = 1000 + cache.id
      while (usedIds.has(id)) id++
      usedIds.add(id)
    }
    slicerIds.set(cache.id, id)
  }
  const takenNames = new Set(wb.names.all().map((n) => n.name.toUpperCase()))
  if (ownSlicers) {
    // The file's slicers parts are written anew.
    for (const p of [...parts.values()]) if (p.contentType === CT_SLICER) remove(p.name)
    const taken = takenNames
    for (const sheet of wb.sheets)
      for (const o of sheet.objects) {
        if (isChart(o) || o.kind !== "slicer") continue
        const nameFor = (column: string) => {
          const name = o.cache && !taken.has(o.cache.toUpperCase()) ? o.cache : slicerCacheName(column, taken)
          taken.add(name.toUpperCase())
          return name
        }
        let key: string
        let make: () => Cache
        if (o.pivot) {
          const target = pivotSlicerTarget(wb.sheets, wb.pivotCaches, o)
          if (!target) continue
          const field = target.cache.fields[target.field].name
          key = `p:${target.cache.id}:${target.field}:${target.tables.map((t) => `${t.sheet}!${t.table.name}`).join("|")}`
          make = () => ({ kind: "pivot", name: nameFor(field), slicer: o, target })
          needId(target.cache)
        } else {
          const target = slicerTarget(wb.sheets, o)
          if (!target) continue
          const index = target.col - target.table.range.c0
          key = `${target.table.id}:${index}`
          const column = target.table.columns[index].name
          make = () => ({
            kind: "table",
            name: nameFor(column),
            slicer: o,
            tableId: target.table.id,
            column,
            columnId: index + 1,
          })
        }
        let cache = caches.get(key)
        if (!cache) caches.set(key, (cache = make()))
        cacheOf.set(o, cache)
      }
  }

  // One timeline cache per date field and set of PivotTables; its name is the file's when it can stay.
  type TimelineCache = { name: string; timeline: SheetTimeline; target: PivotSlicerTarget }
  const timelineCaches = new Map<string, TimelineCache>()
  const timelineCacheOf = new Map<SheetTimeline, TimelineCache>()
  if (ownTimelines) {
    for (const p of [...parts.values()]) if (p.contentType === CT_TIMELINE) remove(p.name)
    for (const sheet of wb.sheets)
      for (const o of sheet.objects) {
        if (isChart(o) || o.kind !== "timeline") continue
        const target = pivotSlicerTarget(wb.sheets, wb.pivotCaches, o)
        if (!target) continue
        const key = `${target.cache.id}:${target.field}:${target.tables.map((t) => `${t.sheet}!${t.table.name}`).join("|")}`
        let cache = timelineCaches.get(key)
        if (!cache) {
          const name =
            o.cache && !takenNames.has(o.cache.toUpperCase())
              ? o.cache
              : timelineCacheName(target.cache.fields[target.field].name, takenNames)
          takenNames.add(name.toUpperCase())
          timelineCaches.set(key, (cache = { name, timeline: o, target }))
          needId(target.cache)
        }
        timelineCacheOf.set(o, cache)
      }
  }

  // PivotTable caches that some PivotTable uses: as read when unchanged, else written anew with their records.
  const numFmt = (code?: string) => styles.numberFormatId(code)
  const cacheEntries: string[] = []
  const cachePart = new Map<number, string>()
  /** Caches written anew, as written (every axis field's shared items listed). */
  const cacheWritten = new Map<number, ReturnType<typeof cacheForWriting>>()
  for (let cache of wb.pivotCaches) {
    const tables = wb.sheets.flatMap((s) => s.pivots.filter((t) => t.cacheId === cache.id))
    if (!tables.length || cachePart.has(cache.id)) continue
    const was = source?.pivots?.caches.get(cache.id)
    const slicerId = slicerIds.get(cache.id)
    if (slicerId !== undefined && cache.slicerId !== slicerId) cache = { ...cache, slicerId }
    const out = cacheForWriting(cache, tables)
    const same =
      was !== undefined &&
      was.key === pivotCacheKey(cache) &&
      was.records === cache.records &&
      parts.has(partKey(was.partName)) &&
      out.fields.every((f, i) => f === cache.fields[i])
    const def = was?.partName ?? freeName((n) => `/xl/pivotCache/pivotCacheDefinition${n}.xml`)
    taken.add(partKey(def))
    if (!same) {
      relGroups.set(partKey(def), [])
      // Save source data with file off: the definition alone, no records part.
      if (out.saveData === false) put(def, CT_PIVOT_CACHE, pivotCacheXml(out, "", numFmt))
      else {
        const records = was?.recordsPart ?? freeName((n) => `/xl/pivotCache/pivotCacheRecords${n}.xml`)
        taken.add(partKey(records))
        const recordsRel = addRel(def, REL_PIVOT_RECORDS, records)
        put(def, CT_PIVOT_CACHE, pivotCacheXml(out, recordsRel, numFmt))
        put(records, CT_PIVOT_RECORDS, pivotRecordsXml(out))
      }
      cacheWritten.set(cache.id, out)
    }
    cachePart.set(cache.id, def)
    cacheEntries.push(`<pivotCache cacheId="${cache.id}" r:id="${addRel(main, REL_PIVOT_CACHE, def)}"/>`)
  }

  wb.sheets.forEach((sheet, i) => {
    if (sheet.chartSheet) {
      // A chart sheet: its view, margins and the drawing with its one chart.
      const cs = source?.chartSheets?.get(sheet)
      const partName = cs?.partName ?? freeName((n) => `/xl/chartsheets/sheet${n}.xml`)
      taken.add(partKey(partName))
      relGroups.set(partKey(partName), [])
      const drawing = writeDrawing(
        wb,
        sheet,
        partName,
        cs?.drawing,
        {
          put,
          removeTree,
          addRel,
          rels,
          freeName,
          putBytes: (name, contentType, data) => parts.set(partKey(name), { name, contentType, data }),
          fileName: fileName.split(/[\\/]/).pop(),
        },
        true,
      )
      const view = views[i] ?? DEFAULT_VIEW
      put(partName, CT_CHARTSHEET, chartSheetXml(sheet, i === active || view.selected, view.zoom, drawing))
      const rid = addRel(main, REL_CHARTSHEET, partName)
      sheetEntries.push(
        `<sheet${attrText({ name: sheet.name, sheetId: i + 1, state: sheet.hidden ? "hidden" : undefined })} r:id="${rid}"/>`,
      )
      return
    }
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
    // Signature lines and embedded objects share the notes' VML drawing (Excel keeps them there).
    const lines = sheet.objects.filter(isSignatureLine)
    const embedded = sheet.objects.filter(isEmbeddedObject)
    const oleItems: string[] = []
    if (notes || lines.length || embedded.length) {
      const vml = freeName((n) => `/xl/drawings/vmlDrawing${n}.vml`)
      const putBytes = (name: string, contentType: string, data: Uint8Array) =>
        parts.set(partKey(name), { name, contentType, data })
      let shapeXml = ""
      let shapeId = Math.max(1024 * (i + 1), notes ? lastShapeId(notes.vml) : 0)
      for (const line of lines) {
        const rel = addRel(vml, REL_IMAGE, putPicture(line.src, { freeName, putBytes }))
        shapeXml += signatureShapeXml(line, ++shapeId, rel, boxOf(sheet, line.anchor))
      }
      for (const o of embedded) {
        const id = ++shapeId
        const { relId, progId } = writeEmbedding(o.embed, {
          embed: (name, contentType, relType, data) => {
            const part = freeName(name)
            putBytes(part, contentType, data)
            return addRel(partName, relType, part)
          },
        })
        const media = putPicture(o.src, { freeName, putBytes })
        const imageRelId = addRel(partName, REL_IMAGE, media)
        shapeXml += oleShapeXml(o, id, addRel(vml, REL_IMAGE, media), boxOf(sheet, o.anchor))
        oleItems.push(oleObjectXml(o, { shapeId: id, relId, imageRelId, progId }))
      }
      const pictures = lines.length + embedded.length > 0
      const xml = notes ? (pictures ? withPictures(notes.vml, shapeXml) : notes.vml) : pictureVml(i, shapeXml)
      parts.set(partKey(vml), { name: vml, contentType: CT_VML, data: new TextEncoder().encode(xml) })
      legacyDrawing = addRel(partName, REL_VML, vml)
    }
    if (notes) {
      const comments = freeName((n) => `/xl/comments${n}.xml`)
      put(comments, CT_COMMENTS, notes.legacy)
      addRel(partName, REL_COMMENTS, comments)
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
      const connectionId = t.query ? conns?.queryIds.get(t.query) : undefined
      put(part, CT_TABLE, tableXml(t, styles, rangeRef, toFile, connectionId !== undefined))
      tableParts.push(addRel(partName, REL_TABLE, part))
      relGroups.set(partKey(part), [])
      if (connectionId !== undefined) {
        // The query table: named ExternalData_n (a hidden name places it), its fields the table's columns.
        const qtPart = freeName((n) => `/xl/queryTables/queryTable${n}.xml`)
        usedQueryTables.add(partKey(qtPart))
        const name = `ExternalData_${queryTableNames.length + 1}`
        queryTableNames.push({ sheet: wb.sheets.indexOf(sheet), name, range: t.range })
        put(
          qtPart,
          CT_QUERY_TABLE,
          mashupQueryTableXml(
            name,
            connectionId,
            t.columns.map((c, i) => ({ id: i + 1, name: c.name })),
          ),
        )
        addRel(part, REL_QUERY_TABLE, qtPart)
      }
    }
    // PivotTables: as read when neither they nor their cache changed, moved when only their place did, else anew.
    const pivotsRead = source?.pivots?.tables.get(sheet)
    for (const t of sheet.pivots) {
      const def = cachePart.get(t.cacheId)
      const cache = wb.pivotCaches.find((c) => c.id === t.cacheId)
      if (!def || !cache) continue
      const was = t.part ? pivotsRead?.get(t.part) : undefined
      const rewritten = cacheWritten.get(t.cacheId)
      const part =
        was && !taken.has(partKey(was.partName)) ? was.partName : freeName((n) => `/xl/pivotTables/pivotTable${n}.xml`)
      taken.add(partKey(part))
      if (was && part === was.partName && !rewritten && was.key === pivotTableKey(t) && parts.has(partKey(part))) {
        if (was.place !== pivotPlaceKey(t)) put(part, CT_PIVOT_TABLE, movedPivotXml(was.xml, t))
      } else
        put(
          part,
          CT_PIVOT_TABLE,
          pivotTableXml(t, rewritten ?? cacheForWriting(cache, [t]), numFmt, was ? keptPivotChildren(was.xml) : []),
        )
      relGroups.set(partKey(part), [])
      addRel(part, REL_PIVOT_CACHE, def)
      addRel(partName, REL_PIVOT_TABLE, part)
    }
    for (const r of own) if (r.type === REL_QUERY_TABLE && r.partName) usedQueryTables.add(partKey(r.partName))
    for (const q of sheet.queryTables) {
      const part = freeName((n) => `/xl/queryTables/queryTable${n}.xml`)
      usedQueryTables.add(partKey(part))
      put(part, CT_QUERY_TABLE, queryTableXml(q, conns!.ids.get(q)!))
      addRel(partName, REL_QUERY_TABLE, part)
    }
    // Links out of the workbook: their addresses go in the sheet's relationships.
    const hyperlinkRels = sheet.hyperlinks.map((l) => {
      if (!l.address) return null
      const id = newId(own)
      own.push({ id, type: REL_HYPERLINK, target: hyperlinkTarget(l.address), targetMode: "External", partName: null })
      return id
    })
    const drawing = writeDrawing(wb, sheet, partName, s?.drawing, {
      put,
      removeTree,
      addRel,
      rels,
      freeName,
      putBytes: (name, contentType, data) => parts.set(partKey(name), { name, contentType, data }),
      fileName: fileName.split(/[\\/]/).pop(),
    })
    // Page Layout › Background: its picture in a media part of its own.
    const picture = sheet.background
      ? addRel(
          partName,
          REL_IMAGE,
          putPicture(sheet.background, {
            freeName,
            putBytes: (name, contentType, data) => parts.set(partKey(name), { name, contentType, data }),
          }),
        )
      : null
    // The sheet's table slicers, in their own part.
    let slicers: { relId: string | null; pivotRelId: string | null } | undefined
    if (ownSlicers) {
      const list = sheet.objects.flatMap((o) =>
        !isChart(o) && o.kind === "slicer" && cacheOf.has(o) ? [{ slicer: o, cache: cacheOf.get(o)! }] : [],
      )
      const written: { relId: string | null; pivotRelId: string | null } = { relId: null, pivotRelId: null }
      slicers = written
      // Table slicers and PivotTable slicers each have a part of their own.
      for (const kind of ["pivot", "table"] as const) {
        const own = list.filter((x) => x.cache.kind === kind).map((x) => ({ slicer: x.slicer, cache: x.cache.name }))
        if (!own.length) continue
        const part = freeName((n) => `/xl/slicers/slicer${n}.xml`)
        put(part, CT_SLICER, slicerPartXml(own))
        const relId = addRel(partName, REL_SLICER, part)
        if (kind === "pivot") written.pivotRelId = relId
        else written.relId = relId
      }
    }
    // The sheet's timelines, in their own part.
    let timelines: { relId: string | null } | undefined
    if (ownTimelines) {
      const list = sheet.objects.flatMap((o) =>
        !isChart(o) && o.kind === "timeline" && timelineCacheOf.has(o)
          ? [{ timeline: o, cache: timelineCacheOf.get(o)!.name }]
          : [],
      )
      timelines = { relId: null }
      if (list.length) {
        const part = freeName((n) => `/xl/timelines/timeline${n}.xml`)
        put(part, CT_TIMELINE, timelinePartXml(list))
        timelines.relId = addRel(partName, REL_TIMELINE, part)
      }
    }
    const view = { ...(views[i] ?? DEFAULT_VIEW), selected: i === active || (views[i]?.selected ?? false) }
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
      oleObjects: oleObjectsXml(oleItems),
      drawing,
      tableParts,
      links,
      hyperlinkRels,
      picture,
      customSheetViews: customSheetViewsXml(wb, sheet),
      slicers,
      timelines,
    })
    dynamic ||= written.dynamic
    put(partName, CT_WORKSHEET, written.xml)
    const rid = addRel(main, REL_WORKSHEET, partName)
    sheetEntries.push(
      `<sheet${attrText({ name: sheet.name, sheetId: i + 1, state: sheet.hidden ? (s?.state === "veryHidden" ? "veryHidden" : "hidden") : undefined })} r:id="${rid}"/>`,
    )
  })

  const cacheRels: string[] = []
  const pivotCacheRels: string[] = []
  for (const cache of caches.values()) {
    const part = freeName((n) => `/xl/slicerCaches/slicerCache${n}.xml`)
    if (cache.kind === "pivot") {
      const { target, slicer } = cache
      const shown = new Map(pivotSlicerItems(target, { ...slicer, hideNoData: false }).map((i) => [i.key, i]))
      put(
        part,
        CT_SLICER_CACHE,
        pivotSlicerCacheXml({
          name: cache.name,
          field: target.cache.fields[target.field].name,
          slicer,
          pivotCacheId: slicerIds.get(target.cache.id)!,
          tables: target.tables.map((t) => ({ tabId: t.sheet + 1, name: t.table.name })),
          items: pivotSlicerOrder(target.cache, target.field, slicer.sort === "descending").map((x) => ({
            x,
            selected: shown.get(`x:${x}`)?.selected ?? true,
            data: shown.get(`x:${x}`)?.data ?? true,
          })),
        }),
      )
      pivotCacheRels.push(addRel(main, REL_SLICER_CACHE, part))
    } else {
      put(part, CT_SLICER_CACHE, slicerCacheXml(cache))
      cacheRels.push(addRel(main, REL_SLICER_CACHE, part))
    }
  }

  const timelineCacheRels: string[] = []
  for (const cache of timelineCaches.values()) {
    const { target, timeline } = cache
    const part = freeName((n) => `/xl/timelineCaches/timelineCache${n}.xml`)
    put(
      part,
      CT_TIMELINE_CACHE,
      timelineCacheXml({
        name: cache.name,
        field: target.cache.fields[target.field].name,
        uid: timeline.cacheUid,
        pivotCacheId: slicerIds.get(target.cache.id)!,
        tables: target.tables.map((t) => ({ tabId: t.sheet + 1, name: t.table.name })),
        picked: timelineSelection(target),
        bounds: timelineBounds(target.cache, target.field),
      }),
    )
    timelineCacheRels.push(addRel(main, REL_TIMELINE_CACHE, part))
  }

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
  const namesXml = definedNamesXml(
    wb,
    source,
    links,
    [...caches.values(), ...timelineCaches.values()].map((c) => c.name),
    queryTableNames,
  )
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
    customWorkbookViews: customWorkbookViewsXml(
      wb,
      wb.sheets.map((_, i) => i + 1),
    ),
    workbookProtection: protectionXml(wb),
    bookViews: bookViewsXml(source, active),
    ...(source?.pivots || cacheEntries.length
      ? {
          pivotCaches:
            cacheEntries.length || source?.pivots?.kept.length
              ? `<pivotCaches>${[...cacheEntries, ...(source?.pivots?.kept ?? [])].join("")}</pivotCaches>`
              : "",
        }
      : {}),
    calcPr: `<calcPr${attrText({ calcId: 191029, calcMode: wb.calc.mode === "manual" ? "manual" : undefined })}/>`,
  }
  let wbXml: string
  if (source) {
    const doc = source.workbook
    const kept: Kept[] = doc.root.children
      .filter((c) => !(c.local in generated))
      .map((c) => ({ local: c.local, xml: raw(doc, c) }))
    if (ownSlicers || ownTimelines) {
      const ext = kept.find((k) => k.local === "extLst")
      const uris = [
        ...(ownSlicers ? [URI_WB_SLICER_CACHES, URI_WB_PIVOT_SLICER_CACHES] : []),
        ...(ownTimelines ? [URI_WB_TIMELINE_CACHES] : []),
      ]
      const rest = ext ? withoutExts(ext.xml, uris) : ""
      const add =
        workbookPivotSlicersExt(pivotCacheRels) +
        workbookSlicersExt(cacheRels) +
        workbookTimelinesExt(timelineCacheRels)
      const xml = add ? withExt(rest || undefined, add) : rest
      if (ext) ext.xml = xml
      else if (xml) kept.push({ local: "extLst", xml })
    }
    // A theme of its own: Excel reads defaultThemeVersion as "Office's theme" and would show that instead.
    if (newTheme)
      for (const k of kept) if (k.local === "workbookPr") k.xml = k.xml.replace(/\s+defaultThemeVersion="[^"]*"/, "")
    const open = doc.source.slice(doc.root.start, doc.root.tagEnd).replace(/\/?>$/, "")
    wbXml = `${open}>${assemble(WORKBOOK_ORDER, kept, generated)}</${doc.root.name}>`
  } else {
    wbXml =
      `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
      assemble(
        WORKBOOK_ORDER,
        [
          { local: "workbookPr", xml: newTheme ? `<workbookPr/>` : `<workbookPr defaultThemeVersion="202300"/>` },
          ...(cacheRels.length || pivotCacheRels.length || timelineCacheRels.length
            ? [
                {
                  local: "extLst",
                  xml: withExt(
                    undefined,
                    workbookPivotSlicersExt(pivotCacheRels) +
                      workbookSlicersExt(cacheRels) +
                      workbookTimelinesExt(timelineCacheRels),
                  ),
                },
              ]
            : []),
        ],
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

function definedNamesXml(
  wb: Workbook,
  source: XlsxSource | null,
  links: LinkTable | null,
  slicerCaches: readonly string[] = [],
  queryTableNames: readonly { sheet: number; name: string; range: Range }[] = [],
): string {
  const items: string[] = []
  // Each slicer cache has a name of its own, #N/A.
  for (const name of slicerCaches) items.push(`<definedName name="${name}">#N/A</definedName>`)
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
  for (const q of queryTableNames) {
    const r = q.range
    const abs = (row: number, col: number) => "$" + columnName(col) + "$" + (row + 1)
    const ref = quoteSheet(wb.sheets[q.sheet].name) + "!" + abs(r.r0, r.c0) + ":" + abs(r.r1, r.c1)
    items.push(`<definedName name="${q.name}" localSheetId="${q.sheet}" hidden="1">${escapeXmlText(ref)}</definedName>`)
  }
  // Custom views' hidden rows, columns, print areas and titles.
  items.push(...customViewNamesXml(wb))
  const viewPrefixes = wb.customViews.map((v) => viewNamePrefix(v.guid))
  for (const k of source?.keptNames ?? []) {
    // A deleted custom view's other names (its filter) go with it.
    const keptName = /\bname="([^"]*)"/.exec(k.xml)?.[1] ?? ""
    if (/^Z_[0-9A-F_]+_\.wvu\./i.test(keptName) && !viewPrefixes.some((p) => keptName.startsWith(p))) continue
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

/** A chart sheet's part: its view (zoomed to fit the window, as Excel opens one), margins, paper and drawing. */
function chartSheetXml(sheet: Sheet, selected: boolean, zoom: number, drawing: string | null): string {
  const m = sheet.pageSetup.margins
  const zoomAttr = zoom && zoom !== 1 ? ` zoomScale="${Math.round(zoom * 100)}"` : ""
  return (
    `<chartsheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetPr/>` +
    `<sheetViews><sheetView${selected ? ` tabSelected="1"` : ""}${zoomAttr} workbookViewId="0" zoomToFit="1"/></sheetViews>` +
    `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>` +
    `<pageSetup orientation="${sheet.pageSetup.orientation}"/>` +
    (drawing ? `<drawing r:id="${drawing}"/>` : "") +
    `</chartsheet>`
  )
}
