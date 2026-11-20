import { readCustomViews, type CustomSheetViewRead } from "./custom-views"
import {
  copyTheme,
  tableAt,
  Workbook,
  type DefinedName,
  type Sheet,
  type Theme,
  type PivotCache,
  type PivotTable,
  type PivotValue,
  type WorkbookProtection,
  isChart,
  objectName,
} from "@cells/core"
import { decodeXml, parseXmlEvents } from "@cells/ooxml"
import { findPart, openPackage, relationshipsFor, type OpcPackage, type OpcRelationship } from "../opc"
import { readTheme, readThemeColors } from "./theme"
import { readComments, readPersons } from "./comments"
import {
  NS_R,
  REL_COMMENTS,
  REL_PERSON,
  REL_SHARED_STRINGS,
  REL_STYLES,
  REL_THEME,
  REL_TABLE,
  REL_HYPERLINK,
  REL_THREADED,
  REL_VML,
  REL_WORKSHEET,
} from "./constants"
import { at, kid, kids, parseTree, raw, type Kept, type XDoc } from "./dom"
import { readVmlSignatures } from "./vml"
import { embedFrom, readOleObjects, REL_OLE_OBJECT, REL_PACKAGE, vmlAnchor, vmlImageRel } from "./ole-objects"
import { pictureUrl, readDrawing, REL_DRAWING, REL_IMAGE, type DrawingSource } from "./drawing"
import { linksFromFile, readExternalLink } from "./links"
import {
  queryTableFrom,
  readConnections,
  readQueryTable,
  REL_CONNECTIONS,
  REL_QUERY_TABLE,
  type ConnectionIn,
} from "./connections"
import { readPrintArea, readPrintTitles, type PageSetupExtras } from "./page-setup"
import { formulaFromFile } from "./formula"
import { hyperlinkAddress } from "./hyperlinks"
import {
  pivotCacheKey,
  pivotPlaceKey,
  pivotTableKey,
  readPivotCache,
  readPivotTable,
  REL_PIVOT_CACHE,
  REL_PIVOT_RECORDS,
  REL_PIVOT_TABLE,
} from "./pivot"
import { readSlicerCache, readSlicerPart, REL_SLICER, REL_SLICER_CACHE, type SlicerCacheRead } from "./slicers"
import { readPasswordHash, readSheet, unescapeText, DEFAULT_VIEW, type SheetView } from "./sheet-read"
import { StylesIn } from "./styles"
import { readTable, structuredFromFile } from "./table"

/** What a saved file needs from the file it was opened from, to write back what Cells didn't read. */
export type SheetSource = {
  partName: string
  rootTag: string
  kept: Kept[]
  keptConditional: string[]
  rowExtras: Map<number, string>
  colExtras: Map<number, string>
  /** "veryHidden" sheets stay very hidden. */
  state: string
  /** The sheet's relationships other than the ones Cells writes (comments, VML, threaded comments, tables). */
  rels: OpcRelationship[]
  /** Each table's part, by table id, so a save writes it back in the same place. */
  tables: { id: number; partName: string }[]
  /** The sheet's drawing: its charts' parts and the anchors Cells keeps as they came. */
  drawing?: DrawingSource
  /** Print settings Cells doesn't model, written back as read. */
  pageExtras?: PageSetupExtras
}

export type XlsxSource = {
  pkg: OpcPackage
  workbook: XDoc
  styles: StylesIn | null
  /** Cells style id → the file's xf index, for formats read and not changed. */
  reuse: Map<number, number>
  /** Plain text → the shared string as read (with its runs), so rich text survives a save. */
  rich: Map<string, string>
  /** Defined names Cells doesn't manage (sheet-scoped, hidden, print areas), as read, with the sheet they belong to. */
  keptNames: { xml: string; sheet: Sheet | null }[]
  sheets: WeakMap<Sheet, SheetSource>
  /** The workbook had dynamic-array formulas (needs metadata.xml on save). */
  dynamic: boolean
  /** Connections Cells doesn't model (databases, the web, Power Query), written back as read. */
  keptConnections: ConnectionIn[]
  /** The file's links to other workbooks are in wb.links (false: some are of other kinds, kept as read). */
  linksModelled: boolean
  /** The file's slicers are table slicers Cells reads and writes anew (false: kept as read). */
  slicersModelled?: boolean
  /** The theme as read, so a save rewrites the theme part only when it changed. */
  theme?: Theme
  /** PivotTables as read, so a save writes back the parts of those that didn't change. */
  pivots?: PivotSource
}

export type PivotSource = {
  /** Modelled caches by id: their parts, and what they were when read. */
  caches: Map<number, { partName: string; recordsPart: string | null; key: string; records: PivotValue[][] }>
  /** Modelled PivotTables by sheet and the part each was read from (PivotTable.part). */
  tables: WeakMap<Sheet, Map<string, { partName: string; xml: string; key: string; place: string }>>
  /** The workbook's `<pivotCache>` entries for caches kept as read (their relationships stay). */
  kept: string[]
}

export type XlsxRead = {
  workbook: Workbook
  views: SheetView[]
  activeSheet: number
  source: XlsxSource
  warnings: string[]
}

export class XlsxError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "XlsxError"
  }
}

/** Whether any formula of the sheet may hold a structured reference (to another sheet's table). */
function hasBracket(sheet: Sheet): boolean {
  let found = false
  sheet.cells.forEachFormula((_r, _c, input) => {
    if (input.includes("[")) found = true
  })
  return found
}

const relOfType = (rels: OpcRelationship[], type: string) => rels.find((r) => r.type === type && r.partName)

function partText(pkg: OpcPackage, name: string | null | undefined): string | undefined {
  if (!name) return undefined
  const part = findPart(pkg, name)
  return part ? decodeXml(part.data) : undefined
}

/** The shared strings, streamed: plain text of each `si` (runs joined, phonetic runs skipped), and the rich ones raw. */
function readSharedStrings(xml: string, rich: Map<string, string>): string[] {
  const out: string[] = []
  let depth = 0
  let siDepth = -1
  let siStart = 0
  let runs = false
  let inT = false
  let skip = 0
  let text = ""
  parseXmlEvents(xml, {
    start(tag) {
      depth++
      if (tag.local === "si") {
        siDepth = depth
        siStart = tag.start
        text = ""
        runs = false
      } else if (siDepth > 0) {
        if (tag.local === "rPh" || tag.local === "phoneticPr") skip++
        else if (tag.local === "r") runs = true
        else if (tag.local === "t" && !skip) inT = true
      }
    },
    end(tag) {
      if (tag.local === "si" && depth === siDepth) {
        const plain = unescapeText(text)
        out.push(plain)
        if (runs && !rich.has(plain)) rich.set(plain, xml.slice(siStart, tag.end))
        siDepth = -1
      } else if (tag.local === "rPh" || tag.local === "phoneticPr") skip--
      else if (tag.local === "t") inT = false
      depth--
    },
    text(t) {
      if (inT) text += t
    },
  })
  return out
}

export type OpenOptions = {
  /**
   * Called now and then with the share of the file read (0–1). It is awaited: a UI passes a function that waits for
   * the next frame, so the window stays responsive and can draw a progress bar while a big file opens.
   */
  progress?: (done: number) => void | Promise<void>
  /** Skip calculating after loading (a worker that calculates on its own does it). */
  calculate?: boolean
}

/** Opens an .xlsx (or .xlsm, .xltx, .xltm) file. */
export async function openXlsx(
  bytes: Uint8Array,
  fileName = "Book1.xlsx",
  options: OpenOptions = {},
): Promise<XlsxRead> {
  const pkg = await openPackage(bytes, fileName)
  await options.progress?.(0.05)
  const warnings: string[] = pkg.warnings.filter((w) => w.code === "broken-part").map((w) => w.message)
  const wbXml = partText(pkg, pkg.mainPartName)
  if (!wbXml) throw new XlsxError("The file has no workbook part.")
  const wbDoc = parseTree(wbXml)
  const root = wbDoc.root
  const wbRels = relationshipsFor(pkg, pkg.mainPartName)

  const themeXml = partText(pkg, relOfType(wbRels, REL_THEME)?.partName)
  const theme = readThemeColors(themeXml)
  const stylesXml = partText(pkg, relOfType(wbRels, REL_STYLES)?.partName)
  const styles = stylesXml ? new StylesIn(stylesXml, theme) : null
  const rich = new Map<string, string>()
  const sstXml = partText(pkg, relOfType(wbRels, REL_SHARED_STRINGS)?.partName)
  const strings = sstXml ? readSharedStrings(sstXml, rich) : []
  const personsXml = partText(pkg, relOfType(wbRels, REL_PERSON)?.partName)
  const connections = readConnections(partText(pkg, relOfType(wbRels, REL_CONNECTIONS)?.partName) ?? null)
  /** Query tables found on sheets, placed once the defined names are read. */
  const linked: { sheet: Sheet; relId: string; name: string; connectionId: number }[] = []
  const persons = personsXml ? readPersons(personsXml) : new Map<string, string>()

  if (at(kid(root, "workbookPr"), "date1904") === "1" || at(kid(root, "workbookPr"), "date1904") === "true")
    warnings.push("The workbook uses the 1904 date system; Cells shows its dates from 1900, 4 years and 1 day earlier.")

  const sheetNodes = kids(kid(root, "sheets"), "sheet")
  if (!sheetNodes.length) throw new XlsxError("The workbook has no sheets.")
  const wb = new Workbook(sheetNodes.map((s) => at(s, "name") ?? "Sheet"))
  wb.theme = readTheme(themeXml, theme)

  const reuse = new Map<number, number>()
  const xfIds = new Map<number, number>()
  const styleId = (xf: number) => {
    let id = xfIds.get(xf)
    if (id !== undefined) return id
    id = styles && xf > 0 && xf < styles.cellXfs.length ? wb.styles.intern(styles.xfStyle(xf)) : 0
    if (id !== 0 && !reuse.has(id)) reuse.set(id, xf)
    xfIds.set(xf, id)
    return id
  }

  // Slicers: their caches (a table column, or a PivotTable cache field); a file with any Cells can't read (OLAP)
  // keeps all slicers as read.
  const cacheRels = wbRels.filter((r) => r.type === REL_SLICER_CACHE && r.partName)
  const slicerCaches = new Map<string, SlicerCacheRead>()
  let slicersModelled = true
  for (const r of cacheRels) {
    const c = readSlicerCache(partText(pkg, r.partName) ?? "")
    if (c) slicerCaches.set(c.name, c)
    else slicersModelled = false
  }
  /** Each table's column ids (slicer caches name a column by its id). */
  const tableColumnIds = new Map<number, number[]>()
  const slicerSheets = new Map<Sheet, number>()

  // PivotTables: the caches the workbook lists; each sheet's PivotTables are read with the sheet.
  type CacheRead = {
    id: number
    cache: PivotCache | null
    partName: string
    recordsPart: string | null
    entry: string
    relId: string
  }
  const pivotCaches: CacheRead[] = []
  for (const n of kids(kid(root, "pivotCaches"), "pivotCache")) {
    const relId = at(n, "id", NS_R) ?? ""
    const r = wbRels.find((x) => x.id === relId && x.type === REL_PIVOT_CACHE && x.partName)
    if (!r?.partName) continue
    const id = Number(at(n, "cacheId") ?? 0)
    const recordsPart = relOfType(relationshipsFor(pkg, r.partName), REL_PIVOT_RECORDS)?.partName ?? null
    const defXml = partText(pkg, r.partName) ?? ""
    let cache: PivotCache | null
    try {
      cache = readPivotCache(id, defXml, partText(pkg, recordsPart), (f) => styles?.numberFormat(f))
    } catch {
      cache = null
    }
    pivotCaches.push({ id, cache, partName: r.partName, recordsPart, entry: raw(wbDoc, n), relId })
  }
  type TableRead = {
    sheet: Sheet
    relId: string
    partName: string
    xml: string
    table: PivotTable | null
    cache: string
  }
  const pivotTables: TableRead[] = []

  const views: SheetView[] = []
  const customSheetViews = new Map<Sheet, CustomSheetViewRead[]>()
  const sheetSources = new WeakMap<Sheet, SheetSource>()
  let dynamic = false
  // Progress: by the size of the sheet parts, which hold nearly all of a big file.
  const sizes = sheetNodes.map((node) => {
    const rel = wbRels.find((r) => r.id === at(node, "id", NS_R))
    return rel?.partName ? (findPart(pkg, rel.partName)?.data.length ?? 0) : 0
  })
  const total = sizes.reduce((a, b) => a + b, 0) || 1
  let before = 0
  let chartIds = 0
  for (const [i, node] of sheetNodes.entries()) {
    const sheet = wb.sheets[i]
    const state = at(node, "state") ?? "visible"
    sheet.hidden = state !== "visible"
    const rid = at(node, "id", NS_R)
    const rel = wbRels.find((r) => r.id === rid)
    if (!rel?.partName || rel.type !== REL_WORKSHEET) {
      warnings.push(`${sheet.name} is a chart sheet or dialog sheet; Cells opens it as an empty sheet.`)
      views.push({ ...DEFAULT_VIEW })
      before += sizes[i]
      continue
    }
    const xml = partText(pkg, rel.partName) ?? ""
    const share = sizes[i] / total
    const start = before / total
    const read = await readSheet(xml, {
      sheet,
      styles,
      styleId,
      strings,
      warnings,
      progress: options.progress ? (done) => options.progress!(0.1 + 0.85 * (start + share * done)) : undefined,
    })
    dynamic ||= read.dynamic
    if (read.customViews.length) customSheetViews.set(sheet, read.customViews)
    const rels = relationshipsFor(pkg, rel.partName)
    const vmlRel = rels.find((r) => r.id === read.legacyDrawing)
    const vml = partText(pkg, vmlRel?.partName)
    // Pictures (signature lines, embedded objects) are read; other VML objects are form controls.
    if (vml && /ObjectType="(?!Note"|Pict")/.test(vml))
      warnings.push(`Form controls on ${sheet.name} aren't supported; they are left out when you save.`)
    readComments(sheet, {
      legacy: partText(pkg, relOfType(rels, REL_COMMENTS)?.partName),
      vml,
      threaded: partText(pkg, relOfType(rels, REL_THREADED)?.partName),
      persons,
    })
    for (const h of read.hyperlinks) {
      if (h.relId) {
        const target = rels.find((r) => r.id === h.relId && r.type === REL_HYPERLINK)?.target
        if (target) h.link.address = hyperlinkAddress(target)
      }
      if (h.link.address || h.link.location) sheet.hyperlinks.push(h.link)
    }
    // Page Layout › Background.
    const pictureRel = rels.find((r) => r.id === read.picture && r.type === REL_IMAGE)
    const picturePart = pictureRel?.partName ? findPart(pkg, pictureRel.partName) : undefined
    if (picturePart) sheet.background = pictureUrl(pictureRel!.partName!, picturePart)
    const tables: { id: number; partName: string }[] = []
    for (const r of rels) {
      if (r.type !== REL_TABLE || !r.partName) continue
      const tableXml = partText(pkg, r.partName) ?? ""
      const table = readTable(tableXml, styles, warnings)
      if (!table) {
        warnings.push(`A table on ${sheet.name} couldn't be read; it was left out.`)
        continue
      }
      sheet.tables.push(table)
      tables.push({ id: table.id, partName: r.partName })
      if (cacheRels.length)
        tableColumnIds.set(
          table.id,
          [...tableXml.matchAll(/<(?:\w+:)?tableColumn\b[^>]*?\bid="(\d+)"/g)].map((m) => Number(m[1])),
        )
    }
    for (const r of rels) {
      if (r.type !== REL_PIVOT_TABLE || !r.partName || !pivotCaches.length) continue
      const xml = partText(pkg, r.partName) ?? ""
      const cache = relOfType(relationshipsFor(pkg, r.partName), REL_PIVOT_CACHE)?.partName ?? ""
      let table: PivotTable | null
      try {
        table = readPivotTable(xml, (f) => styles?.numberFormat(f))
      } catch {
        table = null
      }
      pivotTables.push({ sheet, relId: r.id, partName: r.partName, xml, table, cache })
    }
    for (const r of rels) {
      if (r.type !== REL_QUERY_TABLE || !r.partName) continue
      const q = readQueryTable(partText(pkg, r.partName) ?? null)
      if (q && connections.get(q.connectionId)?.text) linked.push({ sheet, relId: r.id, ...q })
    }
    const slicerDefs =
      slicersModelled && cacheRels.length
        ? new Map(
            rels
              .filter((r) => r.type === REL_SLICER && r.partName)
              .flatMap((r) => readSlicerPart(partText(pkg, r.partName) ?? ""))
              .map((d) => [d.name, d] as const),
          )
        : undefined
    if (slicerDefs?.size) slicerSheets.set(sheet, i)
    const drawingRel = rels.find((r) => r.id === read.drawing && r.type === REL_DRAWING)
    const drawingXml = partText(pkg, drawingRel?.partName)
    const drawing =
      drawingRel?.partName && drawingXml
        ? readDrawing(
            drawingRel.partName,
            drawingXml,
            relationshipsFor(pkg, drawingRel.partName),
            sheet,
            (name) => partText(pkg, name),
            () => ++chartIds,
            warnings,
            (name) => {
              const part = name ? findPart(pkg, name) : undefined
              return part ? { data: part.data, contentType: part.contentType } : undefined
            },
            slicerDefs,
          )
        : undefined
    // Signature lines (Insert › Signature Line) live in the notes' VML drawing.
    if (vml && vmlRel?.partName) {
      const vmlRels = relationshipsFor(pkg, vmlRel.partName)
      for (const line of readVmlSignatures(vml)) {
        const target = vmlRels.find((r) => r.id === line.relId)?.partName
        const part = target ? findPart(pkg, target) : undefined
        const id = ++chartIds
        const { relId: _rel, spid: _spid, ...picture } = line
        sheet.objects.push({
          ...picture,
          id,
          name: objectName({ kind: "picture" }, id),
          src: part ? pictureUrl(target!, part) : "",
        })
      }
    }
    // Insert › Object's embedded files: the sheet's oleObjects, their pictures and the embedded parts.
    const oleImageRels = new Set<string>()
    if (read.oleObjects) {
      for (const ole of readOleObjects(read.oleObjects)) {
        const anchor = ole.anchor ?? vmlAnchor(vml, ole.shapeId)
        if (!anchor) continue
        const embedRel = rels.find((r) => r.id === ole.relId)
        const embedPart = embedRel?.partName ? findPart(pkg, embedRel.partName) : undefined
        let src = ""
        const imageRel = ole.imageRelId ? rels.find((r) => r.id === ole.imageRelId) : undefined
        if (imageRel?.partName) {
          oleImageRels.add(imageRel.id)
          const part = findPart(pkg, imageRel.partName)
          if (part) src = pictureUrl(imageRel.partName, part)
        } else if (vmlRel?.partName) {
          const target = relationshipsFor(pkg, vmlRel.partName).find((r) => r.id === vmlImageRel(vml, ole.shapeId))
          const part = target?.partName ? findPart(pkg, target.partName) : undefined
          if (part) src = pictureUrl(target!.partName!, part)
        }
        const id = ++chartIds
        sheet.objects.push({
          kind: "picture",
          id,
          name: `Object ${id}`,
          anchor,
          src,
          embed: embedFrom(
            ole.progId,
            embedPart
              ? { name: embedRel!.partName!, data: embedPart.data, contentType: embedPart.contentType }
              : undefined,
            embedRel?.type ?? "",
          ),
          ...(ole.placement ? { placement: ole.placement } : {}),
        })
      }
    }
    sheetSources.set(sheet, {
      partName: rel.partName,
      rootTag: read.rootTag,
      kept: read.kept,
      keptConditional: read.keptConditional,
      rowExtras: read.rowExtras,
      colExtras: read.colExtras,
      state,
      // The notes' drawing goes (legacyDrawingHF, a header picture, uses the same type and stays).
      rels: rels.filter(
        (r) =>
          r.type !== REL_COMMENTS &&
          r.type !== REL_THREADED &&
          r.type !== REL_TABLE &&
          r.type !== REL_HYPERLINK &&
          !(slicerDefs && r.type === REL_SLICER) &&
          !(drawing && r.id === read.drawing) &&
          !(picturePart && r.id === read.picture) &&
          r.type !== REL_OLE_OBJECT &&
          r.type !== REL_PACKAGE &&
          !oleImageRels.has(r.id) &&
          !(r.type === REL_VML && r.id === read.legacyDrawing),
      ),
      tables,
      drawing,
      pageExtras: read.pageExtras,
    })
    views.push(read.view)
    before += sizes[i]
  }

  // PivotTables Cells models: a cache and all its PivotTables, or (when any of them uses what Cells doesn't model)
  // none of them, kept as read.
  const pivots: PivotSource = { caches: new Map(), tables: new WeakMap(), kept: [] }
  for (const c of pivotCaches) {
    const tables = pivotTables.filter((t) => t.cache.toLowerCase() === c.partName.toLowerCase())
    const modelled =
      c.cache !== null && tables.every((t) => t.table !== null && t.table.fields.length === c.cache!.fields.length)
    if (!modelled) {
      pivots.kept.push(c.entry)
      continue
    }
    const cache = c.cache!
    wb.pivotCaches.push(cache)
    pivots.caches.set(cache.id, {
      partName: c.partName,
      recordsPart: c.recordsPart,
      key: pivotCacheKey(cache),
      records: cache.records,
    })
    for (const t of tables) {
      const table = t.table!
      table.cacheId = cache.id
      table.part = t.partName
      t.sheet.pivots.push(table)
      let byPart = pivots.tables.get(t.sheet)
      if (!byPart) pivots.tables.set(t.sheet, (byPart = new Map()))
      byPart.set(t.partName, {
        partName: t.partName,
        xml: t.xml,
        key: pivotTableKey(table),
        place: pivotPlaceKey(table),
      })
      const s = sheetSources.get(t.sheet)
      if (s) s.rels = s.rels.filter((r) => r.id !== t.relId)
    }
  }

  // Each slicer's table and column (or PivotTables and field), from its cache; one whose target is gone is left out.
  const sheetIds = sheetNodes.map((n) => Number(at(n, "sheetId") ?? 0))
  for (const sheet of slicerSheets.keys()) {
    const lost: number[] = []
    for (const o of sheet.objects) {
      if (isChart(o) || o.kind !== "slicer") continue
      const cache = slicerCaches.get(o.cache ?? "")
      const settings = cache && { sort: cache.sort, customLists: cache.customLists, crossFilter: cache.crossFilter }
      if (cache?.kind === "pivot") {
        // A PivotTable slicer: its PivotTables by sheet name; its picks are their hidden items, read with them.
        const pivotCache = wb.pivotCaches.find((c) => c.slicerId === cache.pivotCacheId)
        const tables = cache.tables.flatMap((t) => {
          const sheet = wb.sheets[sheetIds.indexOf(t.tabId)]
          return sheet?.pivots.some((p) => p.name === t.name) ? [{ sheet: sheet.name, name: t.name }] : []
        })
        if (!pivotCache || !tables.length) {
          lost.push(o.id)
          continue
        }
        Object.assign(o, { table: -1, column: cache.field, ...settings })
        o.pivot = { cacheId: pivotCache.id, field: cache.field, tables }
      } else {
        const table = cache && wb.sheets.flatMap((s) => s.tables).find((t) => t.id === cache.tableId)
        const index = cache && table ? (tableColumnIds.get(table.id) ?? []).indexOf(cache.columnId) : -1
        const column = table?.columns[index]
        if (!cache || !table || !column) {
          lost.push(o.id)
          continue
        }
        Object.assign(o, { table: table.id, column: column.name, ...settings })
      }
      if (cache!.hideNoData) o.hideNoData = true
      if (cache!.uid) o.cacheUid = cache!.uid
    }
    if (lost.length) {
      sheet.objects = sheet.objects.filter((o) => !lost.includes(o.id))
      warnings.push(`A slicer on ${sheet.name} couldn't be read; it was left out.`)
    }
  }

  // Structured references as Excel shows them ([@Qty], no table name inside the table), now that the tables are known.
  for (const sheet of wb.sheets) {
    if (!sheet.tables.length && !hasBracket(sheet)) continue
    const changes: { row: number; col: number; input: string }[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      if (input.includes("[")) {
        const next = structuredFromFile(input, tableAt(sheet, row, col)?.name ?? null)
        if (next !== input) changes.push({ row, col, input: next })
      }
    })
    for (const c of changes) sheet.cells.set(c.row, c.col, { ...sheet.cells.get(c.row, c.col)!, input: c.input })
    for (const t of sheet.tables)
      for (const col of t.columns) if (col.formula) col.formula = structuredFromFile("=" + col.formula, t.name).slice(1)
  }

  // Links to other workbooks: formulas name them by position in the file ([1]Prices!B2), by file name in Cells.
  const linkBooks = kids(kid(root, "externalReferences"), "externalReference").map((n) => {
    const r = wbRels.find((x) => x.id === at(n, "id", NS_R))
    return r?.partName ? readExternalLink(partText(pkg, r.partName) ?? null, relationshipsFor(pkg, r.partName)) : null
  })
  const linksModelled = linkBooks.every((b) => b !== null)
  if (linksModelled && linkBooks.length) {
    const books = linkBooks.filter((b) => b !== null)
    wb.links.list = books
    for (const sheet of wb.sheets) {
      const changes: { row: number; col: number; input: string }[] = []
      sheet.cells.forEachFormula((row, col, input) => {
        const next = linksFromFile(input, books)
        if (next !== input) changes.push({ row, col, input: next })
      })
      for (const c of changes) sheet.cells.set(c.row, c.col, { ...sheet.cells.get(c.row, c.col)!, input: c.input })
    }
  }

  // Defined names: workbook-wide ones Cells manages; the rest stay as read.
  const names: DefinedName[] = []
  const keptNames: XlsxSource["keptNames"] = []
  for (const n of kids(kid(root, "definedNames"), "definedName")) {
    const name = at(n, "name") ?? ""
    const hidden = at(n, "hidden") === "1" || at(n, "hidden") === "true"
    // The filter's hidden name is written anew from each sheet's AutoFilter.
    if (name === "_xlnm._FilterDatabase") continue
    // A slicer cache's name (#N/A) is written anew with its cache.
    if (slicersModelled && slicerCaches.has(name)) continue
    // The print area and titles belong to the sheet's page setup (written anew from it).
    const sheetOf = at(n, "localSheetId") !== undefined ? wb.sheets[Number(at(n, "localSheetId"))] : undefined
    if (sheetOf && name === "_xlnm.Print_Area") {
      const area = readPrintArea(n.text)
      if (area) {
        sheetOf.pageSetup.printArea = area
        continue
      }
    }
    if (sheetOf && name === "_xlnm.Print_Titles") {
      const titles = readPrintTitles(n.text)
      if (titles) {
        sheetOf.pageSetup.titleRows = titles.rows
        sheetOf.pageSetup.titleCols = titles.cols
        continue
      }
    }
    if (at(n, "localSheetId") !== undefined || hidden || name.startsWith("_xlnm.") || name.startsWith("_xl")) {
      const local = at(n, "localSheetId")
      keptNames.push({ xml: raw(wbDoc, n), sheet: local !== undefined ? (wb.sheets[Number(local)] ?? null) : null })
      continue
    }
    const def: DefinedName = { name, ref: linksFromFile(formulaFromFile(n.text), wb.links.list).slice(1) }
    const comment = at(n, "comment")
    if (comment) def.comment = comment
    names.push(def)
  }
  wb.names.replace(names)

  // Text links: each query table placed by its sheet-level name; Cells writes those three parts anew.
  const modelled = new Set<number>()
  for (const l of linked) {
    const k = keptNames.findIndex((n) => n.sheet === l.sheet && at(parseTree(n.xml).root, "name") === l.name)
    if (k < 0) continue
    const area = readPrintArea(parseTree(keptNames[k].xml).root.text)
    const q = area?.length === 1 ? queryTableFrom(l.name, connections.get(l.connectionId)!, area[0]) : null
    if (!q) continue
    l.sheet.queryTables.push(q)
    keptNames.splice(k, 1)
    modelled.add(l.connectionId)
    const s = sheetSources.get(l.sheet)
    if (s) s.rels = s.rels.filter((r) => r.id !== l.relId)
  }
  const keptConnections = [...connections.values()].filter((c) => !modelled.has(c.id))

  // View › Custom Views: the views, each sheet's part, and the hidden names that hold their rows, columns and print
  // area (written anew from the views; other names of a view, such as its filter, stay as read).
  const viewName = (sheet: Sheet, name: string): string | null => {
    const k = keptNames.findIndex((n) => n.sheet === sheet && at(parseTree(n.xml).root, "name") === name)
    if (k < 0) return null
    const text = parseTree(keptNames[k].xml).root.text
    keptNames.splice(k, 1)
    return text
  }
  wb.customViews = readCustomViews(
    kid(root, "customWorkbookViews"),
    wb.sheets,
    sheetNodes.map((n) => Number(at(n, "sheetId") ?? 0)),
    customSheetViews,
    viewName,
  )

  const prot = kid(root, "workbookProtection")
  if (prot && (at(prot, "lockStructure") === "1" || at(prot, "lockStructure") === "true")) {
    const p: WorkbookProtection = {}
    const legacy = at(prot, "workbookPassword")
    if (legacy) p.password = legacy.toUpperCase()
    const hash = readPasswordHash(prot, "workbook")
    if (hash) p.hash = hash
    wb.structure.protection = p
  }

  const bookView = kid(kid(root, "bookViews"), "workbookView")
  let activeSheet = Number(at(bookView, "activeTab") ?? 0)
  if (!(activeSheet >= 0 && activeSheet < wb.sheets.length) || wb.sheets[activeSheet].hidden)
    activeSheet = Math.max(
      0,
      wb.sheets.findIndex((s) => !s.hidden),
    )

  const calcPr = kid(root, "calcPr")
  await options.progress?.(0.95)
  // The file's values were calculated by whoever saved it: they stand until something they depend on changes.
  // Dynamic arrays are calculated again, since their spilled cells aren't read (and Excel asks for it with fullCalcOnLoad).
  const recalc = dynamic || at(calcPr, "fullCalcOnLoad") === "1" || at(calcPr, "fullCalcOnLoad") === "true"
  wb.calc.rebuild(!recalc)
  if (recalc && options.calculate !== false) wb.calc.calculateNow(true)
  if (at(calcPr, "calcMode") === "manual") wb.calc.mode = "manual"
  wb.clearHistory()

  return {
    workbook: wb,
    views,
    activeSheet,
    source: {
      pkg,
      workbook: wbDoc,
      styles,
      reuse,
      rich,
      keptNames,
      sheets: sheetSources,
      dynamic,
      keptConnections,
      linksModelled,
      slicersModelled,
      theme: themeXml ? copyTheme(wb.theme) : undefined,
      pivots: pivotCaches.length ? pivots : undefined,
    },
    warnings: [...new Set(warnings)],
  }
}
