import { readCustomViews, type CustomSheetViewRead } from "./custom-views"
import {
  copyTheme,
  OFFICE_THEME,
  tableAt,
  Workbook,
  type DefinedName,
  type Sheet,
  type Theme,
  type WorkbookProtection,
} from "@cells/core"
import { decodeXml, parseXmlEvents } from "@cells/ooxml"
import { findPart, openPackage, relationshipsFor, type OpcPackage, type OpcRelationship } from "../opc"
import { themeColorsFrom, DEFAULT_THEME_COLORS, type ThemeColors } from "./colors"
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
import { at, kid, kids, parseTree, raw, type Kept, type XDoc, type XNode } from "./dom"
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

/** The theme as shapes use it: its name, colours (by slot) and fonts. */
function themeOf(xml: string | undefined, colors: ThemeColors): Theme {
  const order = [
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
  ] as const
  const theme = copyTheme(OFFICE_THEME)
  if (!xml) return theme
  order.forEach((slot, i) => (theme.colors[slot] = colors[i]))
  const doc = parseTree(xml)
  theme.name = at(doc.root, "name") ?? theme.name
  const find = (n: XNode, local: string): XNode | undefined =>
    n.local === local ? n : n.children.map((c) => find(c, local)).find((x) => x !== undefined)
  const font = (local: string) => at(kid(find(doc.root, local) ?? doc.root, "latin"), "typeface")
  theme.fonts = { major: font("majorFont") || theme.fonts.major, minor: font("minorFont") || theme.fonts.minor }
  return theme
}

/** The theme's colour scheme (`a:clrScheme`). */
function readTheme(xml: string | undefined): ThemeColors {
  if (!xml) return DEFAULT_THEME_COLORS
  const scheme: Record<string, string> = {}
  const find = (n: XNode): XNode | undefined =>
    n.local === "clrScheme" ? n : n.children.map(find).find((x) => x !== undefined)
  const root = find(parseTree(xml).root)
  if (!root) return DEFAULT_THEME_COLORS
  for (const slot of root.children) {
    const color = slot.children[0]
    if (!color) continue
    const hex = color.local === "sysClr" ? at(color, "lastClr") : at(color, "val")
    if (hex && /^[0-9A-Fa-f]{6}$/.test(hex)) scheme[slot.local] = hex.toUpperCase()
  }
  return themeColorsFrom(scheme)
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
  const theme = readTheme(themeXml)
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
  wb.theme = themeOf(themeXml, theme)

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
    if (vml && /ObjectType="(?!Note")/.test(vml))
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
      const table = readTable(partText(pkg, r.partName) ?? "", styles, warnings)
      if (!table) {
        warnings.push(`A table on ${sheet.name} couldn't be read; it was left out.`)
        continue
      }
      sheet.tables.push(table)
      tables.push({ id: table.id, partName: r.partName })
    }
    for (const r of rels) {
      if (r.type !== REL_QUERY_TABLE || !r.partName) continue
      const q = readQueryTable(partText(pkg, r.partName) ?? null)
      if (q && connections.get(q.connectionId)?.text) linked.push({ sheet, relId: r.id, ...q })
    }
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
          )
        : undefined
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
          !(drawing && r.id === read.drawing) &&
          !(picturePart && r.id === read.picture) &&
          !(r.type === REL_VML && r.id === read.legacyDrawing),
      ),
      tables,
      drawing,
      pageExtras: read.pageExtras,
    })
    views.push(read.view)
    before += sizes[i]
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
    },
    warnings: [...new Set(warnings)],
  }
}
