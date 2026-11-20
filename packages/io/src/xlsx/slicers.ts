import type { SheetSlicer } from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import { at, attrText, kid, kids, parseTree, type XNode } from "./dom"

/**
 * Slicers for tables, as Excel 2013 and later write them:
 * - a slicer cache per table column (`xl/slicerCaches/slicerCacheN.xml`, `x15:tableSlicerCache`), listed in the
 *   workbook's extLst and named by a defined name ("Slicer_Region" = #N/A);
 * - each sheet's slicers (`xl/slicers/slicerN.xml`: caption, style, columns, button height), listed in the sheet's
 *   extLst;
 * - each slicer's frame in the sheet's drawing (an `sle:slicer` graphic frame, with a fallback shape for older
 *   Excel).
 * Which items are picked is the table's own AutoFilter. Slicers on PivotTables aren't modelled: a file that has any
 * keeps all its slicers as they were.
 */

export const REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer"
export const REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
export const CT_SLICER = "application/vnd.ms-excel.slicer+xml"
export const CT_SLICER_CACHE = "application/vnd.ms-excel.slicerCache+xml"
/** The workbook's x15 list of slicer caches, and x14's (PivotTable slicers). */
export const URI_WB_SLICER_CACHES = "{46BE6895-7355-4a93-B00E-2C351335B9C9}"
export const URI_WB_PIVOT_SLICER_CACHES = "{BBE1A952-AA13-448e-AADC-164F8A28A991}"
/** A sheet's list of table slicers (x15), and of PivotTable slicers (x14). */
export const URI_SHEET_SLICERS = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"
export const URI_SHEET_PIVOT_SLICERS = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"
const URI_TABLE_CACHE = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"
const URI_HIDE_NO_DATA = "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}"
const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const NS_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
const NS_X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const NS_XR10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"
const NS_SLE = "http://schemas.microsoft.com/office/drawing/2010/slicer"
const NS_SLE15 = "http://schemas.microsoft.com/office/drawing/2012/slicer"
const NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
export const SLICER_GRAPHIC_URI = "http://schemas.microsoft.com/office/drawing/2010/slicer"

const EMU = 9525
const falsy = (v: string | undefined) => v === "0" || v === "false"
const truthy = (v: string | undefined) => v === "1" || v === "true"

/** A slicer cache as read: a table column's, or a PivotTable cache field's, and how its items are listed. */
export type SlicerCacheRead = {
  name: string
  sort: SheetSlicer["sort"]
  customLists: boolean
  crossFilter: SheetSlicer["crossFilter"]
  hideNoData: boolean
  uid?: string
} & (
  | {
      kind: "table"
      tableId: number
      /** The table column's id (1-based, as in the table part). */
      columnId: number
    }
  | {
      kind: "pivot"
      /** The cache field's name. */
      field: string
      /** The PivotTables it filters: their sheet's sheetId and name. */
      tables: { tabId: number; name: string }[]
      /** The x14 id of their PivotTable cache. */
      pivotCacheId: number
    }
)

const crossOf = (v: string | undefined): SheetSlicer["crossFilter"] =>
  v === "none" || v === "showItemsWithNoData" ? v : "showItemsWithDataAtTop"

/** A slicer cache part: a table's or a PivotTable's, or null (an OLAP one, or unreadable). */
export function readSlicerCache(xml: string): SlicerCacheRead | null {
  const root = parseTree(xml).root
  const common = {
    name: at(root, "name") ?? "",
    hideNoData: /slicerCacheHideItemsWithNoData\b/.test(xml),
    uid: at(root, "xr10:uid") ?? at(root, "uid"),
  }
  const tabular = kid(kid(root, "data"), "tabular")
  if (tabular) {
    const pivotCacheId = Number(at(tabular, "pivotCacheId"))
    const field = at(root, "sourceName")
    if (!Number.isFinite(pivotCacheId) || !field) return null
    const tables = kids(kid(root, "pivotTables"), "pivotTable").map((t) => ({
      tabId: Number(at(t, "tabId") ?? 0),
      name: at(t, "name") ?? "",
    }))
    return {
      ...common,
      kind: "pivot",
      field,
      tables,
      pivotCacheId,
      sort: at(tabular, "sortOrder") === "descending" ? "descending" : "ascending",
      customLists: !falsy(at(tabular, "customListSort")),
      crossFilter: crossOf(at(tabular, "crossFilter")),
    }
  }
  const table = kids(kid(root, "extLst"), "ext")
    .map((e) => kid(e, "tableSlicerCache"))
    .find(Boolean)
  if (!table || kid(root, "pivotTables") || kid(root, "data")) return null
  const tableId = Number(at(table, "tableId"))
  const columnId = Number(at(table, "column"))
  if (!Number.isFinite(tableId) || !Number.isFinite(columnId)) return null
  return {
    ...common,
    kind: "table",
    tableId,
    columnId,
    sort: at(table, "sortOrder") === "descending" ? "descending" : "ascending",
    customLists: !falsy(at(table, "customListSort")),
    crossFilter: crossOf(at(table, "crossFilter")),
  }
}

/** One `<slicer>` of a sheet's slicers part. */
export type SlicerRead = {
  name: string
  cache: string
  caption: string
  header: boolean
  style?: string
  columns: number
  /** Button height in px. */
  buttonHeight: number
  startItem?: number
  locked: boolean
  uid?: string
}

export function readSlicerPart(xml: string): SlicerRead[] {
  return kids(parseTree(xml).root, "slicer").map((n: XNode) => {
    const height = Number(at(n, "rowHeight"))
    const start = Number(at(n, "startItem") ?? 0)
    return {
      name: at(n, "name") ?? "",
      cache: at(n, "cache") ?? "",
      caption: at(n, "caption") ?? at(n, "name") ?? "",
      header: !falsy(at(n, "showCaption")),
      style: at(n, "style"),
      columns: Math.max(1, Number(at(n, "columnCount") ?? 1) || 1),
      buttonHeight: Number.isFinite(height) && height > 0 ? Math.round(height / EMU) : 27,
      ...(start > 0 ? { startItem: start } : {}),
      locked: truthy(at(n, "lockedPosition")),
      uid: at(n, "xr10:uid") ?? at(n, "uid"),
    }
  })
}

/** The slicer's name, when a drawing anchor is a slicer frame. */
export function slicerFrameName(anchor: XNode): string | null {
  const alternate = kid(anchor, "AlternateContent")
  for (const choice of kids(alternate, "Choice")) {
    const data = kid(kid(kid(choice, "graphicFrame"), "graphic"), "graphicData")
    if (at(data, "uri") === SLICER_GRAPHIC_URI) return at(kid(data, "slicer"), "name") ?? null
  }
  return null
}

// ── Writing ────────────────────────────────────────────────────────────────────────────────────────────────────

const esc = (v: string) => escapeXmlText(v).replace(/"/g, "&quot;")

/** A slicer's frame in the drawing (inside its anchor), with Excel's fallback for versions without slicers. */
export function slicerFrameXml(
  o: SheetSlicer,
  id: number,
  box: { x: number; y: number; w: number; h: number },
): string {
  const e = (px: number) => Math.round(px * EMU)
  const descr = o.altText ? ` descr="${esc(o.altText)}"` : ""
  const hidden = o.hidden ? ` hidden="1"` : ""
  const locks = o.locked ? `<a:graphicFrameLocks noMove="1" noResize="1"/>` : ""
  // A PivotTable slicer is Excel 2010's (a14); a table slicer Excel 2013's (sle15).
  const choice = o.pivot
    ? `<mc:Choice xmlns:a14="${NS_A14}" Requires="a14">`
    : `<mc:Choice xmlns:sle15="${NS_SLE15}" Requires="sle15">`
  const fallback = o.pivot
    ? "This shape represents a slicer. Slicers are supported in Excel 2010 or later."
    : "This shape represents a table slicer. Table slicers are not supported in this version of Excel."
  return (
    `<mc:AlternateContent xmlns:mc="${NS_MC}">${choice}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(o.name)}"${descr}${hidden}/>` +
    `<xdr:cNvGraphicFramePr>${locks}</xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${SLICER_GRAPHIC_URI}"><sle:slicer xmlns:sle="${NS_SLE}" name="${esc(o.name)}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback>` +
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id + 100000}" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="${e(box.x)}" y="${e(box.y)}"/><a:ext cx="${e(box.w)}" cy="${e(box.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill>` +
    `<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
    `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
    `<a:t>${fallback}</a:t>` +
    `</a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>`
  )
}

/** A sheet's slicers part. */
export function slicerPartXml(slicers: readonly { slicer: SheetSlicer; cache: string }[]): string {
  const items = slicers.map(({ slicer: s, cache }) => {
    const attrs = attrText({
      name: s.name,
      "xr10:uid": s.uid,
      cache,
      caption: s.caption,
      startItem: s.startItem || undefined,
      columnCount: s.columns > 1 ? s.columns : undefined,
      showCaption: s.header ? undefined : "0",
      style: s.style === "SlicerStyleLight1" ? undefined : s.style,
      lockedPosition: s.locked ? "1" : undefined,
      rowHeight: Math.round(s.buttonHeight * EMU),
    })
    return `<slicer${attrs}/>`
  })
  return (
    `<slicers xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x xr10" xmlns:x="${NS_X}" xmlns:xr10="${NS_XR10}">` +
    `${items.join("")}</slicers>`
  )
}

/** A table slicer cache part. */
export function slicerCacheXml(c: {
  name: string
  column: string
  tableId: number
  columnId: number
  slicer: SheetSlicer
}): string {
  const s = c.slicer
  const table = attrText({
    tableId: c.tableId,
    column: c.columnId,
    sortOrder: s.sort === "descending" ? "descending" : undefined,
    customListSort: s.customLists ? undefined : "0",
    crossFilter: s.crossFilter === "showItemsWithDataAtTop" ? undefined : s.crossFilter,
  })
  const hide = s.hideNoData
    ? `<x:ext uri="${URI_HIDE_NO_DATA}" xmlns:x15="${NS_X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext>`
    : ""
  return (
    `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x xr10" xmlns:x="${NS_X}" xmlns:xr10="${NS_XR10}"` +
    `${attrText({ name: c.name, "xr10:uid": s.cacheUid, sourceName: c.column })}>` +
    `<extLst><x:ext uri="${URI_TABLE_CACHE}" xmlns:x15="${NS_X15}"><x15:tableSlicerCache${table}/></x:ext>${hide}</extLst>` +
    `</slicerCacheDefinition>`
  )
}

/** A PivotTable slicer cache part: its PivotTables (sheetId and name) and the field's items, picked or not. */
export function pivotSlicerCacheXml(c: {
  name: string
  field: string
  slicer: SheetSlicer
  pivotCacheId: number
  tables: readonly { tabId: number; name: string }[]
  items: readonly { x: number; selected: boolean; data: boolean }[]
}): string {
  const s = c.slicer
  const tabular = attrText({
    pivotCacheId: c.pivotCacheId,
    sortOrder: s.sort === "descending" ? "descending" : undefined,
    customListSort: s.customLists ? undefined : "0",
    crossFilter: s.crossFilter === "showItemsWithDataAtTop" ? undefined : s.crossFilter,
  })
  const items = c.items.map(
    (i) => `<i${attrText({ x: i.x, s: i.selected ? 1 : undefined, nd: i.data ? undefined : 1 })}/>`,
  )
  const hide = s.hideNoData
    ? `<extLst><x:ext uri="${URI_HIDE_NO_DATA}" xmlns:x15="${NS_X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext></extLst>`
    : ""
  return (
    `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x xr10" xmlns:x="${NS_X}" xmlns:xr10="${NS_XR10}"` +
    `${attrText({ name: c.name, "xr10:uid": s.cacheUid, sourceName: c.field })}>` +
    `<pivotTables>${c.tables.map((t) => `<pivotTable${attrText({ tabId: t.tabId, name: t.name })}/>`).join("")}</pivotTables>` +
    `<data><tabular${tabular}><items count="${items.length}">${items.join("")}</items></tabular></data>${hide}` +
    `</slicerCacheDefinition>`
  )
}

/** The workbook's extLst entry listing the PivotTable slicer caches (x14). */
export const workbookPivotSlicersExt = (relIds: readonly string[]) =>
  relIds.length
    ? `<ext uri="${URI_WB_PIVOT_SLICER_CACHES}" xmlns:x14="${NS_X14}"><x14:slicerCaches>` +
      relIds.map((id) => `<x14:slicerCache r:id="${id}"/>`).join("") +
      `</x14:slicerCaches></ext>`
    : ""

/** A sheet's extLst entry pointing at its PivotTable slicers part (x14). */
export const sheetPivotSlicersExt = (relId: string) =>
  `<ext uri="${URI_SHEET_PIVOT_SLICERS}" xmlns:x14="${NS_X14}"><x14:slicerList><x14:slicer r:id="${relId}"/></x14:slicerList></ext>`

/** The workbook's extLst entry listing the slicer caches (by relationship id). */
export const workbookSlicersExt = (relIds: readonly string[]) =>
  relIds.length
    ? `<ext uri="${URI_WB_SLICER_CACHES}" xmlns:x15="${NS_X15}"><x15:slicerCaches xmlns:x14="${NS_X14}">` +
      relIds.map((id) => `<x14:slicerCache r:id="${id}"/>`).join("") +
      `</x15:slicerCaches></ext>`
    : ""

/** A sheet's extLst entry pointing at its slicers part. */
export const sheetSlicersExt = (relId: string) =>
  `<ext uri="${URI_SHEET_SLICERS}" xmlns:x15="${NS_X15}"><x14:slicerList xmlns:x14="${NS_X14}">` +
  `<x14:slicer r:id="${relId}"/></x14:slicerList></ext>`

/** An extLst without the `<ext>`s of these uris (empty when nothing is left). */
export function withoutExts(extLst: string, uris: readonly string[]): string {
  let out = extLst
  for (const uri of uris) {
    const u = uri.replace(/[{}-]/g, (c) => "\\" + c)
    out = out.replace(new RegExp(`<(\\w+:)?ext\\b[^>]*uri="${u}"[\\s\\S]*?</\\1?ext>`, "gi"), "")
  }
  return /<(\w+:)?ext\b/.test(out) ? out : ""
}

/** Excel's name for a new slicer cache: "Slicer_" and the column name made a valid name, numbered if taken. */
export function slicerCacheName(column: string, taken: ReadonlySet<string>): string {
  const base = "Slicer_" + (column.replace(/[^\p{L}\p{N}_.]/gu, "_") || "Column")
  let name = base
  for (let n = 1; taken.has(name.toUpperCase()); n++) name = base + n
  return name
}
