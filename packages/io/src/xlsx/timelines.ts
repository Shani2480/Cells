import { TIMELINE_LEVELS, type SheetTimeline, type TimelineLevel } from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import { at, attrText, kid, kids, parseTree, type XNode } from "./dom"
import { isoToSerial, serialToIso } from "./pivot"

/**
 * Timelines on PivotTables, as Excel 2013 and later write them:
 * - a timeline cache per date field and set of PivotTables (`xl/timelineCaches/timelineCacheN.xml`): the PivotTables,
 *   the PivotTable cache's x14 id, the dates picked and the span, listed in the workbook's extLst and named by a
 *   defined name ("NativeTimeline_Date" = #N/A);
 * - each sheet's timelines (`xl/timelines/timelineN.xml`: caption, what shows, level, scroll position, style), listed
 *   in the sheet's extLst;
 * - each timeline's frame in the sheet's drawing (a `tsle:timeslicer` graphic frame with a fallback shape).
 * What is picked is the PivotTables' own whole-day date filters on the field.
 */

export const REL_TIMELINE = "http://schemas.microsoft.com/office/2011/relationships/timeline"
export const REL_TIMELINE_CACHE = "http://schemas.microsoft.com/office/2011/relationships/timelineCache"
export const CT_TIMELINE = "application/vnd.ms-excel.timeline+xml"
export const CT_TIMELINE_CACHE = "application/vnd.ms-excel.timelineCache+xml"
/** The workbook's list of timeline caches and a sheet's list of timelines (x15). */
export const URI_WB_TIMELINE_CACHES = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"
export const URI_SHEET_TIMELINES = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}"
export const TIMELINE_GRAPHIC_URI = "http://schemas.microsoft.com/office/drawing/2012/timeslicer"
const NS_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
const NS_X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
const NS_XR10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"

const falsy = (v: string | undefined) => v === "0" || v === "false"
const truthy = (v: string | undefined) => v === "1" || v === "true"
const esc = (v: string) => escapeXmlText(v).replace(/"/g, "&quot;")

export type TimelineCacheRead = {
  name: string
  field: string
  tables: { tabId: number; name: string }[]
  pivotCacheId: number
  uid?: string
}

/** A timeline cache part, or null (one on the Data Model, or unreadable). */
export function readTimelineCache(xml: string): TimelineCacheRead | null {
  const root = parseTree(xml).root
  const state = kid(root, "state")
  const pivotCacheId = Number(at(state, "pivotCacheId"))
  const field = at(root, "sourceName")
  if (!state || !Number.isFinite(pivotCacheId) || !field || kid(state, "olap") !== undefined) return null
  return {
    name: at(root, "name") ?? "",
    field,
    tables: kids(kid(root, "pivotTables"), "pivotTable").map((t) => ({
      tabId: Number(at(t, "tabId") ?? 0),
      name: at(t, "name") ?? "",
    })),
    pivotCacheId,
    uid: at(root, "xr10:uid") ?? at(root, "uid"),
  }
}

/** One `<timeline>` of a sheet's timelines part. */
export type TimelineRead = Pick<
  SheetTimeline,
  "name" | "caption" | "header" | "selectionLabel" | "scrollbar" | "timeLevel" | "level" | "selectionLevel" | "style"
> & { cache: string; scrollPosition?: number; locked: boolean; uid?: string }

const levelOf = (v: string | undefined, fallback: TimelineLevel): TimelineLevel =>
  TIMELINE_LEVELS[Number(v)] ?? fallback

export function readTimelinePart(xml: string): TimelineRead[] {
  return kids(parseTree(xml).root, "timeline").map((n: XNode) => {
    const scroll = at(n, "scrollPosition")
    const serial = scroll ? isoToSerial(scroll) : null
    const out: TimelineRead = {
      name: at(n, "name") ?? "",
      cache: at(n, "cache") ?? "",
      caption: at(n, "caption") ?? at(n, "name") ?? "",
      header: !falsy(at(n, "showHeader")),
      selectionLabel: !falsy(at(n, "showSelectionLabel")),
      timeLevel: !falsy(at(n, "showTimeLevel")),
      scrollbar: !falsy(at(n, "showHorizontalScrollbar")),
      level: levelOf(at(n, "level"), "years"),
      selectionLevel: levelOf(at(n, "selectionLevel"), "years"),
      style: at(n, "style") ?? "TimeSlicerStyleLight1",
      locked: truthy(at(n, "lockedPosition")),
      uid: at(n, "xr10:uid") ?? at(n, "uid"),
    }
    if (serial !== null) out.scrollPosition = Math.floor(serial)
    return out
  })
}

/** The timeline's name, when a drawing anchor is a timeline frame. */
export function timelineFrameName(anchor: XNode): string | null {
  for (const choice of kids(kid(anchor, "AlternateContent"), "Choice")) {
    const data = kid(kid(kid(choice, "graphicFrame"), "graphic"), "graphicData")
    if (at(data, "uri") === TIMELINE_GRAPHIC_URI) return at(kid(data, "timeslicer"), "name") ?? null
  }
  return null
}

// ── Writing ────────────────────────────────────────────────────────────────────────────────────────────────────

/** A timeline's frame in the drawing (inside its anchor), with Excel's fallback for versions without timelines. */
export function timelineFrameXml(
  o: SheetTimeline,
  id: number,
  box: { x: number; y: number; w: number; h: number },
): string {
  const e = (px: number) => Math.round(px * 9525)
  const descr = o.altText ? ` descr="${esc(o.altText)}"` : ""
  const hidden = o.hidden ? ` hidden="1"` : ""
  const locks = o.locked ? `<a:graphicFrameLocks noMove="1" noResize="1"/>` : ""
  return (
    `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:tsle="${TIMELINE_GRAPHIC_URI}" Requires="tsle">` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(o.name)}"${descr}${hidden}/>` +
    `<xdr:cNvGraphicFramePr>${locks}</xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${TIMELINE_GRAPHIC_URI}"><tsle:timeslicer xmlns:tsle="${TIMELINE_GRAPHIC_URI}" name="${esc(o.name)}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback>` +
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id + 100000}" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="${e(box.x)}" y="${e(box.y)}"/><a:ext cx="${e(box.w)}" cy="${e(box.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill>` +
    `<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
    `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
    `<a:t>Timeline: Works in Excel 2013 or higher. Do not move or resize.</a:t>` +
    `</a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>`
  )
}

/** A sheet's timelines part. */
export function timelinePartXml(list: readonly { timeline: SheetTimeline; cache: string }[]): string {
  const items = list.map(({ timeline: t, cache }) => {
    const attrs = attrText({
      name: t.name,
      "xr10:uid": t.uid,
      cache,
      caption: t.caption,
      showHeader: t.header ? undefined : 0,
      showSelectionLabel: t.selectionLabel ? undefined : 0,
      showTimeLevel: t.timeLevel ? undefined : 0,
      showHorizontalScrollbar: t.scrollbar ? undefined : 0,
      level: TIMELINE_LEVELS.indexOf(t.level),
      selectionLevel: TIMELINE_LEVELS.indexOf(t.selectionLevel),
      scrollPosition: t.scrollPosition !== undefined ? serialToIso(t.scrollPosition) : undefined,
      style: t.style === "TimeSlicerStyleLight1" ? undefined : t.style,
      lockedPosition: t.locked ? 1 : undefined,
    })
    return `<timeline${attrs}/>`
  })
  return (
    `<timelines xmlns="${NS_X15}" xmlns:mc="${NS_MC}" mc:Ignorable="x xr10" xmlns:x="${NS_X}" xmlns:xr10="${NS_XR10}">` +
    `${items.join("")}</timelines>`
  )
}

/** A timeline cache part: its PivotTables (sheetId and name), the dates picked, and the span shown. */
export function timelineCacheXml(c: {
  name: string
  field: string
  uid?: string
  pivotCacheId: number
  tables: readonly { tabId: number; name: string }[]
  picked: { from: number; to: number } | null
  bounds: { start: number; end: number } | null
}): string {
  const state = attrText({
    minimalRefreshVersion: 6,
    lastRefreshVersion: 6,
    pivotCacheId: c.pivotCacheId,
    filterType: c.picked ? "dateBetween" : "unknown",
  })
  const selection = c.picked
    ? `<selection${attrText({ startDate: serialToIso(c.picked.from), endDate: serialToIso(c.picked.to) })}/>`
    : ""
  const bounds = c.bounds
    ? `<bounds${attrText({ startDate: serialToIso(c.bounds.start), endDate: serialToIso(c.bounds.end) })}/>`
    : ""
  return (
    `<timelineCacheDefinition xmlns="${NS_X15}" xmlns:mc="${NS_MC}" mc:Ignorable="xr10" xmlns:xr10="${NS_XR10}"` +
    `${attrText({ name: c.name, "xr10:uid": c.uid, sourceName: c.field })}>` +
    `<pivotTables>${c.tables.map((t) => `<pivotTable${attrText({ tabId: t.tabId, name: t.name })}/>`).join("")}</pivotTables>` +
    `<state${state}>${selection}${bounds}</state></timelineCacheDefinition>`
  )
}

/** The workbook's extLst entry listing the timeline caches. */
export const workbookTimelinesExt = (relIds: readonly string[]) =>
  relIds.length
    ? `<ext uri="${URI_WB_TIMELINE_CACHES}" xmlns:x15="${NS_X15}"><x15:timelineCacheRefs>` +
      relIds.map((id) => `<x15:timelineCacheRef r:id="${id}"/>`).join("") +
      `</x15:timelineCacheRefs></ext>`
    : ""

/** A sheet's extLst entry pointing at its timelines part. */
export const sheetTimelinesExt = (relId: string) =>
  `<ext uri="${URI_SHEET_TIMELINES}" xmlns:x15="${NS_X15}"><x15:timelineRefs><x15:timelineRef r:id="${relId}"/></x15:timelineRefs></ext>`

/** Excel's name for a new timeline cache: "NativeTimeline_" and the field made a valid name, numbered if taken. */
export function timelineCacheName(field: string, taken: ReadonlySet<string>): string {
  const base = "NativeTimeline_" + (field.replace(/[^\p{L}\p{N}_.]/gu, "_") || "Date")
  let name = base
  for (let n = 1; taken.has(name.toUpperCase()); n++) name = base + n
  return name
}
