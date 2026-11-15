import {
  tableAt,
  columnName,
  isError,
  MAX_COLS,
  MAX_ROWS,
  type CellData,
  type CFEntry,
  type CFRule,
  type CFThreshold,
  type Range,
  type Sheet,
  type StyleTable,
} from "@cells/core"
import { escapeXml, escapeXmlText } from "../xml"
import { autoFilterXml } from "./autofilter"
import { argb } from "./colors"
import { NS_MAIN, NS_R } from "./constants"
import { assemble, attrText, type Kept } from "./dom"
import { printElementsXml, withFitToPage } from "./page-setup"
import { formulaToFile } from "./formula"
import { PROTECTION_FLAGS, type SheetView } from "./sheet-read"
import type { SheetSource } from "./read"
import type { StylesOut } from "./styles"
import { pxToColWidth, pxToPoints } from "./units"
import { validationExtXml, validationsXml, withExt } from "./validation"
import { structuredToFile } from "./table"

/** The shared string table being written: each distinct text once, rich text as read where it is unchanged. */
export class SharedStrings {
  private readonly index = new Map<string, number>()
  private readonly items: string[] = []
  count = 0

  constructor(private readonly rich: ReadonlyMap<string, string> = new Map()) {}

  add(text: string): number {
    this.count++
    let i = this.index.get(text)
    if (i === undefined) {
      i = this.items.length
      this.index.set(text, i)
      this.items.push(
        this.rich.get(text) ??
          `<si><t${/^\s|\s$|\n/.test(text) ? ` xml:space="preserve"` : ""}>${escapeText(text)}</t></si>`,
      )
    }
    return i
  }

  get size() {
    return this.items.length
  }

  xml(): string {
    return `<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.items.length}">${this.items.join("")}</sst>`
  }
}

/** Characters XML 1.0 can't hold (C0 controls other than tab and line breaks, U+FFFE, U+FFFF). */
// eslint-disable-next-line no-control-regex -- these are exactly the characters XML cannot hold
const CONTROL = new RegExp("[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF]", "g")

/** Text for XML: control characters XML can't hold go as Excel's `_xHHHH_` escapes (and a literal `_x…_` is escaped too). */
export function escapeText(text: string): string {
  let s = text.replace(/_x([0-9A-Fa-f]{4})_/g, "_x005F_x$1_")
  s = s.replace(CONTROL, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`)
  return escapeXmlText(s)
}

/** Schema order of a worksheet's children. */
export const WORKSHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
  "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
  "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
  "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
  "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
  "extLst",
] // prettier-ignore

const ref = (row: number, col: number) => columnName(col) + (row + 1)
export const rangeRef = (r: Range) =>
  r.r0 === r.r1 && r.c0 === r.c1 ? ref(r.r0, r.c0) : `${ref(r.r0, r.c0)}:${ref(r.r1, r.c1)}`

export type SheetWriteContext = {
  sheet: Sheet
  index: number
  view: SheetView
  styleTable: StyleTable
  styles: StylesOut
  strings: SharedStrings
  source: SheetSource | undefined
  /** Spill ranges of dynamic-array formulas on this sheet, by anchor "row,col". */
  spills: ReadonlyMap<string, Range>
  /** The relationship id of the sheet's notes drawing, when it has notes. */
  legacyDrawing: string | null
  /** The relationship id of the sheet's drawing (charts, and pictures kept from the file). */
  drawing?: string | null
  /** The relationship ids of the sheet's table parts. */
  tableParts?: string[]
}

export type SheetWritten = { xml: string; dynamic: boolean }

const numberText = (n: number) => (Number.isFinite(n) ? String(n) : "0")

/** One worksheet part. */
export function writeSheet(ctx: SheetWriteContext): SheetWritten {
  const { sheet, styles, styleTable } = ctx
  const xf = (id: number) => (id ? styles.xf(id, styleTable.get(id)) : 0)
  let dynamic = false

  // ── sheetData, band by band (the store is column-major in 1,024-row chunks) ──
  const bounds = sheet.cells.usedBounds()
  const special = specialRows(sheet, ctx.source?.rowExtras)
  const rowParts: string[] = []
  const maxRow = Math.max(bounds?.r1 ?? -1, special.length ? special[special.length - 1] : -1)
  const BAND = 1024
  let si = 0
  for (let band = 0; band <= maxRow; band += BAND) {
    const r1 = Math.min(band + BAND - 1, MAX_ROWS - 1)
    const byRow = new Map<number, string[]>()
    if (bounds && bounds.r0 <= r1 && bounds.r1 >= band) {
      sheet.cells.forEach({ r0: band, c0: 0, r1, c1: MAX_COLS - 1 }, (row, col, data) => {
        let list = byRow.get(row)
        if (!list) byRow.set(row, (list = []))
        const c = cellXml(row, col, data)
        if (c) list.push(c)
      })
    }
    while (si < special.length && special[si] <= r1) {
      if (!byRow.has(special[si])) byRow.set(special[si], [])
      si++
    }
    for (const row of [...byRow.keys()].sort((a, b) => a - b)) rowParts.push(rowXml(row, byRow.get(row)!))
  }

  function rowXml(row: number, cells: string[]): string {
    const style = sheet.rowStyles.get(row)
    const custom = sheet.rows.hasCustomSize(row)
    const attrs = attrText({
      r: row + 1,
      s: style ? xf(style) : undefined,
      customFormat: style ? true : undefined,
      ht: custom ? pxToPoints(sheet.rows.storedSize(row)) : undefined,
      hidden: sheet.rows.isHidden(row),
      customHeight: custom && !sheet.rows.isAuto(row) ? true : undefined,
    })
    const extra = ctx.source?.rowExtras.get(row) ?? ""
    return cells.length ? `<row${attrs}${extra}>${cells.join("")}</row>` : `<row${attrs}${extra}/>`
  }

  function cellXml(row: number, col: number, data: CellData): string {
    const r = ref(row, col)
    const s = data.style ? xf(data.style) : 0
    const sAttr = s ? ` s="${s}"` : ""
    const value = data.value
    const formula = !data.spill && data.input?.startsWith("=") ? data.input : null
    let t = ""
    let v = ""
    if (value !== undefined) {
      if (typeof value === "number") v = numberText(value)
      else if (typeof value === "boolean") {
        t = "b"
        v = value ? "1" : "0"
      } else if (isError(value)) {
        t = "e"
        v = value.error
      } else if (formula) {
        t = "str"
        v = escapeText(value)
      } else {
        t = "s"
        v = String(ctx.strings.add(value))
      }
    }
    const tAttr = t ? ` t="${t}"` : ""
    if (formula) {
      const spill = ctx.spills.get(`${row},${col}`)
      const spillError = isError(value) && value.error === "#SPILL!"
      const own = formula.includes("[") ? (tableAt(sheet, row, col)?.name ?? null) : null
      const text = escapeXmlText(formulaToFile(formula.includes("[") ? structuredToFile(formula, own) : formula))
      if (spill || spillError) {
        dynamic = true
        const area = spill ? rangeRef(spill) : r
        return `<c r="${r}"${sAttr}${tAttr} cm="1"><f t="array" ref="${area}">${text}</f>${v !== "" ? `<v>${v}</v>` : ""}</c>`
      }
      return `<c r="${r}"${sAttr}${tAttr}><f>${text}</f>${v !== "" ? `<v>${v}</v>` : ""}</c>`
    }
    if (value === undefined) return `<c r="${r}"${sAttr}/>`
    return `<c r="${r}"${sAttr}${tAttr}><v>${v}</v></c>`
  }

  const generated: Record<string, string> = {
    dimension: `<dimension ref="${bounds ? rangeRef({ r0: bounds.r0, c0: bounds.c0, r1: bounds.r1, c1: bounds.c1 }) : "A1"}"/>`,
    sheetViews: viewsXml(ctx.view),
    cols: colsXml(sheet, xf, ctx.source?.colExtras),
    sheetData: rowParts.length ? `<sheetData>${rowParts.join("")}</sheetData>` : `<sheetData/>`,
    sheetProtection: protectionXml(sheet),
    autoFilter: sheet.autoFilter ? autoFilterXml(sheet.autoFilter, styles, rangeRef) : "",
    mergeCells: sheet.merges.size
      ? `<mergeCells count="${sheet.merges.size}">${sheet.merges.all.map((m) => `<mergeCell ref="${rangeRef(m)}"/>`).join("")}</mergeCells>`
      : "",
    conditionalFormatting: conditionalXml(sheet.conditional, styles, ctx.source?.keptConditional ?? []),
    dataValidations: validationsXml(sheet.validations, rangeRef),
    drawing: ctx.drawing ? `<drawing r:id="${ctx.drawing}"/>` : "",
    legacyDrawing: ctx.legacyDrawing ? `<legacyDrawing r:id="${ctx.legacyDrawing}"/>` : "",
    ...printElementsXml(sheet.pageSetup, ctx.source?.pageExtras),
    tableParts: ctx.tableParts?.length
      ? `<tableParts count="${ctx.tableParts.length}">${ctx.tableParts.map((id) => `<tablePart r:id="${id}"/>`).join("")}</tableParts>`
      : "",
  }
  const kept: Kept[] = ctx.source?.kept.map((k) => ({ ...k })) ?? []
  const sheetPr = kept.find((k) => k.local === "sheetPr")
  const tab = sheet.tabColor ? `<tabColor rgb="${argb(sheet.tabColor)}"/>` : ""
  if (sheetPr) sheetPr.xml = withTabColor(sheetPr.xml, tab)
  else if (tab) kept.unshift({ local: "sheetPr", xml: `<sheetPr>${tab}</sheetPr>` })
  // Fit to pages is turned on in sheetPr.
  const fitted = withFitToPage(sheetPr?.xml ?? null, sheet.pageSetup.fitTo !== null)
  if (sheetPr && fitted) sheetPr.xml = fitted
  else if (fitted) kept.unshift({ local: "sheetPr", xml: fitted })
  if (!ctx.source) kept.push({ local: "sheetFormatPr", xml: `<sheetFormatPr defaultRowHeight="15"/>` })
  const validationExt = validationExtXml(sheet.validations, rangeRef)
  if (validationExt) {
    const extLst = kept.find((k) => k.local === "extLst")
    if (extLst) extLst.xml = withExt(extLst.xml, validationExt)
    else kept.push({ local: "extLst", xml: withExt(undefined, validationExt) })
  }
  // Generated elements replace nothing (the reader dropped them); assemble puts them in schema order among the kept ones.
  const body = assemble(WORKSHEET_ORDER, kept, generated)
  const open = ctx.source?.rootTag ?? `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"`
  const rootTag = /xmlns:r=/.test(open) ? open : open + ` xmlns:r="${NS_R}"`
  return { xml: `${rootTag}>${body}</worksheet>`, dynamic }
}

/** Rows with a height, hidden flag or style of their own, ascending. */
function specialRows(sheet: Sheet, extras: ReadonlyMap<number, string> | undefined): number[] {
  const out = new Set<number>()
  for (const [row] of sheet.rowStyles.entries()) out.add(row)
  for (const row of extras?.keys() ?? []) out.add(row)
  const rows = sheet.rows
  for (let r = 0; r < MAX_ROWS; r++) if (rows.hasCustomSize(r) || rows.isHidden(r)) out.add(r)
  return [...out].sort((a, b) => a - b)
}

function withTabColor(xml: string, tab: string): string {
  let out = xml
    .replace(/<(\w+:)?tabColor\b[^>]*\/>/, "")
    .replace(/<(\w+:)?tabColor\b[^>]*>[\s\S]*?<\/(\w+:)?tabColor>/, "")
  if (!tab) return out
  if (/\/>$/.test(out) && !/<\w[^>]*>[\s\S]*</.test(out.slice(1))) return out.replace(/\s*\/>$/, `>${tab}</sheetPr>`)
  out = out.replace(/^(<(?:\w+:)?sheetPr\b[^>]*>)/, `$1${tab}`)
  return out
}

function viewsXml(view: SheetView): string {
  const cell = (p: { row: number; col: number }) => ref(p.row, p.col)
  const attrs = attrText({
    tabSelected: view.selected,
    showGridLines: view.showGridlines ? undefined : "0",
    showRowColHeaders: view.showHeadings ? undefined : "0",
    zoomScale: view.zoom !== 1 ? Math.round(view.zoom * 100) : undefined,
    zoomScaleNormal: view.zoom !== 1 ? Math.round(view.zoom * 100) : undefined,
    topLeftCell: view.freeze
      ? view.freeze.top || view.freeze.left
        ? cell({ row: view.freeze.top, col: view.freeze.left })
        : undefined
      : view.topLeft.row || view.topLeft.col
        ? cell(view.topLeft)
        : undefined,
    workbookViewId: 0,
  })
  const sqref = view.selection.map(rangeRef).join(" ")
  const selAttrs = (pane?: string) =>
    attrText({
      pane,
      activeCell: cell(view.active),
      sqref: sqref !== "A1" ? sqref : undefined,
    })
  let inner = ""
  const f = view.freeze
  if (f && (f.row > f.top || f.col > f.left)) {
    const xs = f.col - f.left
    const ys = f.row - f.top
    const pane = xs && ys ? "bottomRight" : ys ? "bottomLeft" : "topRight"
    const topLeft = { row: Math.max(view.topLeft.row, f.row), col: Math.max(view.topLeft.col, f.col) }
    inner += `<pane${attrText({ xSplit: xs || undefined, ySplit: ys || undefined, topLeftCell: cell(topLeft), activePane: pane, state: "frozen" })}/>`
    if (xs && ys) inner += `<selection pane="topRight"/><selection pane="bottomLeft"/>`
    inner += `<selection${selAttrs(pane)}/>`
  } else if (view.active.row || view.active.col || sqref !== "A1") inner += `<selection${selAttrs()}/>`
  return `<sheetViews><sheetView${attrs}${inner ? `>${inner}</sheetView>` : "/>"}</sheetViews>`
}

function colsXml(sheet: Sheet, xf: (id: number) => number, extras: ReadonlyMap<number, string> | undefined): string {
  const cols = sheet.cols
  const parts: string[] = []
  const used = new Set<number>(sheet.colStyles.entries().map(([c]) => c))
  const candidates: number[] = []
  for (let c = 0; c < MAX_COLS; c++)
    if (cols.hasCustomSize(c) || cols.isHidden(c) || used.has(c) || extras?.has(c)) candidates.push(c)
  let i = 0
  while (i < candidates.length) {
    const c = candidates[i]
    const key = (k: number) =>
      `${cols.hasCustomSize(k) ? cols.storedSize(k) : ""}|${cols.isHidden(k)}|${sheet.colStyles.get(k)}|${extras?.get(k) ?? ""}`
    let j = i
    while (j + 1 < candidates.length && candidates[j + 1] === candidates[j] + 1 && key(candidates[j + 1]) === key(c))
      j++
    const style = sheet.colStyles.get(c)
    const width = pxToColWidth(cols.storedSize(c))
    parts.push(
      `<col${attrText({
        min: c + 1,
        max: candidates[j] + 1,
        width,
        style: style ? xf(style) : undefined,
        hidden: cols.isHidden(c),
        customWidth: cols.hasCustomSize(c),
      })}${extras?.get(c) ?? ""}/>`,
    )
    i = j + 1
  }
  return parts.length ? `<cols>${parts.join("")}</cols>` : ""
}

function protectionXml(sheet: Sheet): string {
  const p = sheet.protection
  if (!p) return ""
  const attrs: Record<string, string | undefined> = {}
  if (p.hash) {
    attrs.algorithmName = p.hash.algorithm
    attrs.hashValue = p.hash.hash
    attrs.saltValue = p.hash.salt
    attrs.spinCount = String(p.hash.spinCount)
  } else if (p.password) attrs.password = p.password
  attrs.sheet = "1"
  for (const f of PROTECTION_FLAGS) {
    const allowed = p[f.key] as boolean
    const locked = !allowed
    if (locked !== f.lockedByDefault) attrs[f.attr] = locked ? "1" : "0"
  }
  return `<sheetProtection${attrText(attrs)}/>`
}

// ── Conditional formatting ──────────────────────────────────────────────────────────────────────────────────

const DATE_FORMULAS: Record<string, (a: string) => string> = {
  today: (a) => `FLOOR(${a},1)=TODAY()`,
  yesterday: (a) => `FLOOR(${a},1)=TODAY()-1`,
  tomorrow: (a) => `FLOOR(${a},1)=TODAY()+1`,
  last7Days: (a) => `AND(TODAY()-FLOOR(${a},1)<=6,FLOOR(${a},1)<=TODAY())`,
  thisMonth: (a) => `AND(MONTH(${a})=MONTH(TODAY()),YEAR(${a})=YEAR(TODAY()))`,
  lastMonth: (a) => `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0-1)),YEAR(${a})=YEAR(EDATE(TODAY(),0-1)))`,
  nextMonth: (a) => `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0+1)),YEAR(${a})=YEAR(EDATE(TODAY(),0+1)))`,
  thisWeek: (a) => `AND(TODAY()-ROUNDDOWN(${a},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${a},0)-TODAY()<=7-WEEKDAY(TODAY()))`,
  lastWeek: (a) => `AND(TODAY()-ROUNDDOWN(${a},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${a},0)<(WEEKDAY(TODAY())+7))`,
  nextWeek: (a) =>
    `AND(ROUNDDOWN(${a},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${a},0)-TODAY()<(15-WEEKDAY(TODAY())))`,
}

const quote = (text: string) => `"${text.replace(/"/g, '""')}"`
const f = (text: string) => `<formula>${escapeXmlText(formulaToFile("=" + text))}</formula>`

function cfvo(t: CFThreshold): string {
  const type = t.type === "automatic" ? "min" : t.type
  const val =
    type === "min" || type === "max"
      ? undefined
      : type === "formula"
        ? formulaToFile("=" + (t.value ?? "0"))
        : (t.value ?? "0")
  return `<cfvo${attrText({ type, val })}/>`
}

function ruleXml(
  rule: CFRule,
  at: string,
  styles: StylesOut,
): { attrs: Record<string, string | number | boolean | undefined>; body: string } {
  const dxf = "style" in rule && rule.style ? styles.dxf(rule.style) : undefined
  switch (rule.kind) {
    case "cellIs":
      return { attrs: { type: "cellIs", dxfId: dxf, operator: rule.operator }, body: rule.values.map(f).join("") }
    case "text": {
      const q = quote(rule.text)
      const map = {
        contains: ["containsText", "containsText", `NOT(ISERROR(SEARCH(${q},${at})))`],
        notContains: ["notContainsText", "notContains", `ISERROR(SEARCH(${q},${at}))`],
        beginsWith: ["beginsWith", "beginsWith", `LEFT(${at},LEN(${q}))=${q}`],
        endsWith: ["endsWith", "endsWith", `RIGHT(${at},LEN(${q}))=${q}`],
      } as const
      const [type, operator, formula] = map[rule.operator]
      return { attrs: { type, dxfId: dxf, operator, text: rule.text }, body: f(formula) }
    }
    case "date":
      return {
        attrs: { type: "timePeriod", dxfId: dxf, timePeriod: rule.period },
        body: f(DATE_FORMULAS[rule.period](at)),
      }
    case "duplicate":
      return { attrs: { type: rule.unique ? "uniqueValues" : "duplicateValues", dxfId: dxf }, body: "" }
    case "top":
      return {
        attrs: { type: "top10", dxfId: dxf, percent: rule.percent, bottom: rule.bottom, rank: rule.rank },
        body: "",
      }
    case "average":
      return {
        attrs: {
          type: "aboveAverage",
          dxfId: dxf,
          aboveAverage: rule.below ? "0" : undefined,
          equalAverage: rule.equal,
        },
        body: "",
      }
    case "blanks":
      return rule.blank
        ? { attrs: { type: "containsBlanks", dxfId: dxf }, body: f(`LEN(TRIM(${at}))=0`) }
        : { attrs: { type: "notContainsBlanks", dxfId: dxf }, body: f(`LEN(TRIM(${at}))>0`) }
    case "errors":
      return rule.errors
        ? { attrs: { type: "containsErrors", dxfId: dxf }, body: f(`ISERROR(${at})`) }
        : { attrs: { type: "notContainsErrors", dxfId: dxf }, body: f(`NOT(ISERROR(${at}))`) }
    case "expression":
      return { attrs: { type: "expression", dxfId: dxf }, body: f(rule.formula) }
    case "colorScale":
      return {
        attrs: { type: "colorScale" },
        body:
          `<colorScale>${rule.points.map((p) => cfvo(p.at)).join("")}` +
          `${rule.points.map((p) => `<color rgb="${argb(p.color)}"/>`).join("")}</colorScale>`,
      }
    case "dataBar":
      return {
        attrs: { type: "dataBar" },
        body: `<dataBar>${cfvo(rule.min)}${cfvo(rule.max)}<color rgb="${argb(rule.color)}"/></dataBar>`,
      }
    case "iconSet":
      return {
        attrs: { type: "iconSet" },
        body:
          `<iconSet${attrText({ iconSet: rule.set, showValue: rule.showValue ? undefined : "0", reverse: rule.reverse })}>` +
          `<cfvo type="percent" val="0"/>${rule.thresholds.map(cfvo).join("")}</iconSet>`,
      }
  }
}

function conditionalXml(entries: readonly CFEntry[], styles: StylesOut, kept: readonly string[]): string {
  let out = ""
  let priority = 1
  for (const e of entries) {
    if (!e.ranges.length) continue
    const first = e.ranges[0]
    const at = ref(first.r0, first.c0)
    const { attrs, body } = ruleXml(e.rule, at, styles)
    const rule = `<cfRule${attrText({ ...attrs, priority: priority++, stopIfTrue: e.stopIfTrue })}${body ? `>${body}</cfRule>` : "/>"}`
    out += `<conditionalFormatting sqref="${escapeXml(e.ranges.map(rangeRef).join(" "))}">${rule}</conditionalFormatting>`
  }
  for (const k of kept)
    out += k.replace(/(<(?:\w+:)?cfRule\b[^>]*?\bpriority=")\d+"/g, (_, head: string) => `${head}${priority++}"`)
  return out
}
