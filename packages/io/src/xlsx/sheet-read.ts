import {
  copyFormula,
  DEFAULT_PROTECTION,
  ERROR_CODES,
  errorValue,
  ICON_SETS,
  MAX_COLS,
  MAX_ROWS,
  parseRange,
  type CellData,
  type CellValue,
  type CFEntry,
  type CFRule,
  type CFThreshold,
  type CellIsOperator,
  type ErrorCode,
  type IconSetName,
  type PasswordHash,
  type Range,
  type Sheet,
  type SheetProtection,
  type TimePeriod,
} from "@cells/core"
import { escapeXml } from "../xml"
import { readAutoFilter } from "./autofilter"
import { scanXml } from "./scan"
import { at, kid, kids, parseFragment, raw, type Kept, type XDoc, type XNode } from "./dom"
import { formulaFromFile } from "./formula"
import type { StylesIn } from "./styles"
import { readValidationExt, readValidations } from "./validation"
import { colWidthToPx, pointsToPx } from "./units"

/** What a sheet's view looked like when saved (Excel keeps these per sheet). */
export type SheetView = {
  /** Frozen panes as Cells keeps them: first unfrozen row and column, and the top-left cell when frozen. */
  freeze: { row: number; col: number; top: number; left: number } | null
  zoom: number
  showGridlines: boolean
  showHeadings: boolean
  topLeft: { row: number; col: number }
  active: { row: number; col: number }
  selection: Range[]
  selected: boolean
}

export const DEFAULT_VIEW: SheetView = {
  freeze: null,
  zoom: 1,
  showGridlines: true,
  showHeadings: true,
  topLeft: { row: 0, col: 0 },
  active: { row: 0, col: 0 },
  selection: [{ r0: 0, c0: 0, r1: 0, c1: 0 }],
  selected: false,
}

export type SheetReadContext = {
  sheet: Sheet
  styles: StylesIn | null
  /** Called between slices of a big sheet with the share of it read; awaited, so the caller can let the UI draw. */
  progress?: (done: number) => void | Promise<void>
  /** Cells style id for a cell format (xf) index of the file. */
  styleId: (xf: number) => number
  strings: readonly string[]
  warnings: string[]
}

export type SheetRead = {
  view: SheetView
  /** The worksheet's start tag (namespace declarations, mc:Ignorable), without ">". */
  rootTag: string
  /** Top-level elements written back as read (Cells regenerates the rest). */
  kept: Kept[]
  /** The vmlDrawing relationship of the sheet's notes (legacyDrawing), dropped and written anew. */
  legacyDrawing: string | null
  /** The file had dynamic-array formulas: their spilled cells need calculating. */
  dynamic: boolean
  /** Conditional formats Cells doesn't model, as `conditionalFormatting` elements to write back as read. */
  keptConditional: string[]
  /** Row and column attributes Cells doesn't model (outline levels, collapsed, dyDescent), as read, by index. */
  rowExtras: Map<number, string>
  colExtras: Map<number, string>
}

/** Row and column attributes Cells writes itself; the others are kept. */
const ROW_OWN = new Set(["r", "spans", "s", "customFormat", "ht", "customHeight", "hidden"])
const COL_OWN = new Set(["min", "max", "width", "style", "hidden", "customWidth"])

/** Elements Cells reads and writes itself; every other child of `worksheet` is kept as read. */
const OWN = new Set([
  "dimension",
  "sheetViews",
  "cols",
  "sheetData",
  "sheetProtection",
  "autoFilter",
  "mergeCells",
  "conditionalFormatting",
  "dataValidations",
  "legacyDrawing",
])

/** Unicode escapes Excel writes in text (`_x000D_`). */
export const unescapeText = (s: string) =>
  s.indexOf("_x") < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)))

/** "AB12" → row 11, col 27 (no allocation; -1 when it isn't a cell). */
function decodeRef(ref: string): { row: number; col: number } | null {
  let col = 0
  let i = 0
  for (; i < ref.length; i++) {
    const c = ref.charCodeAt(i)
    if (c >= 65 && c <= 90) col = col * 26 + (c - 64)
    else if (c >= 97 && c <= 122) col = col * 26 + (c - 96)
    else if (c !== 36) break
  }
  let row = 0
  for (; i < ref.length; i++) {
    const c = ref.charCodeAt(i)
    if (c >= 48 && c <= 57) row = row * 10 + (c - 48)
    else if (c !== 36) return null
  }
  if (!col || !row) return null
  return { row: row - 1, col: col - 1 }
}

const ISO_EPOCH = Date.UTC(1899, 11, 30)

function errorOf(text: string, warnings: string[]): CellValue {
  if ((ERROR_CODES as readonly string[]).includes(text)) return errorValue(text as ErrorCode)
  warnings.push(`The error ${text} isn't known to Cells; it shows as #VALUE!.`)
  return errorValue("#VALUE!")
}

/** Reads one worksheet part into `ctx.sheet`. */
export async function readSheet(xml: string, ctx: SheetReadContext): Promise<SheetRead> {
  const { sheet, warnings } = ctx
  const kept: Kept[] = []
  let rootTag = ""
  let view: SheetView = { ...DEFAULT_VIEW }
  let legacyDrawing: string | null = null
  let dynamic = false
  const conditional: XDoc[] = []
  const keptConditional: string[] = []
  const rowExtras = new Map<number, string>()
  const colExtras = new Map<number, string>()

  let depth = 0
  let topStart = 0
  let topLocal = ""
  // sheetData state
  let inData = false
  let row = -1
  let col = -1
  let cellType = ""
  let cellStyle = 0
  let cellMeta = false
  let fAttrs: { t?: string; ref?: string; si?: string } | null = null
  let fText = ""
  let vText = ""
  let isText = ""
  let capture: "v" | "f" | "t" | null = null
  let inIs = false
  const shared = new Map<string, { row: number; col: number; formula: string }>()
  const arrays: Range[] = []
  let lastRow = -1

  const inArray = (r: number, c: number) => {
    for (let i = arrays.length - 1; i >= 0; i--) {
      const a = arrays[i]
      if (r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1 && !(r === a.r0 && c === a.c0)) return true
    }
    return false
  }

  const finishCell = () => {
    const data: CellData = {}
    if (cellStyle) data.style = cellStyle
    let formula: string | undefined
    if (fAttrs) {
      const t = fAttrs.t
      if (t === "shared" && fAttrs.si !== undefined) {
        if (fText) {
          formula = formulaFromFile(fText)
          shared.set(fAttrs.si, { row, col, formula })
        } else {
          const master = shared.get(fAttrs.si)
          if (master) formula = copyFormula(master.formula, row - master.row, col - master.col)
        }
      } else if (t === "dataTable") {
        warnings.push("What-If data tables ({=TABLE()}) aren't supported; their values are kept as numbers.")
      } else if (fText) formula = formulaFromFile(fText)
      if (t === "array" && fAttrs.ref) {
        const r = parseRange(fAttrs.ref)
        if (r && (cellMeta || r.r0 !== r.r1 || r.c0 !== r.c1)) {
          arrays.push(r)
          dynamic = true
        }
      }
    }
    // A spilled cell belongs to its array formula: the calculation fills it.
    if (!formula && arrays.length && inArray(row, col)) {
      if (cellStyle) sheet.cells.set(row, col, { style: cellStyle })
      return
    }
    let value: CellValue | undefined
    switch (cellType) {
      case "s": {
        const s = ctx.strings[Number(vText)]
        value = s ?? ""
        break
      }
      case "str":
        value = unescapeText(vText)
        break
      case "inlineStr":
        value = unescapeText(isText)
        break
      case "b":
        value = vText === "1" || vText === "true"
        break
      case "e":
        value = errorOf(vText, warnings)
        break
      case "d": {
        const ms = Date.parse(vText.endsWith("Z") || /[+-]\d\d:\d\d$/.test(vText) ? vText : vText + "Z")
        value = Number.isNaN(ms) ? vText : (ms - ISO_EPOCH) / 86_400_000
        break
      }
      default:
        if (vText !== "") value = Number(vText)
    }
    if (value !== undefined) data.value = value
    if (formula) {
      data.input = formula
      if (value === undefined) data.value = 0
    } else if (typeof value === "string" && value !== "" && needsApostrophe(value)) {
      // Text that would read as a number or formula when typed keeps a leading apostrophe in the input.
      data.input = "'" + value
    }
    if (data.value !== undefined || data.style) sheet.cells.set(row, col, data)
  }

  const handlers: Parameters<typeof scanXml>[1] = {
    start(tag) {
      depth++
      const local = tag.local
      if (depth === 1) {
        rootTag = xml.slice(tag.start, tag.end).replace(/\/?>$/, "")
        return
      }
      if (depth === 2) {
        topStart = tag.start
        topLocal = local
        if (local === "sheetData") inData = true
        return
      }
      if (!inData) return
      const attr = (name: string) => tag.attrs.get(name)
      if (local === "row") {
        const r = attr("r")
        row = r ? Number(r) - 1 : lastRow + 1
        lastRow = row
        col = -1
        if (row >= MAX_ROWS) return
        let extra = ""
        if (tag.attrs.size > 3)
          for (const [name, value] of tag.attrs)
            if (!ROW_OWN.has(name) && !name.startsWith("xmlns")) extra += ` ${name}="${escapeXml(value)}"`
        if (extra) rowExtras.set(row, extra)
        const ht = attr("ht")
        const custom = attr("customHeight")
        if (ht !== undefined) {
          const px = pointsToPx(Number(ht))
          if (custom === "1" || custom === "true") sheet.rows.setSize(row, px)
          else if (px !== sheet.rows.size(row)) sheet.rows.setSize(row, px, true)
        }
        const hidden = attr("hidden")
        if (hidden === "1" || hidden === "true") sheet.rows.setHidden(row, true)
        const s = attr("s")
        const cf = attr("customFormat")
        if (s && (cf === "1" || cf === "true")) sheet.rowStyles.set(row, ctx.styleId(Number(s)))
      } else if (local === "c") {
        const r = attr("r")
        const ref = r ? decodeRef(r) : null
        if (ref) {
          row = ref.row
          col = ref.col
        } else col++
        cellType = attr("t") ?? ""
        const s = attr("s")
        cellStyle = s ? ctx.styleId(Number(s)) : 0
        cellMeta = attr("cm") !== undefined
        fAttrs = null
        fText = ""
        vText = ""
        isText = ""
        inIs = false
      } else if (local === "f") {
        fAttrs = { t: attr("t"), ref: attr("ref"), si: attr("si") }
        capture = "f"
      } else if (local === "v") capture = "v"
      else if (local === "is") inIs = true
      else if (local === "t" && inIs) capture = "t"
      else if (local === "rPh") capture = null
    },
    end(local, end) {
      if (depth === 2) {
        if (local === "sheetData") inData = false
        else {
          const source = xml.slice(topStart, end)
          handleTop(topLocal, source)
        }
      } else if (inData) {
        if (local === "c") {
          if (row < MAX_ROWS && col < MAX_COLS && col >= 0) finishCell()
        } else if (local === "f" || local === "v" || local === "t") capture = null
        else if (local === "is") inIs = false
      }
      depth--
    },
    text(text) {
      if (capture === "v") vText += text
      else if (capture === "f") fText += text
      else if (capture === "t") isText += text
    },
  }
  const state = { pos: 0, depth: 0 }
  const SLICE = 2_000_000
  while (!scanXml(xml, handlers, state, ctx.progress ? SLICE : Infinity)) await ctx.progress?.(state.pos / xml.length)

  function handleTop(local: string, source: string) {
    if (!OWN.has(local)) {
      if (local === "sheetPr") {
        const color = kid(parseFragment(source, rootTag).root, "tabColor")
        const rgb = color && ctx.styles ? at(color, "rgb") : undefined
        if (rgb) sheet.tabColor = "#" + rgb.slice(-6).toUpperCase()
      }
      if (local === "extLst") {
        const ext = readValidationExt(dropConditionalExt(source), rootTag)
        sheet.validations.push(...ext.entries)
        source = ext.rest
      }
      if (source) kept.push({ local, xml: source })
      return
    }
    switch (local) {
      case "sheetViews":
        view = readView(parseFragment(source, rootTag).root, warnings)
        break
      case "cols":
        readCols(parseFragment(source, rootTag).root)
        break
      case "mergeCells":
        for (const m of kids(parseFragment(source, rootTag).root, "mergeCell")) {
          const r = parseRange(at(m, "ref") ?? "")
          if (r && (r.r0 !== r.r1 || r.c0 !== r.c1)) sheet.merges.add(r)
        }
        break
      case "sheetProtection":
        sheet.protection = readProtection(parseFragment(source, rootTag).root)
        break
      case "autoFilter":
        sheet.autoFilter = readAutoFilter(parseFragment(source, rootTag).root, ctx.styles, warnings)
        break
      case "conditionalFormatting":
        conditional.push(parseFragment(source, rootTag))
        break
      case "dataValidations":
        sheet.validations.unshift(...readValidations(parseFragment(source, rootTag).root))
        break
      case "legacyDrawing":
        legacyDrawing =
          at(
            parseFragment(source, rootTag).root,
            "id",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          ) ?? null
        break
    }
  }

  function readCols(root: XNode) {
    for (const c of kids(root, "col")) {
      const min = Number(at(c, "min") ?? 0) - 1
      const max = Math.min(MAX_COLS, Number(at(c, "max") ?? 0)) - 1
      if (min < 0 || max < min) continue
      const width = at(c, "width")
      const px = width !== undefined ? colWidthToPx(Number(width)) : null
      const hidden = at(c, "hidden") === "1" || at(c, "hidden") === "true"
      const s = at(c, "style")
      const style = s ? ctx.styleId(Number(s)) : 0
      let extra = ""
      for (const [name, value] of Object.entries(c.attrs))
        if (!COL_OWN.has(name)) extra += ` ${name}="${escapeXml(value)}"`
      for (let i = min; i <= max; i++) {
        if (extra) colExtras.set(i, extra)
        if (px !== null && px !== sheet.cols.size(i)) sheet.cols.setSize(i, px)
        if (hidden) sheet.cols.setHidden(i, true)
        if (style) sheet.colStyles.set(i, style)
      }
    }
  }

  // Conditional formats: rules in priority order across the sheet.
  const rules: { priority: number; entry: CFEntry }[] = []
  for (const doc of conditional) {
    const cf = doc.root
    const ranges = (at(cf, "sqref") ?? "")
      .split(/\s+/)
      .map((r) => parseRange(r))
      .filter((r): r is Range => r !== null)
    if (!ranges.length) continue
    const unknown: string[] = []
    for (const r of kids(cf, "cfRule")) {
      const rule = readRule(r, ctx.styles, warnings)
      if (!rule) {
        unknown.push(raw(doc, r))
        continue
      }
      rules.push({
        priority: Number(at(r, "priority") ?? 0),
        entry: { id: 0, ranges, rule, stopIfTrue: at(r, "stopIfTrue") === "1" || at(r, "stopIfTrue") === "true" },
      })
    }
    if (unknown.length)
      keptConditional.push(doc.source.slice(cf.start, cf.tagEnd) + unknown.join("") + "</" + cf.name + ">")
  }
  // Excel groups the elements by range; ids follow priority, the order Cells keeps them in.
  sheet.conditional = rules.sort((a, b) => a.priority - b.priority).map((r, i) => ({ ...r.entry, id: i + 1 }))

  return { view, rootTag, kept, legacyDrawing, dynamic, keptConditional, rowExtras, colExtras }
}

/** Text that the entry parser would turn into something else ("123", "=1", "TRUE") keeps an apostrophe. */
function needsApostrophe(text: string): boolean {
  if (/^[=+-]/.test(text)) return true
  if (/^\s*[-+]?[$€£¥]?\s*[\d.,]+%?\s*$/.test(text)) return true
  return /^(true|false)$/i.test(text)
}

/** Drops the x14 conditional formats from extLst: Cells writes its own rules, and Excel would show both. */
function dropConditionalExt(source: string): string {
  const out = source.replace(/<(\w+:)?ext\b[^>]*\{78C0D931-6437-407d-A8EE-F0AAD8E25EB1\}[\s\S]*?<\/\1?ext>/gi, "")
  return /<(\w+:)?ext\b/.test(out) ? out : ""
}

function readView(root: XNode, warnings: string[]): SheetView {
  const v = kid(root, "sheetView")
  if (!v) return { ...DEFAULT_VIEW }
  const cell = (text: string | undefined) => {
    const r = text ? decodeRef(text) : null
    return r ?? { row: 0, col: 0 }
  }
  const topLeft = cell(at(v, "topLeftCell"))
  const pane = kid(v, "pane")
  let freeze: SheetView["freeze"] = null
  if (pane) {
    const state = at(pane, "state")
    const xs = Math.round(Number(at(pane, "xSplit") ?? 0))
    const ys = Math.round(Number(at(pane, "ySplit") ?? 0))
    if (state === "frozen" || state === "frozenSplit")
      freeze = { row: topLeft.row + ys, col: topLeft.col + xs, top: topLeft.row, left: topLeft.col }
    else warnings.push("Split panes aren't supported yet; the sheet opens without them.")
  }
  const selections = kids(v, "selection")
  const sel = selections.find((s) => at(s, "pane") === at(pane, "activePane")) ?? selections[selections.length - 1]
  const active = cell(at(sel, "activeCell"))
  const ranges = (at(sel, "sqref") ?? "")
    .split(/\s+/)
    .map((r) => parseRange(r))
    .filter((r): r is Range => r !== null)
  const zoom = Number(at(v, "zoomScale") ?? 100)
  return {
    freeze,
    zoom: zoom > 0 ? zoom / 100 : 1,
    showGridlines: at(v, "showGridLines") !== "0" && at(v, "showGridLines") !== "false",
    showHeadings: at(v, "showRowColHeaders") !== "0" && at(v, "showRowColHeaders") !== "false",
    topLeft: freeze
      ? {
          row: Math.max(freeze.row, cell(at(pane, "topLeftCell")).row),
          col: Math.max(freeze.col, cell(at(pane, "topLeftCell")).col),
        }
      : topLeft,
    active,
    selection: ranges.length ? ranges : [{ r0: active.row, c0: active.col, r1: active.row, c1: active.col }],
    selected: at(v, "tabSelected") === "1" || at(v, "tabSelected") === "true",
  }
}

/**
 * Protect Sheet's options. In the file most flags say what is locked (formatCells="0" allows formatting); the two
 * selection flags and objects/scenarios say what is protected when "1".
 */
export const PROTECTION_FLAGS: { key: keyof SheetProtection; attr: string; lockedByDefault: boolean }[] = [
  { key: "editObjects", attr: "objects", lockedByDefault: false },
  { key: "editScenarios", attr: "scenarios", lockedByDefault: false },
  { key: "formatCells", attr: "formatCells", lockedByDefault: true },
  { key: "formatColumns", attr: "formatColumns", lockedByDefault: true },
  { key: "formatRows", attr: "formatRows", lockedByDefault: true },
  { key: "insertColumns", attr: "insertColumns", lockedByDefault: true },
  { key: "insertRows", attr: "insertRows", lockedByDefault: true },
  { key: "insertHyperlinks", attr: "insertHyperlinks", lockedByDefault: true },
  { key: "deleteColumns", attr: "deleteColumns", lockedByDefault: true },
  { key: "deleteRows", attr: "deleteRows", lockedByDefault: true },
  { key: "selectLocked", attr: "selectLockedCells", lockedByDefault: false },
  { key: "sort", attr: "sort", lockedByDefault: true },
  { key: "autoFilter", attr: "autoFilter", lockedByDefault: true },
  { key: "pivotTables", attr: "pivotTables", lockedByDefault: true },
  { key: "selectUnlocked", attr: "selectUnlockedCells", lockedByDefault: false },
]

export function readPasswordHash(n: XNode, prefix = ""): PasswordHash | undefined {
  const algorithm = at(n, prefix ? prefix + "AlgorithmName" : "algorithmName")
  const hash = at(n, prefix ? prefix + "HashValue" : "hashValue")
  if (!algorithm || !hash) return undefined
  return {
    algorithm,
    hash,
    salt: at(n, prefix ? prefix + "SaltValue" : "saltValue") ?? "",
    spinCount: Number(at(n, prefix ? prefix + "SpinCount" : "spinCount") ?? 0),
  }
}

function readProtection(n: XNode): SheetProtection | null {
  if (at(n, "sheet") !== "1" && at(n, "sheet") !== "true") return null
  const p: SheetProtection = { ...DEFAULT_PROTECTION }
  delete p.password
  for (const f of PROTECTION_FLAGS) {
    const v = at(n, f.attr)
    const locked = v === undefined ? f.lockedByDefault : v === "1" || v === "true"
    ;(p as Record<string, unknown>)[f.key] = !locked
  }
  const password = at(n, "password")
  if (password) p.password = password.toUpperCase()
  const hash = readPasswordHash(n)
  if (hash) p.hash = hash
  return p
}

const PERIODS = new Set([
  "yesterday",
  "today",
  "tomorrow",
  "last7Days",
  "lastWeek",
  "thisWeek",
  "nextWeek",
  "lastMonth",
  "thisMonth",
  "nextMonth",
])

function threshold(n: XNode): CFThreshold {
  const type = at(n, "type") ?? "min"
  const value = at(n, "val")
  const t = (
    ["min", "max", "num", "percent", "percentile", "formula"].includes(type)
      ? type
      : type === "autoMin"
        ? "min"
        : type === "autoMax"
          ? "max"
          : "num"
  ) as CFThreshold["type"]
  return value !== undefined && t !== "min" && t !== "max"
    ? { type: t, value: t === "formula" ? formulaFromFile(value).slice(1) : value }
    : { type: t }
}

function readRule(r: XNode, styles: StylesIn | null, warnings: string[]): CFRule | null {
  const type = at(r, "type")
  const dxf = at(r, "dxfId")
  const style = dxf !== undefined && styles ? styles.dxfStyle(Number(dxf)) : {}
  const formulas = kids(r, "formula").map((f) => formulaFromFile(f.text).slice(1))
  const truthy = (v: string | undefined, dflt: boolean) => (v === undefined ? dflt : v === "1" || v === "true")
  switch (type) {
    case "cellIs":
      return { kind: "cellIs", operator: (at(r, "operator") ?? "equal") as CellIsOperator, values: formulas, style }
    case "containsText":
      return { kind: "text", operator: "contains", text: at(r, "text") ?? "", style }
    case "notContainsText":
      return { kind: "text", operator: "notContains", text: at(r, "text") ?? "", style }
    case "beginsWith":
      return { kind: "text", operator: "beginsWith", text: at(r, "text") ?? "", style }
    case "endsWith":
      return { kind: "text", operator: "endsWith", text: at(r, "text") ?? "", style }
    case "timePeriod": {
      const period = at(r, "timePeriod") ?? ""
      return PERIODS.has(period) ? { kind: "date", period: period as TimePeriod, style } : null
    }
    case "duplicateValues":
    case "uniqueValues":
      return { kind: "duplicate", unique: type === "uniqueValues", style }
    case "top10":
      return {
        kind: "top",
        bottom: truthy(at(r, "bottom"), false),
        percent: truthy(at(r, "percent"), false),
        rank: Number(at(r, "rank") ?? 10),
        style,
      }
    case "aboveAverage":
      if (at(r, "stdDev")) {
        return null
      }
      return {
        kind: "average",
        below: !truthy(at(r, "aboveAverage"), true),
        equal: truthy(at(r, "equalAverage"), false),
        style,
      }
    case "containsBlanks":
    case "notContainsBlanks":
      return { kind: "blanks", blank: type === "containsBlanks", style }
    case "containsErrors":
    case "notContainsErrors":
      return { kind: "errors", errors: type === "containsErrors", style }
    case "expression":
      return formulas[0] ? { kind: "expression", formula: formulas[0], style } : null
    case "colorScale": {
      const cs = kid(r, "colorScale")
      const cfvos = kids(cs, "cfvo")
      const colors = kids(cs, "color")
      const points = cfvos.map((c, i) => ({ at: threshold(c), color: colorOf(colors[i], styles) ?? "#FFFFFF" }))
      return points.length >= 2 ? { kind: "colorScale", points } : null
    }
    case "dataBar": {
      const db = kid(r, "dataBar")
      const cfvos = kids(db, "cfvo")
      return {
        kind: "dataBar",
        min: cfvos[0] ? threshold(cfvos[0]) : { type: "min" },
        max: cfvos[1] ? threshold(cfvos[1]) : { type: "max" },
        color: colorOf(kid(db, "color"), styles) ?? "#638EC6",
        gradient: true,
      }
    }
    case "iconSet": {
      const is = kid(r, "iconSet")
      const name = at(is, "iconSet") ?? "3TrafficLights1"
      if (!(name in ICON_SETS)) return null
      return {
        kind: "iconSet",
        set: name as IconSetName,
        reverse: truthy(at(is, "reverse"), false),
        showValue: truthy(at(is, "showValue"), true),
        thresholds: kids(is, "cfvo").slice(1).map(threshold),
      }
    }
  }
  // Kept as read: Excel shows it, Cells doesn't draw it.
  warnings.push(`A conditional format of type ${type} isn't shown in Cells; it is kept in the file.`)
  return null
}

function colorOf(n: XNode | undefined, styles: StylesIn | null): string | null {
  if (!n) return null
  const rgb = at(n, "rgb")
  if (rgb) return "#" + rgb.slice(-6).toUpperCase()
  if (styles) {
    const theme = at(n, "theme")
    if (theme !== undefined) return "#" + (styles.theme[Number(theme)] ?? "000000").toUpperCase()
  }
  return null
}
