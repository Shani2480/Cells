import type { PdfItem } from "@cells/io"
import {
  copyTable,
  findTable,
  nextTableId,
  queryFormula,
  readQueryPlan,
  type CellWrite,
  type Op,
  type QueryColumnType,
  type QueryPlan,
  type QueryStep,
  type QueryValue,
  EN_US,
  regional,
  stepName,
  type Regional,
  type Table,
  type Workbook,
  type WorkbookQuery,
} from "@cells/core"

/**
 * Data › Get Data, run by Cells: a query's plan (see core's power-query.ts) evaluated step by step, and its result
 * loaded into a table on a sheet of its own, as Excel's Close & Load does; Refresh runs it again into the same table.
 */

export type QueryResult = { columns: string[]; rows: QueryValue[][]; types: (QueryColumnType | null)[] }

/** What a query can read: this workbook's tables, files on the web, and a PDF's tables (read from its path). */
export type QueryContext = {
  wb: Workbook
  fetch(url: string): Promise<{ bytes: Uint8Array; type: string }>
  pdf(path: string): Promise<PdfItem[]>
}

/** A source with several tables: a web page's (by place), or a PDF's (by Id). */
export type SourceItem = QueryResult & { id?: string }

/** A PDF's table or page as Pdf.Tables gives it: text columns Column1, Column2…, empty cells null. */
export function pdfResult(item: PdfItem): SourceItem {
  const width = item.rows[0]?.length ?? 0
  return {
    ...result(
      columnNames(width),
      item.rows.map((r) => r.map((v) => (v === "" ? null : v))),
    ),
    id: item.id,
  }
}

export class QueryError extends Error {}

const columnNames = (n: number) => Array.from({ length: n }, (_, i) => `Column${i + 1}`)

function result(columns: string[], rows: QueryValue[][]): QueryResult {
  return { columns, rows, types: columns.map(() => null) }
}

/** A workbook table's header and data as a query sees them (errors as null). */
export function tableResult(wb: Workbook, name: string): QueryResult | null {
  const found = findTable(wb.sheets, name)
  if (!found) return null
  const { table } = found
  const sheet = wb.sheets[found.sheet]
  const r = table.range
  const first = table.header ? r.r0 + 1 : r.r0
  const last = table.totals ? r.r1 - 1 : r.r1
  const rows: QueryValue[][] = []
  for (let row = first; row <= last; row++) {
    const line: QueryValue[] = []
    for (let col = r.c0; col <= r.c1; col++) {
      const v = sheet.cells.value(row, col)
      line.push(v === undefined || v === null || typeof v === "object" ? null : (v as QueryValue))
    }
    rows.push(line)
  }
  return result(
    table.columns.map((c) => c.name),
    rows,
  )
}

/** CSV text into rows (quotes as RFC 4180 has them). */
export function parseCsv(text: string, delimiter: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let cell = ""
  let quoted = false
  for (let i = 0; i < text.length; i++) {
    const c = text[i]
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') {
        cell += '"'
        i++
      } else if (c === '"') quoted = false
      else cell += c
    } else if (c === '"' && cell === "") quoted = true
    else if (c === delimiter) {
      row.push(cell)
      cell = ""
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++
      row.push(cell)
      rows.push(row)
      row = []
      cell = ""
    } else cell += c
  }
  if (cell !== "" || row.length) {
    row.push(cell)
    rows.push(row)
  }
  return rows
}

const ENCODINGS: Record<number, string> = { 65001: "utf-8", 1252: "windows-1252", 1200: "utf-16le", 1201: "utf-16be" }

function decode(bytes: Uint8Array, codePage: number): string {
  try {
    return new TextDecoder(ENCODINGS[codePage] ?? "utf-8").decode(bytes).replace(/^\uFEFF/, "")
  } catch {
    return new TextDecoder().decode(bytes)
  }
}

/** A web page's tables, as Web.Page reads them: a header row of th cells names the columns. */
export function htmlTables(html: string): QueryResult[] {
  const doc = new DOMParser().parseFromString(html, "text/html")
  return [...doc.querySelectorAll("table")].map((table) => {
    const lines = [...table.querySelectorAll("tr")].filter((tr) => tr.closest("table") === table)
    const cells = lines.map((tr) => [...tr.children].filter((c) => /^T[HD]$/.test(c.tagName)))
    const header = cells[0]?.length && cells[0].every((c) => c.tagName === "TH")
    const body = (header ? cells.slice(1) : cells).map((r) => r.map((c) => (c.textContent ?? "").trim()))
    const width = Math.max(header ? cells[0].length : 0, ...body.map((r) => r.length))
    const names = header ? uniqueNames(cells[0].map((c) => (c.textContent ?? "").trim())) : columnNames(width)
    while (names.length < width) names.push(`Column${names.length + 1}`)
    return result(
      names,
      body.map((r) => names.map((_, i) => (r[i] === undefined || r[i] === "" ? null : r[i]))),
    )
  })
}

/** JSON's list of records as a table: a column per key, in the order keys first appear. */
export function jsonTable(text: string): QueryResult {
  let data: unknown
  try {
    data = JSON.parse(text)
  } catch {
    throw new QueryError("The file isn't JSON.")
  }
  const list = Array.isArray(data)
    ? data
    : data && typeof data === "object"
      ? Object.values(data).find(Array.isArray)
      : null
  if (!Array.isArray(list) || !list.every((r) => r && typeof r === "object" && !Array.isArray(r)))
    throw new QueryError("The JSON isn't a list of records, so it can't be a table.")
  const columns: string[] = []
  for (const r of list) for (const k of Object.keys(r as object)) if (!columns.includes(k)) columns.push(k)
  const value = (v: unknown): QueryValue =>
    v === null || v === undefined
      ? null
      : typeof v === "string" || typeof v === "number" || typeof v === "boolean"
        ? v
        : JSON.stringify(v)
  return result(
    columns,
    list.map((r) => columns.map((k) => value((r as Record<string, unknown>)[k]))),
  )
}

/** Header names made unique as Power Query does ("Name", "Name_1"). */
function uniqueNames(names: string[]): string[] {
  const out: string[] = []
  names.forEach((raw, i) => {
    let name = raw.trim() || `Column${i + 1}`
    for (let n = 1; out.includes(name); n++) name = `${raw.trim()}_${n}`
    out.push(name)
  })
  return out
}

// ── Steps ──

const DAY = 86400000
const EPOCH = Date.UTC(1899, 11, 30)

const escape = (c: string) => c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")

/**
 * Text as a number, as Power Query's type number reads it in the person's region (measured in Excel): the region's
 * thousands separator is ignored and its decimal separator read, "(5.00)" and "5.00-" are negative, "12%" is 0.12.
 * Cells also takes a currency sign ($ € £ ¥) and spaces between groups of three digits ("120 500"), which PDFs and
 * statements are full of. Null when it isn't a number.
 */
export function numberFrom(text: string, r: Regional = regional(), strict = false): number | null {
  let t = text.trim().replace(/[\u00a0\u202f]/g, " ")
  if (strict && /[$€£¥]|\d \d/.test(t)) return null
  if (!t) return null
  let sign = 1
  const paren = /^\((.*)\)$/.exec(t)
  if (paren) {
    sign = -1
    t = paren[1].trim()
  }
  if (/-$/.test(t)) {
    sign = -sign
    t = t.slice(0, -1).trim()
  }
  let scale = 1
  if (/%$/.test(t)) {
    scale = 0.01
    t = t.slice(0, -1).trim()
  }
  t = t.replace(/^([-+]?)\s*[$€£¥]\s*/, "$1").replace(/\s*[$€£¥]$/, "")
  if (/^[-+]?\d{1,3}( \d{3})+([.,]\d+)?$/.test(t)) t = t.replace(/ /g, "")
  if (r.group) t = t.replace(new RegExp(`(\\d)${escape(r.group)}(?=\\d{3})`, "g"), "$1")
  if (r.decimal !== ".") t = t.replace(r.decimal, ".")
  if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return null
  return sign * Number(t) * scale
}

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]

/**
 * Text as a date serial, as Power Query's type date reads it (measured in Excel): ISO dates always; numbers split by
 * / - . in the region's day, month and year order; "13 Feb 2026" and "Feb 13, 2026" with English month names. Null
 * when it isn't a date (13/02/2026 in a month-first region, say).
 */
export function dateFrom(text: string, r: Regional = regional()): number | null {
  const t = text.trim()
  const serial = (y: number, m: number, d: number) => {
    if (y < 100) y += y < 30 ? 2000 : 1900
    if (m < 1 || m > 12 || d < 1 || d > 31) return null
    const ms = Date.UTC(y, m - 1, d)
    if (new Date(ms).getUTCDate() !== d) return null
    return (ms - EPOCH) / DAY
  }
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ]\d{1,2}:\d{2}(?::\d{2})?)?$/.exec(t)
  if (m) return serial(+m[1], +m[2], +m[3])
  m = /^(\d{1,4})[/.-](\d{1,2})[/.-](\d{1,4})$/.exec(t)
  if (m) {
    const [a, b, c] = [+m[1], +m[2], +m[3]]
    if (m[1].length === 4) return serial(a, b, c)
    return r.dateOrder === "dmy" ? serial(c, b, a) : r.dateOrder === "mdy" ? serial(c, a, b) : serial(a, b, c)
  }
  const month = (name: string) => {
    const i = MONTHS.indexOf(name.slice(0, 3).toLowerCase())
    return i < 0 ? null : i + 1
  }
  m = /^(\d{1,2})[ -]([A-Za-z]{3,9})\.?[ -,]*(\d{2,4})$/.exec(t)
  if (m && month(m[2])) return serial(+m[3], month(m[2]) as number, +m[1])
  m = /^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{2,4})$/.exec(t)
  if (m && month(m[1])) return serial(+m[3], month(m[1]) as number, +m[2])
  return null
}

/** The locales a "Changed Type with Locale" step reads text in: their separators and date order. */
const CULTURES: Record<string, Regional> = {
  "en-US": EN_US,
  "en-GB": { ...EN_US, dateOrder: "dmy", shortDate: "dd/mm/yyyy" },
  "de-DE": { decimal: ",", group: ".", dateSeparator: ".", dateOrder: "dmy", shortDate: "dd.mm.yyyy" },
  "fr-FR": { decimal: ",", group: " ", dateSeparator: "/", dateOrder: "dmy", shortDate: "dd/mm/yyyy" },
  "nl-NL": { decimal: ",", group: ".", dateSeparator: "-", dateOrder: "dmy", shortDate: "d-m-yyyy" },
  "ja-JP": { ...EN_US, dateOrder: "ymd", shortDate: "yyyy/mm/dd" },
}

/** A locale's reading rules (an unknown one reads as the person's region). */
const cultureRules = (culture: string | undefined): Regional =>
  (culture && CULTURES[culture]) ||
  (culture && CULTURES[Object.keys(CULTURES).find((k) => k.startsWith(culture.slice(0, 2))) ?? ""]) ||
  regional()

function toType(v: QueryValue, t: QueryColumnType, r: Regional = regional()): QueryValue {
  if (v === null || t === "any") return v
  switch (t) {
    case "text":
      return typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)
    case "number":
    case "whole": {
      const n = typeof v === "number" ? v : typeof v === "boolean" ? Number(v) : numberFrom(v, r)
      if (n === null || !Number.isFinite(n)) return null
      return t === "whole" ? Math.round(n) : n
    }
    case "logical":
      if (typeof v === "boolean") return v
      if (/^true$/i.test(String(v))) return true
      if (/^false$/i.test(String(v))) return false
      return null
    case "date":
    case "datetime": {
      if (typeof v === "number") return t === "date" ? Math.floor(v) : v
      const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(v))
      if (!m) return t === "date" ? dateFrom(String(v), r) : null
      const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0))
      const serial = (ms - EPOCH) / DAY
      return t === "date" ? Math.floor(serial) : serial
    }
  }
}

function compare(a: QueryValue, b: QueryValue): number {
  if (a === b) return 0
  if (a === null) return -1
  if (b === null) return 1
  if (typeof a === "number" && typeof b === "number") return a - b
  if (typeof a === typeof b) return String(a) < String(b) ? -1 : 1
  const rank = (v: QueryValue) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2)
  return rank(a) - rank(b)
}

function passes(v: QueryValue, test: Extract<QueryStep, { op: "filter" }>): boolean {
  const want = test.value
  switch (test.test) {
    case "=":
      return v === want
    case "<>":
      return v !== want
    case ">":
      return v !== null && compare(v, want) > 0
    case ">=":
      return v !== null && compare(v, want) >= 0
    case "<":
      return v !== null && compare(v, want) < 0
    case "<=":
      return v !== null && compare(v, want) <= 0
    case "contains":
      return v !== null && String(v).includes(String(want))
    case "starts":
      return v !== null && String(v).startsWith(String(want))
    case "ends":
      return v !== null && String(v).endsWith(String(want))
  }
}

function column(r: QueryResult, name: string): number {
  const i = r.columns.indexOf(name)
  if (i < 0) throw new QueryError(`The column '${name}' of the table wasn't found.`)
  return i
}

/** One step on a table (a web page's list of tables only takes the "table" step). */
export function applyStep(input: QueryResult | SourceItem[], step: QueryStep): QueryResult | SourceItem[] {
  if (Array.isArray(input)) {
    if (step.op === "item") {
      const t = input.find((i) => i.id === step.id)
      if (!t) throw new QueryError(`The PDF has no ${step.id} any more.`)
      return t
    }
    if (step.op !== "table") throw new QueryError("Pick one of the source's tables first.")
    const t = input[step.index]
    if (!t) throw new QueryError(`The page has no table ${step.index + 1} any more.`)
    return t
  }
  const r = input
  switch (step.op) {
    case "table":
      throw new QueryError("This step only applies to a web page.")
    case "item":
      throw new QueryError("This step only applies to a PDF.")
    case "promote": {
      const [head, ...rest] = r.rows
      const names = uniqueNames((head ?? []).map((v, i) => (v === null ? `Column${i + 1}` : String(v))))
      return result(names, rest)
    }
    case "types": {
      const types = [...r.types]
      const at = step.columns.map(([name, t]) => {
        const i = column(r, name)
        types[i] = t
        return [i, t] as const
      })
      const rows = r.rows.map((row) => {
        const out = [...row]
        const rules = cultureRules(step.culture)
        for (const [i, t] of at) out[i] = toType(row[i], t, rules)
        return out
      })
      return { columns: r.columns, rows, types }
    }
    case "select":
    case "remove": {
      const keep =
        step.op === "select"
          ? step.columns.map((c) => column(r, c))
          : r.columns.map((_, i) => i).filter((i) => !step.columns.includes(r.columns[i]))
      if (step.op === "remove") for (const c of step.columns) column(r, c)
      return {
        columns: keep.map((i) => r.columns[i]),
        rows: r.rows.map((row) => keep.map((i) => row[i] ?? null)),
        types: keep.map((i) => r.types[i]),
      }
    }
    case "filter": {
      const i = column(r, step.column)
      return { ...r, rows: r.rows.filter((row) => passes(row[i] ?? null, step)) }
    }
    case "sort": {
      const keys = step.keys.map((k) => ({ i: column(r, k.column), desc: k.desc }))
      const rows = [...r.rows].sort((a, b) => {
        for (const k of keys) {
          const d = compare(a[k.i] ?? null, b[k.i] ?? null)
          if (d) return k.desc ? -d : d
        }
        return 0
      })
      return { ...r, rows }
    }
    case "replace": {
      const at = step.columns.map((name) => column(r, name))
      const rows = r.rows.map((row) => {
        const out = [...row]
        for (const i of at)
          if (typeof out[i] === "string" && step.find) out[i] = (out[i] as string).split(step.find).join(step.with)
        return out
      })
      return { ...r, rows }
    }
    case "distinct": {
      const seen = new Set<string>()
      const rows = r.rows.filter((row) => {
        const key = JSON.stringify(row)
        if (seen.has(key)) return false
        seen.add(key)
        return true
      })
      return { ...r, rows }
    }
    case "first":
      return { ...r, rows: r.rows.slice(0, Math.max(0, step.count)) }
  }
}

/** The source's data: a table, or a web page's tables. */
export async function sourceData(plan: QueryPlan, ctx: QueryContext): Promise<QueryResult | SourceItem[]> {
  const s = plan.source
  if (s.kind === "table") {
    const t = tableResult(ctx.wb, s.name)
    if (!t) throw new QueryError(`Couldn't find a table named '${s.name}' in this workbook.`)
    return t
  }
  if (s.kind === "pdf") return (await ctx.pdf(s.path)).map(pdfResult)
  const got = await ctx.fetch(s.url)
  if (s.kind === "csv") {
    const rows = parseCsv(decode(got.bytes, s.encoding), s.delimiter)
    const width = s.columns ?? Math.max(0, ...rows.map((r) => r.length))
    return result(
      columnNames(width),
      rows.map((r) => Array.from({ length: width }, (_, i) => (r[i] === undefined || r[i] === "" ? null : r[i]))),
    )
  }
  const text = decode(got.bytes, 65001)
  if (s.kind === "json") return jsonTable(text)
  return htmlTables(text)
}

/** A plan run through to its last step (or `upTo` steps: the editor's preview of an earlier step). */
export async function runPlan(plan: QueryPlan, ctx: QueryContext, upTo = plan.steps.length): Promise<QueryResult> {
  let data = await sourceData(plan, ctx)
  for (const { step } of plan.steps.slice(0, upTo)) data = applyStep(data, step)
  if (Array.isArray(data)) throw new QueryError("Pick one of the source's tables.")
  return data
}

/** A column's type as text in a locale reads: whole numbers, numbers, true/false, dates, else text. */
function columnType(values: Exclude<QueryValue, null>[], r: Regional): QueryColumnType {
  const all = (test: (v: Exclude<QueryValue, null>) => boolean) => values.length > 0 && values.every(test)
  const n = (v: QueryValue) => (typeof v === "number" ? v : typeof v === "string" ? numberFrom(v, r) : null)
  // Text with decimals ("5.00") is a decimal number even when its value is whole, as Power Query detects it.
  const decimals = new RegExp(`${escape(r.decimal)}\\d|%`)
  if (all((v) => Number.isInteger(n(v)) && !(typeof v === "string" && decimals.test(v)))) return "whole"
  if (all((v) => n(v) !== null)) return "number"
  if (all((v) => typeof v === "boolean" || /^(true|false)$/i.test(String(v)))) return "logical"
  if (all((v) => typeof v === "string" && dateFrom(v, r) !== null)) return "date"
  return "text"
}

const valuesOf = (r: QueryResult, i: number) =>
  r.rows.map((row) => row[i]).filter((v): v is Exclude<QueryValue, null> => v !== null && v !== "")

/** The types Power Query would detect for each column, in the person's region. */
export function detectTypes(r: QueryResult): [string, QueryColumnType][] {
  return r.columns.map((name, i) => [name, columnType(valuesOf(r, i), regional())])
}

const sameRules = (a: Regional, b: Regional) =>
  a.decimal === b.decimal && a.group === b.group && a.dateOrder === b.dateOrder

/**
 * The steps that type new data. Numbers written with a currency sign or spaced groups ("€ 450.00", "120 500") get
 * Replace Values steps first that take those out, so Excel's type number reads them too. Then "Changed Type" in the
 * person's region, and, for columns that only read as numbers or dates in another locale (a UK statement's 13/02/2026
 * in a US region), "Changed Type with Locale" in that one, as Excel's Change Type › Using Locale writes it.
 */
export function typeSteps(r: QueryResult, taken: readonly string[] = []): { name: string; step: QueryStep }[] {
  const here = regional()
  const steps: { name: string; step: QueryStep }[] = []
  const names = [...taken]
  const add = (step: QueryStep) => {
    const name = stepName(step, names)
    names.push(name)
    steps.push({ name, step })
  }
  // Signs and spaces Excel's type number doesn't take, by what's taken out (one step each, as the editor makes them).
  const cleanups = new Map<string, string[]>()
  r.columns.forEach((name, i) => {
    const values = valuesOf(r, i).filter((v): v is string => typeof v === "string")
    if (!values.length || !values.every((v) => numberFrom(v, here) !== null)) return
    if (values.every((v) => numberFrom(v, here, true) !== null)) return
    const found = new Set<string>()
    for (const v of values) {
      for (const sign of v.match(/\s*[$€£¥]\s*/g) ?? []) found.add(sign.trim())
      if (/\d[ \u00a0\u202f]\d/.test(v)) found.add(" ")
    }
    for (const f of found) cleanups.set(f, [...(cleanups.get(f) ?? []), name])
  })
  let data = r
  for (const find of [...cleanups.keys()].sort((a, b) => (a === " " ? 1 : b === " " ? -1 : 0))) {
    const step: QueryStep = { op: "replace", columns: cleanups.get(find) as string[], find, with: "" }
    add(step)
    data = applyStep(data, step) as QueryResult
  }
  const main: [string, QueryColumnType][] = []
  const other = new Map<string, [string, QueryColumnType][]>()
  data.columns.forEach((name, i) => {
    const values = valuesOf(data, i)
    const type = columnType(values, here)
    if (type === "text" && values.length) {
      const culture = Object.keys(CULTURES).find(
        (k) => !sameRules(CULTURES[k], here) && columnType(values, CULTURES[k]) !== "text",
      )
      if (culture) {
        other.set(culture, [...(other.get(culture) ?? []), [name, columnType(values, CULTURES[culture])]])
        return
      }
    }
    main.push([name, type])
  })
  if (main.length) add({ op: "types", columns: main })
  for (const [culture, columns] of other) add({ op: "types", columns, culture })
  return steps
}

// ── Loading and refreshing ──

/** A sheet name from a query's name (Excel names the new sheet after the query). */
export function sheetNameFor(wb: Workbook, base: string): string {
  const clean =
    base
      .replace(/[\\/?*[\]:]/g, " ")
      .trim()
      .slice(0, 31) || "Query"
  const taken = new Set(wb.sheets.map((s) => s.name.toLowerCase()))
  if (!taken.has(clean.toLowerCase())) return clean
  for (let n = 2; ; n++) {
    const name = `${clean.slice(0, 31 - String(n).length - 3)} (${n})`
    if (!taken.has(name.toLowerCase())) return name
  }
}

/** A table name from a query's name ("Sales North" → "Sales_North"), unique among tables and defined names. */
export function tableNameFor(wb: Workbook, base: string): string {
  let clean = base.replace(/[^A-Za-z0-9_.\\]/g, "_").replace(/^[^A-Za-z_\\]/, "_$&")
  if (/^[A-Za-z]{1,3}\d+$/.test(clean) || /^[rc]$/i.test(clean)) clean = `_${clean}`
  const taken = new Set([
    ...wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())),
    ...wb.names.all().map((n) => n.name.toLowerCase()),
  ])
  if (!taken.has(clean.toLowerCase())) return clean
  for (let n = 2; ; n++) if (!taken.has(`${clean}_${n}`.toLowerCase())) return `${clean}_${n}`
}

const DATE_FORMATS: Partial<Record<QueryColumnType, string>> = { date: "m/d/yyyy", datetime: "m/d/yyyy h:mm" }

/** The cells of a result at (row, col): its header row, then its rows. */
function resultCells(r: QueryResult, row: number, col: number): CellWrite[] {
  const cells: CellWrite[] = r.columns.map((name, i) => ({ row, col: col + i, data: { value: name } }))
  r.rows.forEach((line, y) =>
    r.columns.forEach((_, x) => {
      const v = line[x]
      if (v !== null && v !== undefined && v !== "") cells.push({ row: row + 1 + y, col: col + x, data: { value: v } })
    }),
  )
  return cells
}

/** Date columns get a date format (the cells hold serial numbers). */
function dateFormats(r: QueryResult, sheet: number, row: number, col: number): Op[] {
  const ops: Op[] = []
  r.types.forEach((t, x) => {
    const fmt = t ? DATE_FORMATS[t] : undefined
    if (fmt && r.rows.length)
      ops.push({
        kind: "format",
        sheet,
        ranges: [{ r0: row + 1, c0: col + x, r1: row + r.rows.length, c1: col + x }],
        patch: { numberFormat: fmt },
      })
  })
  return ops
}

/**
 * Close & Load: the query added (or changed) and its result put in a new table on a sheet of its own, after the
 * active one. Returns the ops and the new sheet's index.
 */
export function loadOps(
  wb: Workbook,
  query: WorkbookQuery,
  r: QueryResult,
  after: number,
): { ops: Op[]; sheet: number } {
  const sheet = after + 1
  const name = tableNameFor(wb, query.name)
  const range = { r0: 0, c0: 0, r1: Math.max(1, r.rows.length), c1: Math.max(0, r.columns.length - 1) }
  const table: Table = {
    id: nextTableId(wb.sheets),
    name,
    range,
    header: true,
    totals: false,
    columns: r.columns.map((c) => ({ name: c })),
    style: "TableStyleMedium7",
    firstColumn: false,
    lastColumn: false,
    rowStripes: true,
    columnStripes: false,
    autoFilter: { range: { ...range }, columns: {} },
    query: query.name,
  }
  const queries = [...wb.queries.filter((q) => q.name !== query.name), { ...query, table: name }]
  return {
    sheet,
    ops: [
      { kind: "addSheet", sheet, name: sheetNameFor(wb, query.name) },
      { kind: "setCells", sheet, cells: resultCells(r, 0, 0) },
      ...dateFormats(r, sheet, 0, 0),
      { kind: "tables", sheet, tables: [table] },
      { kind: "queries", queries },
    ],
  }
}

/** Refresh: the query's table filled again from its top-left cell, resized, and what the old data covered cleared. */
export function refreshTableOps(wb: Workbook, query: WorkbookQuery, r: QueryResult): Op[] {
  const found = query.table ? findTable(wb.sheets, query.table) : null
  if (!found) return []
  const sheet = wb.sheets[found.sheet]
  const old = found.table.range
  const range = {
    r0: old.r0,
    c0: old.c0,
    r1: old.r0 + Math.max(1, r.rows.length),
    c1: old.c0 + Math.max(0, r.columns.length - 1),
  }
  const cells = resultCells(r, old.r0, old.c0)
  const written = new Set(cells.map((c) => `${c.row},${c.col}`))
  for (let row = Math.min(old.r0, range.r0); row <= Math.max(old.r1, range.r1); row++)
    for (let col = Math.min(old.c0, range.c0); col <= Math.max(old.c1, range.c1); col++) {
      const inside = (row <= old.r1 && col <= old.c1) || (row <= range.r1 && col <= range.c1)
      if (inside && !written.has(`${row},${col}`) && sheet.cells.get(row, col)?.value !== undefined)
        cells.push({ row, col, data: undefined })
    }
  const tables = sheet.tables.map((t) =>
    t.name === found.table.name
      ? {
          ...copyTable(t),
          range,
          columns: r.columns.map((name) => t.columns.find((c) => c.name === name) ?? { name }),
          autoFilter: t.autoFilter ? { ...t.autoFilter, range: { ...range } } : null,
        }
      : copyTable(t),
  )
  return [
    { kind: "setCells", sheet: found.sheet, cells },
    ...dateFormats(r, found.sheet, old.r0, old.c0),
    { kind: "tables", sheet: found.sheet, tables },
  ]
}

/** A query's plan, when Cells can run it. */
export const planOf = (q: WorkbookQuery): QueryPlan | null => readQueryPlan(q.formula)

/** The query with a new plan (its M written anew). */
export const withPlan = (q: WorkbookQuery, plan: QueryPlan): WorkbookQuery => ({ ...q, formula: queryFormula(plan) })
