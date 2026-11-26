/**
 * Data › Get Data: the workbook's queries (Power Query), as Excel keeps them: each a name, a description and its M
 * formula, loaded into a table on a sheet or kept as a connection only. Cells keeps every query as written; the ones
 * in the part of M it knows (a table of this workbook or a file on the web, then the steps the Query Editor's
 * buttons make) it can show step by step, change and refresh. Its own queries are written as Excel's editor writes
 * them, so Excel can refresh them too.
 */

export type QueryColumnType = "text" | "number" | "whole" | "date" | "datetime" | "logical" | "any"

export type QuerySource =
  /** A table (or named range) of this workbook: Excel.CurrentWorkbook(){[Name="…"]}[Content]. */
  | { kind: "table"; name: string }
  /** A CSV file on the web: Csv.Document(Web.Contents(url), [Delimiter, Columns, Encoding, QuoteStyle]). */
  | { kind: "csv"; url: string; delimiter: string; encoding: number; columns?: number }
  /** A web page's tables: Web.Page(Web.Contents(url)); a "table" step then picks one. */
  | { kind: "html"; url: string }
  /** JSON on the web, a list of records: Table.FromRecords(Json.Document(Web.Contents(url))). */
  | { kind: "json"; url: string }
  /** A PDF's tables and pages: Pdf.Tables(File.Contents(path), [Implementation="1.3"]); an "item" step picks one. */
  | { kind: "pdf"; path: string }

export type FilterTest = "=" | "<>" | ">" | ">=" | "<" | "<=" | "contains" | "starts" | "ends"
export type QueryValue = string | number | boolean | null

export type QueryStep =
  /** A web page's nth table: Source{n}[Data]. */
  | { op: "table"; index: number }
  /** A PDF's table or page by its Id ("Table001", "Page001"): Source{[Id="Table001"]}[Data]. */
  | { op: "item"; id: string }
  | { op: "promote" }
  /** Replace Values in text: Table.ReplaceValue(prev, find, with, Replacer.ReplaceText, {columns}). */
  | { op: "replace"; columns: string[]; find: string; with: string }
  /** Column types; `culture` ("en-GB") reads text as that locale does (Excel's Change Type › Using Locale). */
  | { op: "types"; columns: [string, QueryColumnType][]; culture?: string }
  | { op: "select"; columns: string[] }
  | { op: "remove"; columns: string[] }
  | { op: "filter"; column: string; test: FilterTest; value: QueryValue }
  | { op: "sort"; keys: { column: string; desc: boolean }[] }
  | { op: "distinct" }
  | { op: "first"; count: number }

/** A query Cells understands: where its data comes from, and the steps after that, each with its name. */
export type QueryPlan = { source: QuerySource; steps: { name: string; step: QueryStep }[] }

export type WorkbookQuery = {
  name: string
  description: string
  /** The query's M formula (let … in …), as written in the file or as Cells writes its plan. */
  formula: string
  /** The table on a sheet it fills (Excel's "Load To… Table"), or none: a connection only. */
  table: string | null
}

export const copyQueries = (list: readonly WorkbookQuery[]): WorkbookQuery[] => list.map((q) => ({ ...q }))

// ── M: a small reader for the expressions the Query Editor writes ───────────────────────────────────────────────

type MNode =
  | { t: "str"; v: string }
  | { t: "num"; v: number }
  | { t: "lit"; v: boolean | null }
  | { t: "id"; v: string }
  | { t: "list"; items: MNode[] }
  | { t: "rec"; fields: [string, MNode][] }
  | { t: "call"; fn: string; args: MNode[] }
  | { t: "index"; of: MNode; item: MNode }
  | { t: "field"; of: MNode | null; name: string }
  | { t: "each"; body: MNode }
  | { t: "bin"; op: string; a: MNode; b: MNode }
  | { t: "type"; v: string }

class MReader {
  private at = 0
  constructor(private readonly s: string) {}

  private space() {
    for (;;) {
      const m = /^(?:\s+|\/\/[^\n]*|\/\*[\s\S]*?\*\/)/.exec(this.s.slice(this.at))
      if (!m) return
      this.at += m[0].length
    }
  }
  peek(re: RegExp): RegExpExecArray | null {
    this.space()
    return re.exec(this.s.slice(this.at))
  }
  take(re: RegExp): RegExpExecArray | null {
    const m = this.peek(re)
    if (m) this.at += m[0].length
    return m
  }
  expect(re: RegExp) {
    if (!this.take(re)) throw new Error(`M: expected ${re} at ${this.at}`)
  }
  done() {
    this.space()
    return this.at >= this.s.length
  }

  /** A name: identifier (dotted) or #"quoted". */
  name(): string | null {
    const q = this.take(/^#"((?:[^"]|"")*)"/)
    if (q) return q[1].replace(/""/g, '"')
    const id = this.take(/^[A-Za-z_][\w.]*/)
    return id ? id[0] : null
  }

  expr(): MNode {
    let a = this.unary()
    for (;;) {
      const op = this.take(/^(<>|>=|<=|=|>|<|and\b|or\b|&)/)
      if (!op) return a
      a = { t: "bin", op: op[1], a, b: this.unary() }
    }
  }

  private unary(): MNode {
    if (this.take(/^each\b/)) return { t: "each", body: this.expr() }
    if (this.take(/^type\b/)) {
      const n = this.name()
      if (!n) throw new Error("M: type")
      return { t: "type", v: n }
    }
    let node = this.primary()
    for (;;) {
      if (this.take(/^\{/)) {
        const item = this.expr()
        this.expect(/^\}/)
        node = { t: "index", of: node, item }
      } else if (this.peek(/^\[/) && (node.t === "index" || node.t === "id" || node.t === "call")) {
        this.take(/^\[/)
        const n = this.name()
        this.expect(/^\]/)
        node = { t: "field", of: node, name: n ?? "" }
      } else return node
    }
  }

  private primary(): MNode {
    const str = this.take(/^"((?:[^"]|"")*)"/)
    if (str) return { t: "str", v: str[1].replace(/""/g, '"') }
    const num = this.take(/^-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?/)
    if (num) return { t: "num", v: Number(num[0]) }
    if (this.take(/^true\b/)) return { t: "lit", v: true }
    if (this.take(/^false\b/)) return { t: "lit", v: false }
    if (this.take(/^null\b/)) return { t: "lit", v: null }
    if (this.take(/^\(/)) {
      const e = this.expr()
      this.expect(/^\)/)
      return e
    }
    if (this.take(/^\{/)) {
      const items: MNode[] = []
      if (!this.take(/^\}/)) {
        do items.push(this.expr())
        while (this.take(/^,/))
        this.expect(/^\}/)
      }
      return { t: "list", items }
    }
    if (this.take(/^\[/)) {
      // A record [a=1, b=2], or a field of the row in an each ([Region]).
      const fields: [string, MNode][] = []
      const first = this.name()
      if (first !== null && this.take(/^\]/)) return { t: "field", of: null, name: first }
      if (first !== null) {
        this.expect(/^=/)
        fields.push([first, this.expr()])
        while (this.take(/^,/)) {
          const n = this.name()
          this.expect(/^=/)
          fields.push([n ?? "", this.expr()])
        }
      }
      this.expect(/^\]/)
      return { t: "rec", fields }
    }
    const n = this.name()
    if (n === null) throw new Error(`M: unexpected text at ${this.at}`)
    if (this.take(/^\(/)) {
      const args: MNode[] = []
      if (!this.take(/^\)/)) {
        do args.push(this.expr())
        while (this.take(/^,/))
        this.expect(/^\)/)
      }
      return { t: "call", fn: n, args }
    }
    return { t: "id", v: n }
  }
}

const TYPE_NAMES: Record<string, QueryColumnType> = {
  text: "text",
  number: "number",
  "Int64.Type": "whole",
  date: "date",
  datetime: "datetime",
  logical: "logical",
  any: "any",
}

const TYPE_M: Record<QueryColumnType, string> = {
  text: "type text",
  number: "type number",
  whole: "Int64.Type",
  date: "type date",
  datetime: "type datetime",
  logical: "type logical",
  any: "type any",
}

const str = (n: MNode | undefined) => (n?.t === "str" ? n.v : null)
const strs = (n: MNode | undefined) =>
  n?.t === "list" && n.items.every((i) => i.t === "str") ? n.items.map(str) : null
const isRef = (n: MNode | undefined, prev: string) => n?.t === "id" && n.v === prev
const webUrl = (n: MNode | undefined) =>
  n?.t === "call" && n.fn === "Web.Contents" && n.args.length === 1 ? str(n.args[0]) : null
const value = (n: MNode): QueryValue | undefined => (n.t === "str" || n.t === "num" || n.t === "lit" ? n.v : undefined)

function readSource(e: MNode): QuerySource | null {
  // Excel.CurrentWorkbook(){[Name="Sales"]}[Content]
  if (e.t === "field" && e.name === "Content" && e.of?.t === "index") {
    const call = e.of.of
    const key = e.of.item
    if (call.t === "call" && call.fn === "Excel.CurrentWorkbook" && !call.args.length && key.t === "rec") {
      const name = key.fields.length === 1 && key.fields[0][0] === "Name" ? str(key.fields[0][1]) : null
      if (name) return { kind: "table", name }
    }
  }
  if (e.t === "call" && e.fn === "Csv.Document") {
    const url = webUrl(e.args[0])
    const opts = e.args[1]
    if (!url || (opts && opts.t !== "rec")) return null
    const field = (k: string) => (opts?.t === "rec" ? opts.fields.find((f) => f[0] === k)?.[1] : undefined)
    const delimiter = str(field("Delimiter")) ?? ","
    const enc = field("Encoding")
    const cols = field("Columns")
    return {
      kind: "csv",
      url,
      delimiter,
      encoding: enc?.t === "num" ? enc.v : 65001,
      ...(cols?.t === "num" ? { columns: cols.v } : {}),
    }
  }
  if (e.t === "call" && e.fn === "Web.Page" && e.args.length === 1) {
    const url = webUrl(e.args[0])
    return url ? { kind: "html", url } : null
  }
  // Pdf.Tables(File.Contents("…"), [Implementation="1.3"]): other options are Excel's alone.
  if (e.t === "call" && e.fn === "Pdf.Tables" && (e.args.length === 1 || e.args.length === 2)) {
    const file = e.args[0]
    const path = file.t === "call" && file.fn === "File.Contents" && file.args.length === 1 ? str(file.args[0]) : null
    const opts = e.args[1]
    const plain = !opts || (opts.t === "rec" && opts.fields.every(([k]) => k === "Implementation"))
    return path && plain ? { kind: "pdf", path } : null
  }
  if (e.t === "call" && e.fn === "Table.FromRecords" && e.args.length === 1) {
    const inner = e.args[0]
    const url = inner.t === "call" && inner.fn === "Json.Document" ? webUrl(inner.args[0]) : null
    return url ? { kind: "json", url } : null
  }
  return null
}

const TESTS: Record<string, FilterTest> = { "=": "=", "<>": "<>", ">": ">", ">=": ">=", "<": "<", "<=": "<=" }
const TEXT_TESTS: Record<string, FilterTest> = {
  "Text.Contains": "contains",
  "Text.StartsWith": "starts",
  "Text.EndsWith": "ends",
}

function readFilter(body: MNode): QueryStep | null {
  if (body.t === "bin" && TESTS[body.op] && body.a.t === "field" && body.a.of === null) {
    const v = value(body.b)
    return v === undefined ? null : { op: "filter", column: body.a.name, test: TESTS[body.op], value: v }
  }
  if (body.t === "call" && TEXT_TESTS[body.fn] && body.args.length === 2) {
    const [f, v] = body.args
    if (f.t === "field" && f.of === null && v.t === "str")
      return { op: "filter", column: f.name, test: TEXT_TESTS[body.fn], value: v.v }
  }
  return null
}

function readStep(e: MNode, prev: string): QueryStep | null {
  if (e.t === "field" && e.name === "Data" && e.of?.t === "index" && isRef(e.of.of, prev) && e.of.item.t === "num")
    return { op: "table", index: e.of.item.v }
  if (e.t === "field" && e.name === "Data" && e.of?.t === "index" && isRef(e.of.of, prev) && e.of.item.t === "rec") {
    const f = e.of.item.fields
    const id = f.length === 1 && f[0][0] === "Id" ? str(f[0][1]) : null
    return id ? { op: "item", id } : null
  }
  if (e.t !== "call" || !isRef(e.args[0], prev)) return null
  const [, a, b] = e.args
  switch (e.fn) {
    case "Table.PromoteHeaders":
      return { op: "promote" }
    case "Table.TransformColumnTypes": {
      if (a?.t !== "list") return null
      const columns: [string, QueryColumnType][] = []
      for (const pair of a.items) {
        if (pair.t !== "list" || pair.items.length !== 2) return null
        const name = str(pair.items[0])
        const t = pair.items[1]
        const type = t.t === "type" ? TYPE_NAMES[t.v] : t.t === "id" ? TYPE_NAMES[t.v] : undefined
        if (name === null || !type) return null
        columns.push([name, type])
      }
      if (b === undefined) return { op: "types", columns }
      const culture = str(b)
      return culture && e.args.length === 3 ? { op: "types", columns, culture } : null
    }
    case "Table.SelectColumns":
    case "Table.RemoveColumns": {
      const cols = strs(a)
      if (!cols || b !== undefined) return null
      return { op: e.fn === "Table.SelectColumns" ? "select" : "remove", columns: cols as string[] }
    }
    case "Table.SelectRows":
      return a?.t === "each" ? readFilter(a.body) : null
    case "Table.Sort": {
      if (a?.t !== "list") return null
      const keys: { column: string; desc: boolean }[] = []
      for (const k of a.items) {
        if (k.t !== "list" || k.items.length !== 2) return null
        const col = str(k.items[0])
        const dir = k.items[1]
        if (col === null || dir.t !== "id" || !/^Order\.(Ascending|Descending)$/.test(dir.v)) return null
        keys.push({ column: col, desc: dir.v === "Order.Descending" })
      }
      return { op: "sort", keys }
    }
    case "Table.ReplaceValue": {
      const [, f, w, how, cols] = e.args
      const find = str(f)
      const by = str(w)
      const columns = strs(cols)
      if (find === null || by === null || !columns || how?.t !== "id" || how.v !== "Replacer.ReplaceText") return null
      return e.args.length === 5 ? { op: "replace", columns: columns as string[], find, with: by } : null
    }
    case "Table.Distinct":
      return a === undefined ? { op: "distinct" } : null
    case "Table.FirstN":
      return a?.t === "num" && b === undefined ? { op: "first", count: a.v } : null
  }
  return null
}

/** The plan of a query's formula, or null when it uses M that Cells doesn't run. */
export function readQueryPlan(formula: string): QueryPlan | null {
  try {
    const r = new MReader(formula)
    r.expect(/^let\b/)
    const steps: { name: string; expr: MNode }[] = []
    do {
      const name = r.name()
      if (!name) return null
      r.expect(/^=/)
      steps.push({ name, expr: r.expr() })
    } while (r.take(/^,/))
    r.expect(/^in\b/)
    const result = r.name()
    if (!r.done() || !steps.length || result !== steps[steps.length - 1].name) return null
    const source = readSource(steps[0].expr)
    if (!source) return null
    const plan: QueryPlan = { source, steps: [] }
    for (let i = 1; i < steps.length; i++) {
      const step = readStep(steps[i].expr, steps[i - 1].name)
      if (!step) return null
      plan.steps.push({ name: steps[i].name, step })
    }
    return plan
  } catch {
    return null
  }
}

// ── M: written as the Query Editor writes it ────────────────────────────────────────────────────────────────

const mStr = (v: string) => `"${v.replace(/"/g, '""')}"`
const mName = (v: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(v) ? v : `#${mStr(v)}`)
const mValue = (v: QueryValue) => (typeof v === "string" ? mStr(v) : v === null ? "null" : String(v))
const mList = (items: string[]) => `{${items.join(", ")}}`

function sourceM(s: QuerySource): string {
  switch (s.kind) {
    case "table":
      return `Excel.CurrentWorkbook(){[Name=${mStr(s.name)}]}[Content]`
    case "csv":
      return (
        `Csv.Document(Web.Contents(${mStr(s.url)}),[Delimiter=${mStr(s.delimiter)}, ` +
        `${s.columns !== undefined ? `Columns=${s.columns}, ` : ""}Encoding=${s.encoding}, QuoteStyle=QuoteStyle.Csv])`
      )
    case "html":
      return `Web.Page(Web.Contents(${mStr(s.url)}))`
    case "json":
      return `Table.FromRecords(Json.Document(Web.Contents(${mStr(s.url)})))`
    case "pdf":
      return `Pdf.Tables(File.Contents(${mStr(s.path)}), [Implementation="1.3"])`
  }
}

function stepM(step: QueryStep, prev: string): string {
  const p = mName(prev)
  switch (step.op) {
    case "table":
      return `${p}{${step.index}}[Data]`
    case "item":
      return `${p}{[Id=${mStr(step.id)}]}[Data]`
    case "promote":
      return `Table.PromoteHeaders(${p}, [PromoteAllScalars=true])`
    case "types":
      return `Table.TransformColumnTypes(${p},{${step.columns.map(([c, t]) => `{${mStr(c)}, ${TYPE_M[t]}}`).join(", ")}}${step.culture ? `, ${mStr(step.culture)}` : ""})`
    case "select":
      return `Table.SelectColumns(${p},${mList(step.columns.map(mStr))})`
    case "remove":
      return `Table.RemoveColumns(${p},${mList(step.columns.map(mStr))})`
    case "filter": {
      const field = `[${step.column.replace(/]/g, "")}]`
      const fn = { contains: "Text.Contains", starts: "Text.StartsWith", ends: "Text.EndsWith" }[step.test as string]
      const test = fn ? `${fn}(${field}, ${mValue(step.value)})` : `${field} ${step.test} ${mValue(step.value)}`
      return `Table.SelectRows(${p}, each (${test}))`
    }
    case "sort":
      return `Table.Sort(${p},{${step.keys.map((k) => `{${mStr(k.column)}, Order.${k.desc ? "Descending" : "Ascending"}}`).join(", ")}})`
    case "replace":
      return `Table.ReplaceValue(${p},${mStr(step.find)},${mStr(step.with)},Replacer.ReplaceText,${mList(step.columns.map(mStr))})`
    case "distinct":
      return `Table.Distinct(${p})`
    case "first":
      return `Table.FirstN(${p},${step.count})`
  }
}

/** The step names the Query Editor gives. */
export const STEP_NAMES: Record<QueryStep["op"], string> = {
  table: "Data",
  item: "Item",
  promote: "Promoted Headers",
  replace: "Replaced Value",
  types: "Changed Type",
  select: "Removed Other Columns",
  remove: "Removed Columns",
  filter: "Filtered Rows",
  sort: "Sorted Rows",
  distinct: "Removed Duplicates",
  first: "Kept First Rows",
}

/** A step's name, unique among the names before it ("Filtered Rows1" for a second filter, as the editor does). */
export function stepName(step: QueryStep, taken: readonly string[]): string {
  const base =
    step.op === "table"
      ? `Data${step.index}`
      : step.op === "item"
        ? step.id
        : step.op === "types" && step.culture
          ? "Changed Type with Locale"
          : STEP_NAMES[step.op]
  if (!taken.includes(base)) return base
  for (let n = 1; ; n++) if (!taken.includes(`${base}${n}`)) return `${base}${n}`
}

/** A plan's M formula, formatted as Excel's editor formats it. */
export function queryFormula(plan: QueryPlan): string {
  const lines = [`    Source = ${sourceM(plan.source)}`]
  let prev = "Source"
  for (const { name, step } of plan.steps) {
    lines.push(`    ${mName(name)} = ${stepM(step, prev)}`)
    prev = name
  }
  return `let\r\n${lines.join(",\r\n")}\r\nin\r\n    ${mName(prev)}`
}

/** A query name Excel accepts, unique among the workbook's queries ("Sales" → "Sales (2)"). */
export function queryName(base: string, taken: readonly string[]): string {
  const clean = base.replace(/[."[\]]/g, " ").trim() || "Query1"
  const used = new Set(taken.map((t) => t.toLowerCase()))
  if (!used.has(clean.toLowerCase())) return clean
  for (let n = 2; ; n++) if (!used.has(`${clean} (${n})`.toLowerCase())) return `${clean} (${n})`
}
