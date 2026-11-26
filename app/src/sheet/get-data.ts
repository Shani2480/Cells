import {
  currentRegion,
  queryName,
  readQueryPlan,
  stepName,
  tableAt,
  type Op,
  type QueryPlan,
  type QuerySource,
  type QueryStep,
  type QueryValue,
  type WorkbookQuery,
} from "@cells/core"
import type { PdfItem } from "@cells/io"
import type { WebPort } from "@cells/platform"
import { pdfItems, type PdfRead, type Recognizer } from "@/pdf/pdf"
import type { SheetController } from "./controller"
import { FILTER_BUTTON_ROOM } from "./pivots"
import {
  applyStep,
  dateFrom,
  typeSteps,
  numberFrom,
  pdfResult,
  type SourceItem,
  htmlTables,
  jsonTable,
  loadOps,
  parseCsv,
  QueryError,
  refreshTableOps,
  runPlan,
  withPlan,
  type QueryContext,
  type QueryResult,
} from "./power-query"

/**
 * Data › Get & Transform Data: From Table/Range and From Web make a query, the Query Editor shapes it, and Close &
 * Load puts it in a table on a new sheet (or keeps it as a connection only). Recent Sources reopens a source used
 * before; Existing Connections loads a query or goes to its data. Refresh runs a query again into its table.
 */

/** What the Query Editor opens on: a query (new or one of the workbook's) and its plan. */
export type EditorRequest = { query: WorkbookQuery; plan: QueryPlan; editing: string | null }

/** A source used before, newest first (kept on this device). */
export type RecentSource =
  { kind: "web"; url: string } | { kind: "table"; name: string } | { kind: "pdf"; path: string; name: string }

/** One of the Navigator's items: a web page's table, or a PDF's table or page, with its data. */
export type NavigatorItem = {
  name: string
  kind: "Table" | "Page"
  data: QueryResult
  /** What it loads as (headings promoted, types set), for the preview. */
  preview?: QueryResult
  /** The preview's cells recognition wasn't sure of, as [row, column] (a heading row isn't counted). */
  unsure?: [number, number][]
  /** The step that picks it from its source, and the name it gives the query. */
  step: QueryStep
  base: string
}

/** The Navigator: a source with several tables, while it's open. */
export type NavigatorState = {
  title: string
  source: QuerySource
  items: NavigatorItem[]
  /** PDF pages that are scans and weren't read (text recognition is off). */
  unread: number[]
}

const RECENT_KEY = "cells.recentSources"
const RECENT_MAX = 10

let webPort: WebPort | null = null
/** The app sets how files on the web are read (the desktop app fetches them itself). */
export function setWebPort(port: WebPort) {
  webPort = port
}

let fileReader: ((path: string) => Promise<Uint8Array>) | null = null
/** The app sets how a file is read again from its path (Refresh of a PDF query); a browser can't. */
export function setFileReader(read: (path: string) => Promise<Uint8Array>) {
  fileReader = read
}

/** A file's name from its path or content link ("/home/alex/Report.pdf" → "Report.pdf"). */
export const fileName = (path: string) =>
  decodeURIComponent(path.split(/[/\\]/).pop() ?? path)
    .split(/[:/]/)
    .pop() ?? path

/**
 * Whether a table's first row is its headings (Power Query promotes them): every cell filled with text that isn't a
 * number or date, over a column whose values below are numbers or dates.
 */
export function hasHeadings(r: QueryResult): boolean {
  const [head, ...rest] = r.rows
  if (!head || !rest.length) return false
  const plain = (v: QueryValue) => typeof v === "string" && numberFrom(v) === null && dateFrom(v) === null
  if (!head.every((v) => v !== null && plain(v))) return false
  return head.some((_, c) => {
    const below = rest.map((row) => row[c]).filter((v) => v !== null && v !== "")
    return below.length > 0 && below.filter((v) => !plain(v)).length >= 0.8 * below.length
  })
}

export function recentSources(): RecentSource[] {
  try {
    const list = JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]") as RecentSource[]
    return Array.isArray(list) ? list.slice(0, RECENT_MAX) : []
  } catch {
    return []
  }
}

function remember(source: RecentSource) {
  const key = (s: RecentSource) => (s.kind === "web" ? s.url : s.kind === "pdf" ? s.path : s.name)
  const same = (s: RecentSource) => s.kind === source.kind && key(s) === key(source)
  try {
    localStorage.setItem(
      RECENT_KEY,
      JSON.stringify([source, ...recentSources().filter((s) => !same(s))].slice(0, RECENT_MAX)),
    )
  } catch {
    // Storage may be full or off; the list is only a convenience.
  }
}

export function forgetRecentSources() {
  try {
    localStorage.removeItem(RECENT_KEY)
  } catch {
    // Nothing to forget.
  }
}

/** A query name from a web address: its last path part without the extension, or the site. */
export function nameFromUrl(url: string): string {
  try {
    const u = new URL(url)
    const last = decodeURIComponent(u.pathname.split("/").filter(Boolean).pop() ?? "")
    return last.replace(/\.[^.]+$/, "") || u.hostname
  } catch {
    return "Query"
  }
}

/** What a web file is: CSV, JSON or a page, from its type, then its address, then its first characters. */
export function webKind(url: string, type: string, bytes: Uint8Array): "csv" | "json" | "html" {
  const t = type.toLowerCase()
  if (/json/.test(t)) return "json"
  if (/csv|text\/plain|tab-separated/.test(t)) return "csv"
  if (/html|xml/.test(t)) return "html"
  const path = url.split(/[?#]/)[0].toLowerCase()
  if (/\.json$/.test(path)) return "json"
  if (/\.(csv|txt|tsv)$/.test(path)) return "csv"
  const head = new TextDecoder().decode(bytes.subarray(0, 512)).trimStart()
  if (/^[[{]/.test(head)) return "json"
  if (/^</.test(head)) return "html"
  return "csv"
}

let recognizer: Recognizer | null = null
/** Text recognition for scanned PDF pages, while it's turned on (null: off). */
export function setRecognizer(r: Recognizer | null) {
  recognizer = r
}

let turnOnRecognition: (() => void) | null = null
/** How text recognition is turned on from a prompt (the app saves the setting, which sets the recognizer). */
export function setRecognitionSwitch(on: (() => void) | null) {
  turnOnRecognition = on
}

/** A file read again from its path; null where it can't be (moved, or a browser that only knows its name). */
async function readFile(path: string): Promise<Uint8Array | null> {
  try {
    return fileReader ? await fileReader(path) : null
  } catch {
    return null
  }
}

const sameBytes = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((v, i) => v === b[i])

export class GetData {
  /** The Query Editor's query, while it's open. */
  editor: EditorRequest | null = null
  /** The Navigator (From Web's page tables, From PDF's tables and pages), while it's open. */
  navigator: NavigatorState | null = null
  /** What's under way that takes a while ("Recognizing text: page 2 of 5…"), for the status bar. */
  busy: string | null = null
  /** PDFs read this session, by path: the editor's previews don't read the file again; Refresh does. */
  private pdfs = new Map<string, { bytes: Uint8Array; items: PdfItem[] }>()

  constructor(private readonly c: SheetController) {}

  context(fresh = false): QueryContext {
    return {
      wb: this.c.wb,
      fetch: (url) => {
        if (!webPort) throw new QueryError("Files on the web can't be read here.")
        return webPort.fetch(url)
      },
      pdf: (path) => this.pdfAt(path, fresh),
    }
  }

  /** A PDF's items: read again from its path for a refresh, else the copy read this session. */
  private async pdfAt(path: string, fresh: boolean): Promise<PdfItem[]> {
    const known = this.pdfs.get(path)
    const bytes = fresh || !known ? await readFile(path) : null
    if (fresh || !known) {
      if (!bytes && !known)
        throw new QueryError(`Cells couldn't read ${fileName(path)}. Choose Get Data › From PDF to pick it again.`)
    }
    if (known && (!bytes || sameBytes(bytes, known.bytes))) return known.items
    const read = await this.readPdf(bytes as Uint8Array)
    this.pdfs.set(path, { bytes: bytes as Uint8Array, items: read.items })
    return read.items
  }

  /** The tables and pages of a PDF (scanned pages read when text recognition is on). */
  private async readPdf(bytes: Uint8Array): Promise<PdfRead> {
    const say = (text: string | null) => {
      this.busy = text
      this.c.changed()
    }
    try {
      return await pdfItems(bytes, recognizer, (page, of) =>
        say(of > 1 ? `Recognizing text: page ${page} of ${of}…` : "Recognizing text…"),
      )
    } catch (e) {
      throw new QueryError(message(e))
    } finally {
      if (this.busy) say(null)
    }
  }

  /** Whether scanned pages can be read once the person agrees (text recognition is off, and can be turned on). */
  get canTurnOnRecognition(): boolean {
    return !recognizer && turnOnRecognition !== null
  }

  /** Turns text recognition on and reads the Navigator's PDF again, its scanned pages too. */
  async readScans(): Promise<void> {
    const nav = this.navigator
    if (!nav || nav.source.kind !== "pdf" || !turnOnRecognition) return
    const known = this.pdfs.get(nav.source.path)
    if (!known) return
    turnOnRecognition()
    this.navigator = null
    await this.fromPdf(nav.source.path, known.bytes)
  }

  private names(): string[] {
    return this.c.wb.queries.map((q) => q.name)
  }

  /** Opens the editor on a new query of a plan. */
  private edit(base: string, plan: QueryPlan) {
    this.editor = {
      query: { name: queryName(base, this.names()), description: "", formula: "", table: null },
      plan,
      editing: null,
    }
    this.c.shell.openDialog?.("queryEditor")
  }

  /** From Table/Range: the selection's table (a table is made of the data around it first, as Excel asks to). */
  fromTable(): boolean {
    const { row, col } = this.c.selection.active
    let table = tableAt(this.c.sheet, row, col)
    if (!table) {
      const range = currentRegion(this.c.sheet, { row, col })
      if (range.r0 === range.r1 && range.c0 === range.c1 && this.c.sheet.cells.value(row, col) === undefined) {
        this.c.shell.alert("Select a cell in your data first, then choose From Table/Range.")
        return false
      }
      const ops = this.c.tables.createOps(range, true, "TableStyleMedium2")
      if (typeof ops === "string") {
        this.c.shell.alert(ops)
        return false
      }
      this.c.transact(ops, { label: "Create Table", before: this.c.selection })
      table = tableAt(this.c.sheet, row, col)
      if (!table) return false
    }
    // A query's own table: its query opens (as Excel does) when Cells can run it.
    if (table.query && this.c.wb.queries.some((q) => q.name === table.query && readQueryPlan(q.formula)))
      return this.editQuery(table.query)
    this.openTable(table.name)
    return true
  }

  /** The editor on a workbook table, with the types Power Query detects. */
  openTable(name: string) {
    remember({ kind: "table", name })
    const plan: QueryPlan = { source: { kind: "table", name }, steps: [] }
    void runPlan(plan, this.context()).then(
      (r) => {
        plan.steps.push(...typeSteps(r))
        this.edit(name, plan)
      },
      (e: unknown) => this.c.shell.alert(message(e)),
    )
  }

  /**
   * From Web: reads the address, then opens the editor (CSV and JSON) or the Navigator (a page's tables). Rejects
   * with the reason when it can't be read.
   */
  async fromWeb(url: string): Promise<void> {
    const ctx = this.context()
    const got = await ctx.fetch(url)
    remember({ kind: "web", url })
    const kind = webKind(url, got.type, got.bytes)
    const text = new TextDecoder().decode(got.bytes)
    if (kind === "html") {
      const tables = htmlTables(text)
      if (!tables.length) throw new QueryError("The page has no tables to load.")
      this.navigator = {
        title: url,
        source: { kind: "html", url },
        items: tables.map((data, index) => ({
          name: `Table ${index + 1}`,
          kind: "Table",
          data,
          step: { op: "table", index },
          base: `Table ${index + 1}`,
        })),
        unread: [],
      }
      this.c.shell.openDialog?.("queryNavigator")
      return
    }
    let source: QuerySource
    let data: QueryResult
    const steps: QueryPlan["steps"] = []
    if (kind === "json") {
      source = { kind: "json", url }
      data = jsonTable(text)
    } else {
      const delimiter =
        /\.tsv$/i.test(url.split(/[?#]/)[0]) ||
        (text.split("\n")[0].includes("\t") && !text.split("\n")[0].includes(","))
          ? "\t"
          : ","
      const rows = parseCsv(text, delimiter)
      const columns = Math.max(0, ...rows.map((r) => r.length))
      source = { kind: "csv", url, delimiter, encoding: 65001, columns }
      steps.push({ name: "Promoted Headers", step: { op: "promote" } })
      data = await runPlan({ source, steps }, ctx)
    }
    steps.push(
      ...typeSteps(
        data,
        steps.map((s) => s.name),
      ),
    )
    this.edit(nameFromUrl(url), { source, steps })
  }

  /**
   * From PDF: the file's tables and pages in the Navigator (a PDF with none says so). `path` is where Refresh reads it
   * again: a full path or content link, or only its name where the platform gives none (a browser).
   */
  async fromPdf(path: string, bytes: Uint8Array): Promise<void> {
    const read = await this.readPdf(bytes)
    this.pdfs.set(path, { bytes, items: read.items })
    remember({ kind: "pdf", path, name: fileName(path) })
    const items = read.items.filter((i) => i.rows.length)
    if (!items.length && read.unread.length && turnOnRecognition) {
      const on = turnOnRecognition
      this.c.shell.confirm(
        "This PDF is scanned: its pages are pictures of text. Read them with text recognition on this " +
          "computer? Nothing is sent anywhere. You can turn it off again in File › Options › AI.",
        () => {
          on()
          void this.fromPdf(path, bytes).catch((e: unknown) => this.c.shell.alert(message(e)))
        },
      )
      return
    }
    if (!items.length)
      throw new QueryError(
        read.unread.length
          ? "The PDF's pages are scans, and text recognition isn't available here."
          : "The PDF has no text to load.",
      )
    this.navigator = {
      title: fileName(path),
      source: { kind: "pdf", path },
      items: items.map((i) => ({
        name: i.name,
        kind: i.kind,
        data: pdfResult(i),
        step: { op: "item", id: i.id },
        base: i.id,
      })),
      unread: read.unread,
    }
    for (const [n, item] of this.navigator.items.entries()) {
      const steps = this.itemPlan(this.navigator, item).steps.slice(1)
      let data: QueryResult | SourceItem[] = item.data
      for (const { step } of steps) data = applyStep(data, step)
      if (!Array.isArray(data)) item.preview = data
      const shift = steps.some((s) => s.step.op === "promote") ? 1 : 0
      const unsure = items[n].unsure?.filter(([r]) => r >= shift).map(([r, c]): [number, number] => [r - shift, c])
      if (unsure?.length) item.unsure = unsure
    }
    this.c.shell.openDialog?.("queryNavigator")
  }

  /** The plan for one of the Navigator's items: picked from its source, headings promoted, types detected. */
  private itemPlan(nav: NavigatorState, item: NavigatorItem): QueryPlan {
    const steps: QueryPlan["steps"] = [{ name: stepName(item.step, []), step: item.step }]
    let data = item.data
    if (nav.source.kind === "pdf" && hasHeadings(data)) {
      steps.push({ name: "Promoted Headers", step: { op: "promote" } })
      data = applyStep(data, { op: "promote" }) as QueryResult
    }
    steps.push(
      ...typeSteps(
        data,
        steps.map((s) => s.name),
      ),
    )
    return { source: nav.source, steps }
  }

  /** The Navigator's Transform Data: the item in the Query Editor. */
  pickTable(index: number) {
    const nav = this.navigator
    if (!nav) return
    this.navigator = null
    const item = nav.items[index]
    if (item) this.edit(item.base, this.itemPlan(nav, item))
  }

  /** The Navigator's Load: each item picked in a table on a sheet of its own (one step to undo each). */
  async loadItems(indexes: number[]): Promise<void> {
    const nav = this.navigator
    if (!nav) return
    this.navigator = null
    for (const index of indexes) {
      const item = nav.items[index]
      if (!item) continue
      const query: WorkbookQuery = {
        name: queryName(item.base, this.names()),
        description: "",
        formula: "",
        table: null,
      }
      await this.load({ query, plan: this.itemPlan(nav, item), editing: null }, true)
    }
  }

  /** Recent Sources: the source opened again (a PDF read again from its path). */
  async openRecent(source: RecentSource): Promise<void> {
    if (source.kind === "table") return this.openTable(source.name)
    if (source.kind === "web") return this.fromWeb(source.url)
    const bytes = await readFile(source.path)
    const known = this.pdfs.get(source.path)?.bytes ?? null
    if (!bytes && !known) throw new QueryError("Cells can't open it again from here. Choose Get Data › From PDF.")
    return this.fromPdf(source.path, (bytes ?? known) as Uint8Array)
  }

  /** Edit: the editor on one of the workbook's queries (only the ones Cells can run). */
  editQuery(name: string): boolean {
    const query = this.c.wb.queries.find((q) => q.name === name)
    const plan = query ? readQueryPlan(query.formula) : null
    if (!query || !plan) {
      this.c.shell.alert(
        "Cells can't change this query: it uses steps only Excel's Power Query Editor knows. It stays as it is, and Excel can still refresh it.",
      )
      return false
    }
    this.editor = { query: { ...query }, plan, editing: name }
    this.c.shell.openDialog?.("queryEditor")
    return true
  }

  /**
   * Close & Load: the editor's query saved, and its result loaded (to a table on a new sheet, or the table it already
   * fills), or kept as a connection only. One step to undo.
   */
  async load(request: EditorRequest, toTable: boolean): Promise<void> {
    const c = this.c
    const query = withPlan(request.query, request.plan)
    const others = c.wb.queries.filter((q) => q.name !== request.editing)
    if (others.some((q) => q.name.toLowerCase() === query.name.toLowerCase()))
      throw new QueryError(`There's already a query named '${query.name}'.`)
    const old = request.editing ? c.wb.queries.find((q) => q.name === request.editing) : undefined
    const table = old?.table ?? null
    const result = toTable ? await runPlan(request.plan, this.context()) : null
    if (c.edit) c.commitEdit()
    const ops: Op[] = []
    if (result && table) {
      // Renamed queries keep their table; the table's link follows the new name.
      const saved = { ...query, table }
      const list = c.wb.queries.map((q) => (q.name === request.editing ? saved : q))
      ops.push(
        { kind: "queries", queries: list },
        ...this.relinkOps(table, saved.name),
        ...refreshTableOps(c.wb, saved, result),
      )
      c.transact(ops, { label: "Load Query", before: c.selection })
    } else if (result) {
      const loaded = loadOps(c.wb, { ...query, table: null }, result, c.sheetIndex)
      // The list: the other queries and this one (an edited connection-only query's old entry goes).
      const made = loaded.ops.find((op) => op.kind === "tables")
      const tableName = made?.kind === "tables" ? made.tables[0].name : null
      const ops2 = loaded.ops.map((op): Op =>
        op.kind === "queries" ? { kind: "queries", queries: [...others, { ...query, table: tableName }] } : op,
      )
      const entry = c.transact(ops2, { label: "Load Query", before: c.selection })
      c.activate(loaded.sheet)
      // The columns fitted to the data, as Excel's query tables adjust their column widths (a heading with room
      // for its filter button).
      const fit = c.columnFit
      const widths = []
      for (let col = 0; fit && entry && col < result.columns.length; col++) {
        const heading = fit(col, { r0: 0, r1: 0 }, loaded.sheet)
        const body = fit(col, { r0: 1, r1: result.rows.length }, loaded.sheet)
        const size = Math.max(heading === null ? 0 : heading + FILTER_BUTTON_ROOM, body ?? 0) || null
        if (size !== null && size > c.sheet.cols.size(col)) widths.push({ index: col, size, hidden: false })
      }
      if (entry && widths.length)
        c.wb.amend(entry, [{ kind: "axis", sheet: loaded.sheet, axis: "cols", entries: widths }])
    } else {
      const list = [...others, { ...query, table: null }]
      c.transact([{ kind: "queries", queries: list }], { label: "Create Connection", before: c.selection })
    }
    this.editor = null
  }

  /** The table a query fills, told the query's (new) name. */
  private relinkOps(tableName: string, query: string): Op[] {
    for (const [i, s] of this.c.wb.sheets.entries()) {
      if (!s.tables.some((t) => t.name === tableName && t.query !== query)) continue
      return [{ kind: "tables", sheet: i, tables: s.tables.map((t) => (t.name === tableName ? { ...t, query } : t)) }]
    }
    return []
  }

  /** Existing Connections › Open: a connection-only query loaded to a new sheet. */
  async loadConnection(name: string): Promise<void> {
    const query = this.c.wb.queries.find((q) => q.name === name)
    const plan = query ? readQueryPlan(query.formula) : null
    if (!query || !plan) throw new QueryError("Cells can't run this query; Excel can load it.")
    await this.load({ query, plan, editing: name }, true)
  }

  /** The ops that refresh one query's table, or null when it has none or Cells can't run it. */
  async refreshOps(name: string): Promise<Op[] | null> {
    const query = this.c.wb.queries.find((q) => q.name === name)
    const plan = query?.table ? readQueryPlan(query.formula) : null
    if (!query || !plan) return null
    return refreshTableOps(this.c.wb, query, await runPlan(plan, this.context(true)))
  }

  /** Refresh: one query's table, as one step to undo. */
  async refresh(name: string): Promise<boolean> {
    const ops = await this.refreshOps(name)
    if (!ops) return false
    if (this.c.edit) this.c.commitEdit()
    this.c.transact(ops, { label: "Refresh", before: this.c.selection })
    return true
  }

  /** Refresh All's part: every loaded query Cells can run (the ones that fail are named). */
  async refreshAllOps(): Promise<{ ops: Op[]; done: number; failed: string[] }> {
    const ops: Op[] = []
    const failed: string[] = []
    let done = 0
    for (const q of this.c.wb.queries) {
      if (!q.table) continue
      try {
        const more = await this.refreshOps(q.name)
        if (!more) continue
        ops.push(...more)
        done++
      } catch {
        failed.push(q.name)
      }
    }
    return { ops, done, failed }
  }

  /** Delete: the query goes; its table stays as plain data (as in Excel). */
  remove(name: string) {
    const c = this.c
    const query = c.wb.queries.find((q) => q.name === name)
    if (!query) return
    const ops: Op[] = [{ kind: "queries", queries: c.wb.queries.filter((q) => q.name !== name) }]
    if (query.table)
      for (const [i, s] of c.wb.sheets.entries())
        if (s.tables.some((t) => t.name === query.table))
          ops.push({
            kind: "tables",
            sheet: i,
            tables: s.tables.map((t) => {
              if (t.name !== query.table) return t
              const { query: _q, ...rest } = t
              return rest
            }),
          })
    c.transact(ops, { label: "Delete Query", before: c.selection })
  }

  /** A new step's name in the editor's plan. */
  static stepName = stepName
}

export function message(e: unknown): string {
  return e instanceof Error ? e.message : String(e)
}
