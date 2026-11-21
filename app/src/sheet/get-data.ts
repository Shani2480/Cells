import {
  currentRegion,
  queryName,
  readQueryPlan,
  stepName,
  tableAt,
  type Op,
  type QueryPlan,
  type QuerySource,
  type WorkbookQuery,
} from "@cells/core"
import type { WebPort } from "@cells/platform"
import type { SheetController } from "./controller"
import {
  detectTypes,
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
export type RecentSource = { kind: "web"; url: string } | { kind: "table"; name: string }

const RECENT_KEY = "cells.recentSources"
const RECENT_MAX = 10

let webPort: WebPort | null = null
/** The app sets how files on the web are read (the desktop app fetches them itself). */
export function setWebPort(port: WebPort) {
  webPort = port
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
  const same = (s: RecentSource) =>
    s.kind === source.kind &&
    (s.kind === "web" ? s.url === (source as { url: string }).url : s.name === (source as { name: string }).name)
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

export class GetData {
  /** The Query Editor's query, while it's open. */
  editor: EditorRequest | null = null
  /** From Web's Navigator: the page's tables, while it's open. */
  navigator: { url: string; tables: QueryResult[] } | null = null

  constructor(private readonly c: SheetController) {}

  context(): QueryContext {
    return {
      wb: this.c.wb,
      fetch: (url) => {
        if (!webPort) throw new QueryError("Files on the web can't be read here.")
        return webPort.fetch(url)
      },
    }
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
        plan.steps.push({ name: "Changed Type", step: { op: "types", columns: detectTypes(r) } })
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
      this.navigator = { url, tables }
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
    steps.push({ name: "Changed Type", step: { op: "types", columns: detectTypes(data) } })
    this.edit(nameFromUrl(url), { source, steps })
  }

  /** The Navigator's choice: the page's nth table, with its types. */
  pickTable(index: number) {
    const nav = this.navigator
    if (!nav) return
    this.navigator = null
    const t = nav.tables[index]
    this.edit(`Table ${index + 1}`, {
      source: { kind: "html", url: nav.url },
      steps: [
        { name: `Data${index}`, step: { op: "table", index } },
        { name: "Changed Type", step: { op: "types", columns: detectTypes(t) } },
      ],
    })
  }

  /** Recent Sources: the source opened again. */
  openRecent(source: RecentSource): Promise<void> {
    if (source.kind === "table") {
      this.openTable(source.name)
      return Promise.resolve()
    }
    return this.fromWeb(source.url)
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
      c.transact(ops2, { label: "Load Query", before: c.selection })
      c.activate(loaded.sheet)
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
    return refreshTableOps(this.c.wb, query, await runPlan(plan, this.context()))
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
