import { cachedLinkRows } from "@/sheet/links"
import type { Command, RibbonEnv } from "./registry"

/** Data › Sort & Filter › Advanced, and the Data Tools that come with Part 4. */
export function dataToolCommands(): Command[] {
  return [
    {
      id: "AdvancedFilterDialog",
      label: "Advanced...",
      description: "Filter with a range of conditions you write on the sheet, in place or as a copy somewhere else.",
      keytip: "Q",
      run: (env) => env.dialog("advancedFilter"),
    },
    {
      id: "Consolidate",
      label: "Consolidate...",
      description:
        "Summarise data from separate ranges, on this or other sheets, into one table here, by position or by its labels.",
      keytip: "N",
      run: (env) => env.dialog("consolidate"),
    },
    {
      id: "ForecastInsert",
      label: "Forecast Sheet",
      description:
        "Create a new sheet that predicts where your data over time is going, with a table of forecasts and a chart.",
      keytip: "FC",
      run: (env) => env.dialog("forecast"),
    },
    {
      id: "RefreshAllMenu",
      label: "Refresh All",
      description: "Get the latest data by reading every linked text file and PivotTable source in the workbook again.",
      shortcut: "Ctrl+Alt+F5",
      keytip: "RA",
      run: (env) => void refreshAll(env),
      enabled: (env) =>
        (Boolean(env.files) && hasLinks(env)) || env.c.pivots.any() || env.c.wb.queries.some((q) => q.table),
      menu: ["RefreshAll", "RefreshData", "-", "ConnectionProperties"],
    },
    {
      id: "RefreshAll",
      label: "Refresh All",
      description: "Get the latest data by reading every linked text file and PivotTable source in the workbook again.",
      shortcut: "Ctrl+Alt+F5",
      keytip: "A",
      run: (env) => void refreshAll(env),
      enabled: (env) =>
        (Boolean(env.files) && hasLinks(env)) || env.c.pivots.any() || env.c.wb.queries.some((q) => q.table),
    },
    {
      id: "RefreshData",
      label: "Refresh",
      description: "Get the latest data from the file or PivotTable source of the selected cells.",
      shortcut: "Alt+F5",
      keytip: "R",
      run: (env) => {
        if (env.c.pivots.active()) return void env.c.pivots.refresh()
        const loaded = activeQuery(env)
        if (loaded)
          return void env.c.getData.refresh(loaded).then(
            (ok) =>
              ok ||
              env.c.shell.alert("Cells can't run this query (it uses steps only Excel knows); Excel can refresh it."),
            (e: unknown) =>
              env.c.shell.alert(`Cells couldn't refresh ${loaded}. ${e instanceof Error ? e.message : String(e)}`),
          )
        const q = env.c.data.activeQuery()
        if (q) void env.files?.refreshQuery(q.sheet, q.name)
      },
      enabled: (env) =>
        env.c.pivots.active() !== null ||
        activeQuery(env) !== null ||
        (Boolean(env.files) && env.c.data.activeQuery() !== null),
    },
    {
      id: "ConnectionProperties",
      label: "Connection Properties...",
      description: "Name the connection, describe it, and choose whether it refreshes when the workbook opens.",
      keytip: "O",
      run: (env) => openProperties(env),
      enabled: (env) => env.c.data.activeQuery() !== null,
    },
    {
      id: "ShowQueriesAndConnections",
      label: "Queries & Connections",
      description: "Show the workbook's links to text files, where their data is, and refresh or change them.",
      keytip: "O",
      run: (env) => env.c.data.toggleQueries(),
      pressed: (env) => env.c.data.queriesOpen,
    },
    {
      id: "WorkbookLinks",
      label: "Workbook Links",
      description:
        "See the other workbooks this one's formulas use, update their values, change where they come from or break the links.",
      keytip: "K",
      run: (env) => env.c.data.toggleLinks(),
      pressed: (env) => env.c.data.linksOpen,
      enabled: (env) => hasBookLinks(env) || hasLinkedFormulas(env),
    },
    {
      id: "DataRangeProperties",
      label: "Properties",
      description: "Change how the linked data in the selected cells refreshes, and its names.",
      keytip: "P",
      run: (env) => openProperties(env),
      enabled: (env) => env.c.data.activeQuery() !== null,
    },
  ]
}

const hasBookLinks = (env: RibbonEnv) => env.c.wb.links.list.length > 0 || env.c.data.linksOpen

const hasLinks = (env: RibbonEnv) => env.c.wb.sheets.some((s) => s.queryTables.length > 0)

/** Linked text files first (a PivotTable may count their data), then the PivotTables. */
async function refreshAll(env: RibbonEnv) {
  if (env.files && hasLinks(env)) await env.files.refreshAll()
  await refreshQueries(env)
  env.c.pivots.refreshAll()
}

/** Refresh All's queries: each loaded one Cells can run, in one step to undo; the ones that fail are named. */
async function refreshQueries(env: RibbonEnv) {
  const c = env.c
  if (!c.wb.queries.some((q) => q.table)) return
  const { ops, failed } = await c.getData.refreshAllOps()
  if (ops.length) {
    if (c.edit) c.commitEdit()
    c.transact(ops, { label: "Refresh All", before: c.selection })
  }
  if (failed.length)
    c.shell.alert(`Cells couldn't refresh ${failed.length === 1 ? "the query" : "the queries"} ${failed.join(", ")}.`)
}

/** The loaded query of the active cell's table. */
const activeQuery = (env: RibbonEnv) => {
  const { row, col } = env.c.selection.active
  const t = env.c.sheet.tables.find(
    (x) => x.query && row >= x.range.r0 && row <= x.range.r1 && col >= x.range.c0 && col <= x.range.c1,
  )
  return t?.query ?? null
}

function openProperties(env: RibbonEnv) {
  const q = env.c.data.activeQuery()
  if (!q) return
  env.c.data.queryTarget = q
  env.dialog("queryProperties")
}

/** Formulas that name a workbook with no link yet (typed, not yet updated). Checked only when there's no link. */
const hasLinkedFormulas = (env: RibbonEnv) => cachedLinkRows(env.c.wb).length > 0
