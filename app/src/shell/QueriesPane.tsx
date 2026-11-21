import { allQueryTables, columnName, findTable, readQueryPlan } from "@cells/core"
import { Database, FileText, Pencil, RefreshCw, Settings2, Trash2, X } from "lucide-react"
import { useState, useSyncExternalStore } from "react"
import type { FileCommands } from "@/file/commands"
import type { SheetController } from "@/sheet/controller"
import { message } from "@/sheet/get-data"
import { sourceName } from "@/sheet/queries"

const rangeText = (r: { r0: number; c0: number; r1: number; c1: number }) =>
  `${columnName(r.c0)}${r.r0 + 1}:${columnName(r.c1)}${r.r1 + 1}`

/**
 * Data › Queries & Connections: the workbook's links to text files, each with where its data is, Refresh, Properties
 * and Delete (the link goes; the data stays), and Refresh All on top. Picking a link selects its data. Docked beside
 * the sheet; on a phone it takes the whole screen.
 */
export function QueriesPane({
  controller,
  files,
  onProperties,
  onDone,
}: {
  controller: SheetController
  files?: FileCommands
  onProperties: () => void
  onDone: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const [busy, setBusy] = useState<string | null>(null)
  const links = allQueryTables(controller.wb.sheets)
  const queries = controller.wb.queries
  const refreshAll = async () => {
    if (files && links.length) await files.refreshAll()
    const { ops, failed } = await controller.getData.refreshAllOps()
    if (ops.length) controller.transact(ops, { label: "Refresh All", before: controller.selection })
    if (failed.length) controller.shell.alert(`Cells couldn't refresh ${failed.join(", ")}.`)
  }
  const close = () => {
    controller.data.toggleQueries(false)
    onDone()
  }
  const run = async (key: string, work: () => Promise<unknown>) => {
    setBusy(key)
    try {
      await work()
    } finally {
      setBusy(null)
    }
  }
  return (
    <aside
      className="comments-pane queries-pane"
      aria-label="Queries & Connections"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Queries &amp; Connections</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.6} aria-hidden />
        </button>
      </header>
      <div className="clipboard-pane-actions">
        <button
          type="button"
          className="comments-pane-new"
          disabled={(!(files && links.length) && !queries.some((q) => q.table)) || busy !== null}
          onClick={() => void run("*", refreshAll)}
        >
          <RefreshCw size={16} strokeWidth={1.6} aria-hidden className={busy === "*" ? "queries-spin" : undefined} />
          Refresh All
        </button>
      </div>
      <div className="comments-pane-list" role="list">
        {queries.map((q) => {
          const key = `q:${q.name}`
          const runnable = readQueryPlan(q.formula) !== null
          const found = q.table ? findTable(controller.wb.sheets, q.table) : null
          const rows = found ? found.table.range.r1 - found.table.range.r0 : 0
          return (
            <div key={key} role="listitem" className="queries-item">
              <button
                type="button"
                className="queries-go"
                aria-label={`Query ${q.name}`}
                title={q.description || undefined}
                onClick={() => {
                  if (!found) return void controller.getData.editQuery(q.name)
                  if (found.sheet !== controller.sheetIndex) controller.activate(found.sheet)
                  controller.goTo(found.table.range)
                }}
              >
                <Database size={18} strokeWidth={1.5} aria-hidden className="queries-icon" />
                <span className="queries-text">
                  <strong>{q.name}</strong>
                  <span className="queries-meta">
                    {found ? `${rows.toLocaleString()} ${rows === 1 ? "row" : "rows"} loaded` : "Connection only."}
                  </span>
                  {found ? (
                    <span className="queries-meta">
                      {controller.wb.sheets[found.sheet].name}!{rangeText(found.table.range)}
                    </span>
                  ) : null}
                </span>
              </button>
              <span className="queries-actions">
                {found ? (
                  <button
                    type="button"
                    aria-label={`Refresh ${q.name}`}
                    title={runnable ? "Refresh" : "Only Excel can refresh this query"}
                    disabled={!runnable || busy !== null}
                    onClick={() =>
                      void run(key, () =>
                        controller.getData.refresh(q.name).catch((e: unknown) => controller.shell.alert(message(e))),
                      )
                    }
                  >
                    <RefreshCw
                      size={15}
                      strokeWidth={1.6}
                      aria-hidden
                      className={busy === key ? "queries-spin" : undefined}
                    />
                  </button>
                ) : null}
                <button
                  type="button"
                  aria-label={`Edit ${q.name}`}
                  title={runnable ? "Edit in the Query Editor" : "Only Excel can edit this query"}
                  disabled={!runnable}
                  onClick={() => controller.getData.editQuery(q.name)}
                >
                  <Pencil size={15} strokeWidth={1.6} aria-hidden />
                </button>
                <button
                  type="button"
                  aria-label={`Delete ${q.name}`}
                  title="Delete the query (its data stays)"
                  onClick={() => controller.getData.remove(q.name)}
                >
                  <Trash2 size={15} strokeWidth={1.6} aria-hidden />
                </button>
              </span>
            </div>
          )
        })}
        {!links.length && !queries.length ? (
          <p className="comments-pane-empty">
            This workbook has no connections. Data from Data › From Text/CSV, From Web and From Table/Range stays linked
            to its source and shows up here, so you can refresh it when the source changes.
          </p>
        ) : null}
        {links.map(({ sheet, table }) => {
          const key = `${sheet}:${table.name}`
          const rows = table.range.r1 - table.range.r0 + 1
          const sheetName = controller.wb.sheets[sheet]?.name ?? ""
          return (
            <div key={key} role="listitem" className="queries-item">
              <button
                type="button"
                className="queries-go"
                aria-label={`${table.connection}, ${sheetName}!${rangeText(table.range)}`}
                title={table.description || undefined}
                onClick={() => {
                  if (sheet !== controller.sheetIndex) controller.activate(sheet)
                  controller.goTo(table.range)
                }}
              >
                <FileText size={18} strokeWidth={1.5} aria-hidden className="queries-icon" />
                <span className="queries-text">
                  <strong>{table.connection}</strong>
                  <span className="queries-meta">
                    {rows.toLocaleString()} {rows === 1 ? "row" : "rows"} loaded · {sourceName(table.source)}
                  </span>
                  <span className="queries-meta">
                    {sheetName}!{rangeText(table.range)}
                  </span>
                </span>
              </button>
              <span className="queries-actions">
                <button
                  type="button"
                  aria-label={`Refresh ${table.connection}`}
                  title="Refresh"
                  disabled={!files || busy !== null}
                  onClick={() => void run(key, () => files!.refreshQuery(sheet, table.name))}
                >
                  <RefreshCw
                    size={15}
                    strokeWidth={1.6}
                    aria-hidden
                    className={busy === key ? "queries-spin" : undefined}
                  />
                </button>
                <button
                  type="button"
                  aria-label={`Properties of ${table.connection}`}
                  title="Properties"
                  onClick={() => {
                    controller.data.queryTarget = { sheet, name: table.name }
                    onProperties()
                  }}
                >
                  <Settings2 size={15} strokeWidth={1.6} aria-hidden />
                </button>
                <button
                  type="button"
                  aria-label={`Delete ${table.connection}`}
                  title="Delete the connection (the data stays)"
                  onClick={() => controller.data.setQueryTable(sheet, table.name, null)}
                >
                  <Trash2 size={15} strokeWidth={1.6} aria-hidden />
                </button>
              </span>
            </div>
          )
        })}
      </div>
    </aside>
  )
}
