import { allQueryTables, columnName, readQueryPlan, selectRange } from "@cells/core"
import { Database, FileText, Table2 } from "lucide-react"
import { useState, useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"
import { message } from "@/sheet/get-data"
import { sourceName } from "@/sheet/queries"
import { Dialog } from "./Dialog"

/** From Web: a web address (CSV, JSON or a page with tables). */
export function FromWebDialog({
  controller,
  initial = "",
  onClose,
}: {
  controller: SheetController
  initial?: string
  onClose: () => void
}) {
  const [url, setUrl] = useState(initial)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const address = url.trim()
  const full = address && !/^[a-z]+:\/\//i.test(address) ? `https://${address}` : address
  const accept = async () => {
    if (!full) return
    setBusy(true)
    setError(null)
    try {
      onClose()
      await controller.getData.fromWeb(full)
    } catch (e) {
      controller.shell.alert(`Cells couldn't read ${full}. ${message(e)}`)
    } finally {
      setBusy(false)
    }
  }
  return (
    <Dialog
      title="From Web"
      onCancel={onClose}
      onAccept={() => void accept()}
      width={460}
      acceptLabel={busy ? "Reading…" : "OK"}
    >
      <label className="dialog-label" htmlFor="web-url">
        <u>U</u>RL
      </label>
      <input
        id="web-url"
        className="dialog-field"
        autoFocus
        placeholder="https://example.com/data.csv"
        value={url}
        onChange={(e) => setUrl(e.target.value)}
      />
      <p className="dialog-note">
        A CSV or JSON file, or a web page with tables. Cells reads it now and again each time you refresh.
      </p>
      {error ? (
        <p className="dialog-note is-error" role="alert">
          {error}
        </p>
      ) : null}
    </Dialog>
  )
}

/** The Navigator: a web page's tables, each with a preview; the one picked goes to the Query Editor. */
export function NavigatorDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const nav = controller.getData.navigator
  const [picked, setPicked] = useState(0)
  if (!nav) return null
  const table = nav.tables[picked]
  const cancel = () => {
    controller.getData.navigator = null
    onClose()
  }
  const accept = () => {
    onClose()
    controller.getData.pickTable(picked)
  }
  return (
    <Dialog title="Navigator" onCancel={cancel} onAccept={accept} width={760} acceptLabel="Transform Data">
      <div className="nav-dialog">
        <div className="dialog-list nav-list" role="listbox" aria-label="Tables">
          <p className="nav-site" title={nav.url}>
            {nav.url}
          </p>
          {nav.tables.map((t, i) => (
            <button
              key={i}
              type="button"
              role="option"
              aria-selected={i === picked}
              className={`dialog-list-item${i === picked ? " is-selected" : ""}`}
              onClick={() => setPicked(i)}
              onDoubleClick={accept}
            >
              <Table2 size={14} aria-hidden /> Table {i + 1}
              <span className="nav-size">
                {t.rows.length} × {t.columns.length}
              </span>
            </button>
          ))}
        </div>
        <div className="qe-preview nav-preview" aria-label="Preview">
          <table className="qe-grid">
            <thead>
              <tr>
                {table.columns.map((c) => (
                  <th key={c}>{c}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {table.rows.slice(0, 50).map((r, i) => (
                <tr key={i}>
                  {r.map((v, k) => (
                    <td key={k} className={v === null ? "is-null" : undefined}>
                      {v === null ? "" : String(v)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </Dialog>
  )
}

const rangeText = (r: { r0: number; c0: number; r1: number; c1: number }) =>
  `${columnName(r.c0)}${r.r0 + 1}:${columnName(r.c1)}${r.r1 + 1}`

/**
 * Existing Connections: the workbook's connections. Open loads a connection-only query into a table on a new sheet,
 * or goes to the data a query or text link fills.
 */
export function ExistingConnectionsDialog({
  controller,
  onClose,
}: {
  controller: SheetController
  onClose: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const wb = controller.wb
  type Item = { key: string; name: string; detail: string; kind: "query" | "link"; runnable: boolean }
  const items: Item[] = [
    ...wb.queries.map((q) => ({
      key: `q:${q.name}`,
      name: `Query - ${q.name}`,
      detail: q.description || (q.table ? `Loads table ${q.table}` : "Connection only"),
      kind: "query" as const,
      runnable: readQueryPlan(q.formula) !== null,
    })),
    ...allQueryTables(wb.sheets).map(({ sheet, table }) => ({
      key: `l:${sheet}:${table.name}`,
      name: table.connection,
      detail: `${sourceName(table.source)}, ${wb.sheets[sheet].name}!${rangeText(table.range)}`,
      kind: "link" as const,
      runnable: true,
    })),
  ]
  const [picked, setPicked] = useState(items[0]?.key ?? "")
  const [error, setError] = useState<string | null>(null)
  const current = items.find((i) => i.key === picked)
  const open = async () => {
    if (!current) return
    setError(null)
    if (current.kind === "link") {
      const [, sheet, name] = current.key.split(":")
      const t = wb.sheets[Number(sheet)]?.queryTables.find((q) => q.name === name)
      if (t) {
        controller.activate(Number(sheet))
        controller.setSelection(selectRange(t.range))
      }
      onClose()
      return
    }
    const query = wb.queries.find((q) => `q:${q.name}` === current.key)
    if (!query) return
    if (query.table) {
      for (const [i, s] of wb.sheets.entries()) {
        const t = s.tables.find((x) => x.name === query.table)
        if (!t) continue
        controller.activate(i)
        controller.setSelection(selectRange(t.range))
      }
      onClose()
      return
    }
    try {
      await controller.getData.loadConnection(query.name)
      onClose()
    } catch (e) {
      setError(message(e))
    }
  }
  return (
    <Dialog title="Existing Connections" onCancel={onClose} onAccept={() => void open()} width={520} acceptLabel="Open">
      <p className="dialog-label">Connections in this Workbook</p>
      <div className="dialog-list existing-list" role="listbox" aria-label="Connections">
        {items.length ? (
          items.map((i) => (
            <button
              key={i.key}
              type="button"
              role="option"
              aria-selected={i.key === picked}
              className={`dialog-list-item existing-item${i.key === picked ? " is-selected" : ""}`}
              onClick={() => setPicked(i.key)}
              onDoubleClick={() => void open()}
            >
              {i.kind === "query" ? <Database size={16} aria-hidden /> : <FileText size={16} aria-hidden />}
              <span>
                <strong>{i.name}</strong>
                <span className="existing-detail">{i.detail}</span>
              </span>
            </button>
          ))
        ) : (
          <p className="dialog-note">This workbook has no connections yet. Get Data makes them.</p>
        )}
      </div>
      {current?.kind === "query" && !current.runnable ? (
        <p className="dialog-note">Cells can't run this query (Excel can); Open goes to its data.</p>
      ) : null}
      {error ? (
        <p className="dialog-note is-error" role="alert">
          {error}
        </p>
      ) : null}
    </Dialog>
  )
}
