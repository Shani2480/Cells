import {
  allQueryTables,
  columnName,
  readQueryPlan,
  selectRange,
  type QueryColumnType,
  type QueryValue,
} from "@cells/core"
import { Database, File, FileText, Table2 } from "lucide-react"
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

/**
 * The Navigator: a web page's tables, or a PDF's tables and pages, each with a preview. Load puts the item (or every
 * item ticked, with Select multiple items) in a table on a sheet of its own; Transform Data opens it in the Query
 * Editor first.
 */
export function NavigatorDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const nav = controller.getData.navigator
  const [picked, setPicked] = useState(0)
  const [multiple, setMultiple] = useState(false)
  const [ticked, setTicked] = useState<number[]>([])
  if (!nav) return null
  const item = nav.items[picked]
  const table = item.preview ?? item.data
  const doubtful = new Set((item.unsure ?? []).map(([r, c]) => `${r}:${c}`))
  const chosen = multiple ? ticked : [picked]
  const cancel = () => {
    controller.getData.navigator = null
    onClose()
  }
  const transform = () => {
    onClose()
    controller.getData.pickTable(multiple && ticked.length ? ticked[0] : picked)
  }
  const load = () => {
    if (!chosen.length) return
    onClose()
    controller.getData
      .loadItems(chosen)
      .catch((e: unknown) => controller.shell.alert(`Cells couldn't load the data. ${message(e)}`))
  }
  const tick = (i: number) =>
    setTicked((t) => (t.includes(i) ? t.filter((k) => k !== i) : [...t, i].sort((a, b) => a - b)))
  const pdf = nav.source.kind === "pdf"
  return (
    <Dialog
      title="Navigator"
      onCancel={cancel}
      onAccept={load}
      width={780}
      actions={
        <>
          {multiple ? (
            <span className="qe-status">
              {ticked.length} item{ticked.length === 1 ? "" : "s"} selected
            </span>
          ) : null}
          <button type="button" className="dialog-btn is-primary" onClick={load} disabled={!chosen.length}>
            <u>L</u>oad
          </button>
          <button type="button" className="dialog-btn" onClick={transform} disabled={!chosen.length}>
            <u>T</u>ransform Data
          </button>
          <button type="button" className="dialog-btn" onClick={cancel}>
            Cancel
          </button>
        </>
      }
    >
      <div className="nav-dialog">
        <div className="nav-side">
          <label className="dialog-radio nav-multiple">
            <input
              type="checkbox"
              checked={multiple}
              onChange={(e) => {
                setMultiple(e.target.checked)
                setTicked(e.target.checked ? [picked] : [])
              }}
            />
            Select multiple items
          </label>
          <div className="dialog-list nav-list" role="listbox" aria-label={pdf ? "Tables and pages" : "Tables"}>
            <p className="nav-site" title={nav.title}>
              {pdf ? <FileText size={13} aria-hidden /> : null} {nav.title} [{nav.items.length}]
            </p>
            {nav.items.map((t, i) => (
              <button
                key={i}
                type="button"
                role="option"
                aria-selected={i === picked}
                className={`dialog-list-item${i === picked ? " is-selected" : ""}`}
                onClick={() => {
                  setPicked(i)
                  if (multiple) tick(i)
                }}
                onDoubleClick={transform}
              >
                {multiple ? (
                  <input
                    type="checkbox"
                    tabIndex={-1}
                    aria-label={`Select ${t.name}`}
                    checked={ticked.includes(i)}
                    readOnly
                  />
                ) : null}
                {t.kind === "Page" ? <File size={14} aria-hidden /> : <Table2 size={14} aria-hidden />} {t.name}
                <span className="nav-size">
                  {t.data.rows.length} × {t.data.columns.length}
                </span>
              </button>
            ))}
          </div>
          {nav.unread.length ? (
            <p className="dialog-note nav-unread">
              {nav.unread.length === 1 ? `Page ${nav.unread[0]} is a scan` : `${nav.unread.length} pages are scans`} and
              weren&apos;t read.{" "}
              {controller.getData.canTurnOnRecognition ? (
                <button
                  type="button"
                  className="link-button"
                  onClick={() => {
                    onClose()
                    controller.getData
                      .readScans()
                      .catch((e: unknown) => controller.shell.alert(`Cells couldn't read the scans. ${message(e)}`))
                  }}
                >
                  Read them with text recognition
                </button>
              ) : null}
            </p>
          ) : null}
          {item.unsure?.length ? (
            <p className="dialog-note nav-unsure">
              <span className="nav-unsure-swatch" aria-hidden /> Check the marked cells: text recognition wasn&apos;t
              sure of them.
            </p>
          ) : null}
        </div>
        <div className="nav-pane">
          <p className="nav-pane-title">{item.name}</p>
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
                      <td
                        key={k}
                        className={
                          [
                            v === null ? "is-null" : typeof v === "number" ? "is-number" : "",
                            doubtful.has(`${i}:${k}`) ? "is-unsure" : "",
                          ]
                            .filter(Boolean)
                            .join(" ") || undefined
                        }
                        title={doubtful.has(`${i}:${k}`) ? "Text recognition wasn't sure of this" : undefined}
                      >
                        {shown(v, table.types[k])}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </Dialog>
  )
}

/** A preview value: dates as dates, numbers as they are, nothing for null. */
function shown(v: QueryValue, type: QueryColumnType | null | undefined): string {
  if (v === null) return ""
  if (typeof v === "number" && (type === "date" || type === "datetime")) {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000)
    return type === "date" ? d.toISOString().slice(0, 10) : d.toISOString().slice(0, 16).replace("T", " ")
  }
  return typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)
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
