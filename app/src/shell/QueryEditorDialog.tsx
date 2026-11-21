import { stepName, type FilterTest, type QueryColumnType, type QueryPlan, type QueryStep } from "@cells/core"
import {
  ArrowDownAZ,
  ArrowDownZA,
  CopyMinus,
  Funnel,
  ListChecks,
  ListStart,
  RefreshCw,
  Rows3,
  Shapes,
  TableColumnsSplit,
  X,
} from "lucide-react"
import { useEffect, useMemo, useState, useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"
import { message, type EditorRequest } from "@/sheet/get-data"
import { runPlan, type QueryResult } from "@/sheet/power-query"
import { Dialog } from "./Dialog"

/**
 * The Query Editor (Power Query's, with the buttons people use most): a preview of the data, the applied steps (pick
 * one to see the data at that point, × removes it), and buttons that add steps after the picked one: choose or
 * remove columns, keep the first rows, remove duplicates, use the first row as headers, sort, filter and change a
 * column's type. Close & Load puts the result in a table on a new sheet, or keeps the query as a connection only.
 */

const PREVIEW_ROWS = 200

const TYPES: { type: QueryColumnType; label: string; mark: string }[] = [
  { type: "text", label: "Text", mark: "ABC" },
  { type: "whole", label: "Whole Number", mark: "123" },
  { type: "number", label: "Decimal Number", mark: "1.2" },
  { type: "date", label: "Date", mark: "📅" },
  { type: "datetime", label: "Date/Time", mark: "🕒" },
  { type: "logical", label: "True/False", mark: "✓✗" },
  { type: "any", label: "Any", mark: "ABC123" },
]

const TESTS: { test: FilterTest; label: string }[] = [
  { test: "=", label: "equals" },
  { test: "<>", label: "does not equal" },
  { test: ">", label: "is greater than" },
  { test: ">=", label: "is greater than or equal to" },
  { test: "<", label: "is less than" },
  { test: "<=", label: "is less than or equal to" },
  { test: "contains", label: "contains" },
  { test: "starts", label: "begins with" },
  { test: "ends", label: "ends with" },
]

const SOURCE_LABEL: Record<QueryPlan["source"]["kind"], string> = {
  table: "Table in this workbook",
  csv: "CSV file on the web",
  html: "Web page",
  json: "JSON on the web",
}

function stepText(step: QueryStep): string {
  switch (step.op) {
    case "filter":
      return `${step.column} ${TESTS.find((t) => t.test === step.test)?.label ?? step.test} ${String(step.value)}`
    case "sort":
      return step.keys.map((k) => `${k.column} ${k.desc ? "descending" : "ascending"}`).join(", ")
    case "select":
      return `Keep ${step.columns.join(", ")}`
    case "remove":
      return `Remove ${step.columns.join(", ")}`
    case "first":
      return `First ${step.count} rows`
    case "types":
      return step.columns.map(([c, t]) => `${c}: ${TYPES.find((x) => x.type === t)?.label ?? t}`).join(", ")
    case "table":
      return `Table ${step.index + 1}`
    default:
      return ""
  }
}

type Preview = { state: "loading" } | { state: "ready"; result: QueryResult } | { state: "error"; message: string }

export function QueryEditorDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const request = controller.getData.editor
  if (!request) return null
  return <Editor controller={controller} request={request} onClose={onClose} />
}

function Editor({
  controller,
  request,
  onClose,
}: {
  controller: SheetController
  request: EditorRequest
  onClose: () => void
}) {
  const [plan, setPlan] = useState<QueryPlan>(request.plan)
  const [applied, setApplied] = useState(request.plan.steps.length)
  const [name, setName] = useState(request.query.name)
  const [description, setDescription] = useState(request.query.description)
  const loaded = !!request.query.table
  const [toTable, setToTable] = useState(request.editing ? loaded : true)
  const [column, setColumn] = useState<string | null>(null)
  const [panel, setPanel] = useState<"choose" | "filter" | "first" | "type" | null>(null)
  const [preview, setPreview] = useState<Preview>({ state: "loading" })
  const [reload, setReload] = useState(0)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const ctx = useMemo(() => controller.getData.context(), [controller])
  useEffect(() => {
    let live = true
    setPreview({ state: "loading" })
    runPlan(plan, ctx, applied).then(
      (result) => live && setPreview({ state: "ready", result }),
      (e: unknown) => live && setPreview({ state: "error", message: message(e) }),
    )
    return () => {
      live = false
    }
  }, [plan, applied, ctx, reload])

  const result = preview.state === "ready" ? preview.result : null
  const columns = result?.columns ?? []
  const selected = column && columns.includes(column) ? column : null

  /** A step added after the one shown (or changed in place when it's of the same kind on the same column). */
  const addStep = (step: QueryStep) => {
    const names = ["Source", ...plan.steps.map((s) => s.name)]
    const steps = [...plan.steps]
    steps.splice(applied, 0, { name: stepName(step, names), step })
    setPlan({ ...plan, steps })
    setApplied(applied + 1)
    setPanel(null)
  }
  const removeStep = (i: number) => {
    const steps = plan.steps.filter((_, k) => k !== i)
    setPlan({ ...plan, steps })
    setApplied(Math.min(applied, steps.length))
  }

  const accept = async () => {
    setBusy(true)
    setError(null)
    try {
      const fixed = { ...plan, steps: plan.steps }
      await controller.getData.load(
        { ...request, plan: fixed, query: { ...request.query, name: name.trim(), description } },
        toTable,
      )
      onClose()
    } catch (e) {
      setError(message(e))
    } finally {
      setBusy(false)
    }
  }
  const cancel = () => {
    controller.getData.editor = null
    onClose()
  }

  const tool = (label: string, icon: React.ReactNode, run: () => void, enabled = true, pressed = false) => (
    <button
      type="button"
      className={`qe-tool${pressed ? " is-pressed" : ""}`}
      disabled={!enabled || busy}
      onClick={run}
      title={label}
      aria-pressed={pressed || undefined}
    >
      {icon}
      <span>{label}</span>
    </button>
  )
  const needColumn = !!selected && !!result

  return (
    <Dialog
      title={`Query Editor - ${name || "Query"}`}
      onCancel={cancel}
      onAccept={() => void accept()}
      width={940}
      actions={
        <>
          <span className="qe-status" role="status">
            {preview.state === "ready"
              ? `${columns.length} column${columns.length === 1 ? "" : "s"}, ${result!.rows.length} row${result!.rows.length === 1 ? "" : "s"}`
              : preview.state === "loading"
                ? "Loading the preview…"
                : ""}
          </span>
          <button
            type="button"
            className="dialog-btn is-primary"
            disabled={busy || !name.trim()}
            onClick={() => void accept()}
          >
            {busy ? "Loading…" : "Close & Load"}
          </button>
          <button type="button" className="dialog-btn" onClick={cancel}>
            Cancel
          </button>
        </>
      }
    >
      <div className="qe">
        <div className="qe-names">
          <label>
            <span className="dialog-label">
              <u>N</u>ame
            </span>
            <input className="dialog-field" value={name} onChange={(e) => setName(e.target.value)} aria-label="Name" />
          </label>
          <label>
            <span className="dialog-label">
              <u>D</u>escription
            </span>
            <input
              className="dialog-field"
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              aria-label="Description"
            />
          </label>
        </div>
        <div className="qe-toolbar" role="toolbar" aria-label="Transform">
          {tool(
            "Choose Columns",
            <ListChecks size={18} aria-hidden />,
            () => setPanel(panel === "choose" ? null : "choose"),
            !!result,
            panel === "choose",
          )}
          {tool(
            "Remove Column",
            <TableColumnsSplit size={18} aria-hidden />,
            () => addStep({ op: "remove", columns: [selected!] }),
            needColumn,
          )}
          {tool(
            "Keep First Rows",
            <ListStart size={18} aria-hidden />,
            () => setPanel(panel === "first" ? null : "first"),
            !!result,
            panel === "first",
          )}
          {tool("Remove Duplicates", <CopyMinus size={18} aria-hidden />, () => addStep({ op: "distinct" }), !!result)}
          {tool(
            "Use First Row as Headers",
            <Rows3 size={18} aria-hidden />,
            () => addStep({ op: "promote" }),
            !!result?.rows.length,
          )}
          <span className="qe-sep" aria-hidden />
          {tool(
            "Sort Ascending",
            <ArrowDownAZ size={18} aria-hidden />,
            () => addStep({ op: "sort", keys: [{ column: selected!, desc: false }] }),
            needColumn,
          )}
          {tool(
            "Sort Descending",
            <ArrowDownZA size={18} aria-hidden />,
            () => addStep({ op: "sort", keys: [{ column: selected!, desc: true }] }),
            needColumn,
          )}
          {tool(
            "Filter Rows",
            <Funnel size={18} aria-hidden />,
            () => setPanel(panel === "filter" ? null : "filter"),
            needColumn,
            panel === "filter",
          )}
          {tool(
            "Data Type",
            <Shapes size={18} aria-hidden />,
            () => setPanel(panel === "type" ? null : "type"),
            needColumn,
            panel === "type",
          )}
          <span className="qe-sep" aria-hidden />
          {tool("Refresh Preview", <RefreshCw size={18} aria-hidden />, () => setReload(reload + 1))}
        </div>
        {panel && result ? (
          <StepPanel kind={panel} result={result} column={selected} onAdd={addStep} onCancel={() => setPanel(null)} />
        ) : null}
        <div className="qe-body">
          <div className="qe-preview" aria-label="Preview">
            {preview.state === "error" ? (
              <p className="dialog-note is-error" role="alert">
                {preview.message}
              </p>
            ) : preview.state === "loading" ? (
              <p className="dialog-note">Loading the preview…</p>
            ) : (
              <table className="qe-grid">
                <thead>
                  <tr>
                    <th className="qe-rownum" aria-hidden />
                    {result!.columns.map((c, i) => {
                      const t = TYPES.find((x) => x.type === result!.types[i])
                      return (
                        <th
                          key={c}
                          scope="col"
                          className={c === selected ? "is-selected" : undefined}
                          onClick={() => setColumn(c === selected ? null : c)}
                          aria-selected={c === selected}
                          title={t ? `${c} (${t.label})` : c}
                        >
                          <span className="qe-type" aria-hidden>
                            {t?.mark ?? "ABC123"}
                          </span>
                          {c}
                        </th>
                      )
                    })}
                  </tr>
                </thead>
                <tbody>
                  {result!.rows.slice(0, PREVIEW_ROWS).map((row, r) => (
                    <tr key={r}>
                      <td className="qe-rownum">{r + 1}</td>
                      {row.map((v, k) => (
                        <td
                          key={k}
                          className={`${typeof v === "number" ? "is-number" : ""}${v === null ? " is-null" : ""}${result!.columns[k] === selected ? " is-selected" : ""}`}
                        >
                          {v === null ? "null" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
          <aside className="qe-steps" aria-label="Applied Steps">
            <h3>Query Settings</h3>
            <p className="qe-source">{SOURCE_LABEL[plan.source.kind]}</p>
            <h4>Applied Steps</h4>
            <ol>
              {[{ name: "Source", step: null as QueryStep | null }, ...plan.steps].map((s, i) => (
                <li
                  key={`${s.name}-${i}`}
                  className={i === applied ? "is-current" : i > applied ? "is-later" : undefined}
                >
                  <button type="button" onClick={() => setApplied(i)} title={s.step ? stepText(s.step) : undefined}>
                    {s.name}
                  </button>
                  {i > 0 ? (
                    <button
                      type="button"
                      className="qe-step-remove"
                      aria-label={`Remove step ${s.name}`}
                      title="Remove this step"
                      onClick={() => removeStep(i - 1)}
                    >
                      <X size={12} aria-hidden />
                    </button>
                  ) : null}
                </li>
              ))}
            </ol>
          </aside>
        </div>
        <fieldset className="qe-load">
          <legend className="dialog-label">Load to</legend>
          <label className="fc-check">
            <input type="radio" name="qe-load" checked={toTable} onChange={() => setToTable(true)} />
            <span>{loaded ? `Table ${request.query.table}` : "Table on a new sheet"}</span>
          </label>
          <label className="fc-check">
            <input
              type="radio"
              name="qe-load"
              checked={!toTable}
              onChange={() => setToTable(false)}
              disabled={loaded}
            />
            <span>Only create connection</span>
          </label>
        </fieldset>
        {error ? (
          <p className="dialog-note is-error" role="alert">
            {error}
          </p>
        ) : null}
      </div>
    </Dialog>
  )
}

/** The panel under the toolbar for steps that ask something first. */
function StepPanel({
  kind,
  result,
  column,
  onAdd,
  onCancel,
}: {
  kind: "choose" | "filter" | "first" | "type"
  result: QueryResult
  column: string | null
  onAdd: (step: QueryStep) => void
  onCancel: () => void
}) {
  const [keep, setKeep] = useState<string[]>(result.columns)
  const [test, setTest] = useState<FilterTest>("=")
  const [value, setValue] = useState("")
  const [count, setCount] = useState("10")
  const index = column ? result.columns.indexOf(column) : -1
  const numeric =
    index >= 0 &&
    result.rows.some((r) => r[index] !== null) &&
    result.rows.every((r) => r[index] === null || typeof r[index] === "number")
  const buttons = (ok: () => void, enabled = true) => (
    <span className="qe-panel-buttons">
      <button type="button" className="dialog-btn is-primary" onClick={ok} disabled={!enabled}>
        OK
      </button>
      <button type="button" className="dialog-btn" onClick={onCancel}>
        Cancel
      </button>
    </span>
  )
  if (kind === "choose")
    return (
      <div className="qe-panel" role="group" aria-label="Choose Columns">
        <div className="qe-choose">
          {result.columns.map((c) => (
            <label key={c} className="fc-check">
              <input
                type="checkbox"
                checked={keep.includes(c)}
                onChange={(e) =>
                  setKeep(
                    e.target.checked
                      ? result.columns.filter((x) => keep.includes(x) || x === c)
                      : keep.filter((x) => x !== c),
                  )
                }
              />
              <span>{c}</span>
            </label>
          ))}
        </div>
        {buttons(() => onAdd({ op: "select", columns: keep }), keep.length > 0)}
      </div>
    )
  if (kind === "first")
    return (
      <div className="qe-panel" role="group" aria-label="Keep First Rows">
        <label>
          <span className="dialog-label">Number of rows</span>
          <input
            className="dialog-field qe-small"
            type="number"
            min={0}
            value={count}
            onChange={(e) => setCount(e.target.value)}
          />
        </label>
        {buttons(
          () => onAdd({ op: "first", count: Math.max(0, Math.floor(Number(count))) }),
          Number(count) >= 0 && count !== "",
        )}
      </div>
    )
  if (kind === "type")
    return (
      <div className="qe-panel" role="group" aria-label="Data Type">
        <span className="dialog-label">Type of {column}:</span>
        {TYPES.map((t) => (
          <button
            key={t.type}
            type="button"
            className="dialog-btn"
            onClick={() => onAdd({ op: "types", columns: [[column!, t.type]] })}
          >
            {t.label}
          </button>
        ))}
        <button type="button" className="dialog-btn" onClick={onCancel}>
          Cancel
        </button>
      </div>
    )
  const typed = numeric && value.trim() !== "" && Number.isFinite(Number(value)) ? Number(value) : value
  return (
    <div className="qe-panel" role="group" aria-label="Filter Rows">
      <span className="dialog-label">Keep rows where {column}</span>
      <select
        className="dialog-field qe-small"
        value={test}
        onChange={(e) => setTest(e.target.value as FilterTest)}
        aria-label="Test"
      >
        {TESTS.map((t) => (
          <option key={t.test} value={t.test}>
            {t.label}
          </option>
        ))}
      </select>
      <input
        className="dialog-field"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        aria-label="Value"
        autoFocus
      />
      {buttons(() => onAdd({ op: "filter", column: column!, test, value: typed }))}
    </div>
  )
}
