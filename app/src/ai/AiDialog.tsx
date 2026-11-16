import { useEffect, useRef, useState } from "react"
import { rangeName } from "@cells/core"
import { usePlatform } from "@/platform/context"
import { useSettings } from "@/settings/settings"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "@/shell/Dialog"
import { aiTarget, LOCAL_PROVIDERS, PROVIDER_LABELS } from "./client"
import {
  chartRequest,
  dataRange,
  explainRequest,
  formulaRequest,
  readChartReply,
  readFormulaReply,
  summaryRequest,
  type AiTool,
  type ChartReply,
  type FormulaReply,
  type Request,
} from "./prompts"

const TITLES: Record<AiTool, string> = {
  formula: "Formula from Words",
  explain: "Explain Formula",
  chart: "Suggest a Chart",
  summary: "Summarize Data",
}

type Run =
  | { state: "idle" }
  | { state: "running"; text: string }
  | { state: "done"; text: string }
  | { state: "failed"; message: string }

/**
 * The AI tools (Home › AI Tools): one dialog for each, which says what it sends where, streams the answer, and puts
 * a result into the workbook only when the person chooses to (Insert, Insert Chart). When the tools are off or no
 * provider is set up, it says so and offers File › Options › AI.
 */
export function AiDialog({
  tool,
  controller: c,
  onClose,
  onOptions,
}: {
  tool: AiTool
  controller: SheetController
  onClose: () => void
  onOptions: () => void
}) {
  const platform = usePlatform()
  const settings = useSettings()
  const target = aiTarget(settings)
  const ready = platform.ai.available && settings.aiEnabled && !("problem" in target)
  const [description, setDescription] = useState("")
  const [run, setRun] = useState<Run>({ state: "idle" })
  const [formula, setFormula] = useState("")
  const stop = useRef<AbortController | null>(null)
  // What the request is about, fixed when the dialog opens.
  const [scope] = useState(() => {
    const at = c.selection.active
    return {
      cell: rangeName({ r0: at.row, c0: at.col, r1: at.row, c1: at.col }),
      range: rangeName(dataRange(c)),
      explain: tool === "explain" ? explainRequest(c) : null,
    }
  })

  const send = (request: Request | null) => {
    if (!request || "problem" in target) return
    stop.current?.abort()
    const controller = new AbortController()
    stop.current = controller
    setRun({ state: "running", text: "" })
    let text = ""
    platform.ai
      .generate(
        { ...request, target },
        (chunk) => {
          text += chunk
          if (!controller.signal.aborted) setRun({ state: "running", text })
        },
        controller.signal,
      )
      .then(
        (reply) => {
          if (controller.signal.aborted) return
          setRun({ state: "done", text: reply })
          if (tool === "formula") {
            const read = readFormulaReply(reply)
            if ("formula" in read) setFormula(read.formula)
          }
        },
        (error: unknown) => {
          if (!controller.signal.aborted)
            setRun({ state: "failed", message: error instanceof Error ? error.message : String(error) })
        },
      )
  }

  // The tools that need no words from the person start at once.
  useEffect(() => {
    if (!ready) return
    if (tool === "explain") send(scope.explain)
    else if (tool === "chart") send(chartRequest(c))
    else if (tool === "summary") send(summaryRequest(c))
    return () => stop.current?.abort()
    // Once, when the dialog opens.
  }, [])

  const close = () => {
    stop.current?.abort()
    onClose()
  }

  if (tool === "explain" && !scope.explain)
    return (
      <Dialog title={TITLES[tool]} onCancel={close} onAccept={close} cancelLabel={null} width={420}>
        <p className="dialog-note">Select a cell with a formula, then choose Explain Formula.</p>
      </Dialog>
    )

  if (!ready) {
    const why = !platform.ai.available
      ? "The AI tools need the Cells desktop app."
      : !settings.aiEnabled
        ? "The AI tools are off. Turn them on in File › Options › AI, and choose which AI to use: one on this computer, or a service you have a key for."
        : (target as { problem: string }).problem
    return (
      <Dialog
        title={TITLES[tool]}
        onCancel={close}
        onAccept={platform.ai.available ? onOptions : close}
        acceptLabel={platform.ai.available ? "Options…" : "OK"}
        cancelLabel={platform.ai.available ? "Cancel" : null}
        width={420}
      >
        <p className="dialog-note">{why}</p>
      </Dialog>
    )
  }

  const provider = settings.aiProvider as Exclude<typeof settings.aiProvider, "">
  const local = LOCAL_PROVIDERS.includes(provider)
  const sends =
    tool === "formula"
      ? `the cells near ${scope.cell}`
      : tool === "explain"
        ? `${scope.cell}'s formula and the cells around it`
        : `the cells in ${scope.range}`
  const notice = `Sends ${sends} to ${PROVIDER_LABELS[provider]}${local ? "" : ", outside this computer"}.`
  const running = run.state === "running"
  const done = run.state === "done"
  const reply = done ? run.text : ""
  const formulaReply: FormulaReply | null = tool === "formula" && done ? readFormulaReply(reply) : null
  const chartReply: ChartReply | null = tool === "chart" && done ? readChartReply(reply) : null

  const insertFormula = () => {
    const text = formula.trim()
    if (!text) return
    close()
    c.beginEdit("enter", text)
    c.commitEdit()
  }
  const insertChart = () => {
    if (!chartReply || !("chart" in chartReply)) return
    close()
    const chart = c.charts.insert(chartReply.chart)
    if (chart && chartReply.title) {
      const title = chartReply.title
      c.charts.change((spec) => ({ ...spec, title: { ...spec.title, show: true, text: title } }), "Chart Title")
    }
  }
  const copy = () => void platform.clipboard.write({ text: reply })

  let accept: { label: string; run: () => void }
  if (tool === "formula")
    accept =
      formulaReply && "formula" in formulaReply
        ? { label: "Insert", run: insertFormula }
        : { label: "Write Formula", run: () => description.trim() && send(formulaRequest(c, description)) }
  else if (tool === "chart" && chartReply && "chart" in chartReply) accept = { label: "Insert Chart", run: insertChart }
  else accept = { label: "Copy", run: copy }

  return (
    <Dialog
      title={TITLES[tool]}
      onCancel={close}
      onAccept={() => !running && accept.run()}
      width={520}
      actions={
        <>
          {running ? (
            <button
              type="button"
              className="dialog-btn"
              onClick={() => {
                stop.current?.abort()
                setRun({ state: "failed", message: "Stopped." })
              }}
            >
              Stop
            </button>
          ) : (
            <button
              type="button"
              className="dialog-btn is-primary"
              disabled={(tool !== "formula" && !done) || (tool === "formula" && !done && !description.trim())}
              onClick={accept.run}
            >
              {accept.label}
            </button>
          )}
          {done && tool !== "formula" ? (
            <button
              type="button"
              className="dialog-btn"
              onClick={() =>
                send(tool === "explain" ? scope.explain : tool === "chart" ? chartRequest(c) : summaryRequest(c))
              }
            >
              Try Again
            </button>
          ) : null}
          <button type="button" className="dialog-btn" onClick={close}>
            Close
          </button>
        </>
      }
    >
      <div className="ai-dialog">
        {tool === "formula" ? (
          <>
            <label className="dialog-label" htmlFor="ai-description">
              Describe what {scope.cell} should work out:
            </label>
            <input
              id="ai-description"
              className="dialog-field"
              data-autofocus
              value={description}
              placeholder="e.g. the total of column B where column A says Rent"
              onChange={(e) => {
                setDescription(e.target.value)
                if (done) setRun({ state: "idle" })
              }}
            />
          </>
        ) : null}
        {tool === "formula" && formulaReply ? (
          "formula" in formulaReply ? (
            <>
              <label className="dialog-label" htmlFor="ai-formula">
                Formula:
              </label>
              <input
                id="ai-formula"
                className="dialog-field ai-formula"
                value={formula}
                spellCheck={false}
                onChange={(e) => setFormula(e.target.value)}
              />
              <p className="ai-output">{formulaReply.explanation}</p>
            </>
          ) : (
            <p className="dialog-note is-error" role="alert">
              {formulaReply.problem}
            </p>
          )
        ) : null}
        {tool === "chart" && chartReply ? (
          "chart" in chartReply ? (
            <div className="ai-output" aria-label="Suggestion">
              <strong>{chartReply.label}</strong>
              {chartReply.title ? ` titled "${chartReply.title}"` : ""}
              <br />
              {chartReply.why}
            </div>
          ) : (
            <p className="dialog-note is-error" role="alert">
              {chartReply.problem}
            </p>
          )
        ) : null}
        {(tool === "explain" || tool === "summary") && (running || done) ? (
          <div className="ai-output" aria-live="polite" aria-label="Answer">
            {run.state === "running" || run.state === "done" ? run.text : ""}
          </div>
        ) : null}
        {running && tool !== "explain" && tool !== "summary" ? (
          <p className="dialog-note" aria-live="polite">
            Asking {PROVIDER_LABELS[provider]}…
          </p>
        ) : null}
        {run.state === "failed" ? (
          <p className="dialog-note is-error" role="alert">
            {run.message}
          </p>
        ) : null}
        <p className="dialog-note ai-notice">{notice} AI can be wrong: check what it writes before you rely on it.</p>
      </div>
    </Dialog>
  )
}
