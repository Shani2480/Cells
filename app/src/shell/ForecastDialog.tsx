import { useMemo, useState } from "react"
import { ChartColumn, ChartLine, ChevronDown, ChevronRight } from "lucide-react"
import { columnName, currentRegion, displayText, ETS_AGGREGATIONS, parseInput, type Range } from "@cells/core"
import { renderChart } from "@cells/chart"
import type { SheetController } from "@/sheet/controller"
import {
  createForecastSheet,
  forecastPreview,
  futureTimes,
  historySeries,
  readHistory,
  type ForecastOptions,
} from "@/sheet/forecast"
import { Dialog } from "./Dialog"
import { parseSheetRange } from "./AdvancedFilterDialog"

const abs = (sheetName: string, r: Range) => {
  const quoted = /^[A-Za-z_][\w.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`
  return `=${quoted}!$${columnName(r.c0)}$${r.r0 + 1}:$${columnName(r.c1)}$${r.r1 + 1}`
}

/** The timeline and values columns Forecast Sheet starts from: the selection, or the data around the active cell. */
function initialRanges(controller: SheetController): { timeline: Range; values: Range } | null {
  const sel = controller.selection.ranges[0]
  const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
  const area = single ? currentRegion(controller.sheet, controller.selection.active) : sel
  if (area.c1 <= area.c0 || area.r1 <= area.r0) return null
  return {
    timeline: { r0: area.r0, r1: area.r1, c0: area.c0, c1: area.c0 },
    values: { r0: area.r0, r1: area.r1, c0: area.c0 + 1, c1: area.c0 + 1 },
  }
}

/**
 * Data › Forecast › Forecast Sheet ("Create Forecast Worksheet"): a preview of the forecast from a timeline and its
 * values, its end date and options, and Create, which makes the forecast sheet.
 */
export function ForecastDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const name = controller.sheet.name
  const initial = useMemo(() => initialRanges(controller), [controller])
  const [timelineText, setTimelineText] = useState(initial ? abs(name, initial.timeline) : "")
  const [valuesText, setValuesText] = useState(initial ? abs(name, initial.values) : "")
  const [chart, setChart] = useState<"line" | "column">("line")
  const [showOptions, setShowOptions] = useState(false)
  const [confidenceOn, setConfidenceOn] = useState(true)
  const [confidenceText, setConfidenceText] = useState("95%")
  const [autoSeason, setAutoSeason] = useState(true)
  const [season, setSeason] = useState(2)
  const [stats, setStats] = useState(false)
  const [completion, setCompletion] = useState(1)
  const [aggregation, setAggregation] = useState(1)

  // The history in the ranges as typed.
  const read = useMemo(() => {
    const t = parseSheetRange(controller, timelineText)
    const v = parseSheetRange(controller, valuesText)
    if (!t || !v) return { problem: "Reference isn't valid." }
    if (t.sheet !== v.sheet) return { problem: "The timeline and values ranges must be on the same sheet." }
    const r = readHistory(controller, { sheet: t.sheet, timeline: t.range, values: v.range })
    return "problem" in r
      ? { problem: r.problem }
      : { history: r.history, sheet: t.sheet, timeline: t.range, values: v.range }
  }, [controller, timelineText, valuesText])

  const history = "history" in read ? read.history : null
  const format = (t: number) => (history ? displayText(t, history.timeFormat) : String(t))
  const parseTime = (text: string): number | null => {
    const v = parseInput(text.trim())?.value
    return typeof v === "number" ? v : null
  }

  // Default end: about a third of the history again; start: its last point.
  const defaults = useMemo(() => {
    if (!history) return null
    const full = historySeries(history, { completion: 1, aggregation: 1 })
    if ("problem" in full) return null
    const steps = Math.max(1, Math.ceil(full.series.values.length / 3))
    const times = futureTimes(full.series, Infinity, steps)
    return { start: history.points[history.points.length - 1].t, end: times[times.length - 1] }
  }, [history])
  const [endText, setEndText] = useState<string | null>(null)
  const [startText, setStartText] = useState<string | null>(null)
  const endShown = endText ?? (defaults ? format(defaults.end) : "")
  const startShown = startText ?? (defaults ? format(defaults.start) : "")

  const confidence = (() => {
    if (!confidenceOn) return null
    const v = parseInput(confidenceText.trim())?.value
    if (typeof v !== "number") return NaN
    return v >= 1 ? v / 100 : v
  })()

  const options = (): ForecastOptions | string => {
    if (!("history" in read) || !read.history) return read.problem ?? "Reference isn't valid."
    const end = parseTime(endShown)
    const start = parseTime(startShown)
    if (end === null) return "The forecast end must be a date or a number."
    if (start === null) return "The forecast start must be a date or a number."
    if (confidence !== null && !(confidence > 0 && confidence < 1))
      return "The confidence interval must be between 0% and 100%."
    return {
      sheet: read.sheet,
      timeline: read.timeline,
      values: read.values,
      end,
      start,
      confidence,
      seasonality: autoSeason ? 1 : season,
      stats,
      completion,
      aggregation,
      chart,
    }
  }

  const preview = useMemo(() => {
    const o = options()
    if (typeof o === "string" || !history) return { problem: o as string }
    const p = forecastPreview(history, o)
    if ("problem" in p) return p
    try {
      return { svg: renderChart(p.spec, 560, 250) }
    } catch {
      return { problem: "" }
    }
  }, [read, history, endShown, startShown, confidence, autoSeason, season, completion, aggregation, chart])

  const create = () => {
    const o = options()
    if (typeof o === "string") return controller.shell.alert(o)
    if (!history) return
    const problem = createForecastSheet(controller, history, o)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }

  return (
    <Dialog title="Create Forecast Worksheet" onCancel={onClose} onAccept={create} acceptLabel="Create" width={620}>
      <div className="forecast-types" role="radiogroup" aria-label="Chart type">
        <button
          type="button"
          role="radio"
          aria-checked={chart === "line"}
          aria-label="Create a line chart"
          title="Create a line chart"
          className={`forecast-type${chart === "line" ? " is-selected" : ""}`}
          onClick={() => setChart("line")}
        >
          <ChartLine size={20} strokeWidth={1.5} />
        </button>
        <button
          type="button"
          role="radio"
          aria-checked={chart === "column"}
          aria-label="Create a column chart"
          title="Create a column chart"
          className={`forecast-type${chart === "column" ? " is-selected" : ""}`}
          onClick={() => setChart("column")}
        >
          <ChartColumn size={20} strokeWidth={1.5} />
        </button>
      </div>
      <div className="forecast-preview" aria-label="Preview">
        {"svg" in preview ? (
          <svg
            className="forecast-chart"
            viewBox="0 0 560 250"
            role="img"
            aria-label="Forecast preview"
            dangerouslySetInnerHTML={{ __html: preview.svg }}
          />
        ) : (
          <p className="dialog-hint">{preview.problem}</p>
        )}
      </div>
      <div className="forecast-end">
        <label className="dialog-label" htmlFor="forecast-end">
          Forecast <u>E</u>nd
        </label>
        <input
          id="forecast-end"
          className="dialog-field"
          value={endShown}
          onChange={(e) => setEndText(e.target.value)}
        />
      </div>
      <button
        type="button"
        className="forecast-options-toggle"
        aria-expanded={showOptions}
        onClick={() => setShowOptions(!showOptions)}
      >
        {showOptions ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
        <span>
          <u>O</u>ptions
        </span>
      </button>
      {showOptions ? (
        <div className="forecast-options">
          <div className="forecast-column">
            <label className="dialog-label" htmlFor="forecast-start">
              Forecast <u>S</u>tart
            </label>
            <input
              id="forecast-start"
              className="dialog-field"
              value={startShown}
              onChange={(e) => setStartText(e.target.value)}
            />
            <div className="forecast-inline">
              <label className="dialog-radio">
                <input type="checkbox" checked={confidenceOn} onChange={(e) => setConfidenceOn(e.target.checked)} />
                <span>
                  Confidence <u>I</u>nterval
                </span>
              </label>
              <input
                className="dialog-field forecast-small"
                aria-label="Confidence interval"
                disabled={!confidenceOn}
                value={confidenceText}
                onChange={(e) => setConfidenceText(e.target.value)}
              />
            </div>
            <fieldset className="dialog-group">
              <legend>Seasonality</legend>
              <label className="dialog-radio">
                <input type="radio" name="forecast-season" checked={autoSeason} onChange={() => setAutoSeason(true)} />
                <span>
                  Detect <u>A</u>utomatically
                </span>
              </label>
              <div className="forecast-inline">
                <label className="dialog-radio">
                  <input
                    type="radio"
                    name="forecast-season"
                    checked={!autoSeason}
                    onChange={() => setAutoSeason(false)}
                  />
                  <span>
                    Set <u>M</u>anually
                  </span>
                </label>
                <input
                  type="number"
                  className="dialog-field forecast-small"
                  aria-label="Season length"
                  min={2}
                  max={8760}
                  disabled={autoSeason}
                  value={season}
                  onChange={(e) => setSeason(Math.max(2, Math.min(8760, Math.trunc(Number(e.target.value)) || 2)))}
                />
              </div>
            </fieldset>
            <label className="dialog-radio">
              <input type="checkbox" checked={stats} onChange={(e) => setStats(e.target.checked)} />
              <span>
                Include forecast <u>s</u>tatistics
              </span>
            </label>
          </div>
          <div className="forecast-column">
            <label className="dialog-label" htmlFor="forecast-timeline">
              <u>T</u>imeline Range
            </label>
            <input
              id="forecast-timeline"
              className="dialog-field"
              value={timelineText}
              onChange={(e) => setTimelineText(e.target.value)}
            />
            <label className="dialog-label" htmlFor="forecast-values">
              <u>V</u>alues Range
            </label>
            <input
              id="forecast-values"
              className="dialog-field"
              value={valuesText}
              onChange={(e) => setValuesText(e.target.value)}
            />
            <label className="dialog-label" htmlFor="forecast-fill">
              <u>F</u>ill Missing Points Using
            </label>
            <select
              id="forecast-fill"
              className="dialog-field"
              value={completion}
              onChange={(e) => setCompletion(Number(e.target.value))}
            >
              <option value={1}>Interpolation</option>
              <option value={0}>Zeros</option>
            </select>
            <label className="dialog-label" htmlFor="forecast-aggregate">
              Aggregate <u>D</u>uplicates Using
            </label>
            <select
              id="forecast-aggregate"
              className="dialog-field"
              value={aggregation}
              onChange={(e) => setAggregation(Number(e.target.value))}
            >
              {ETS_AGGREGATIONS.map((label, i) => (
                <option key={label} value={i + 1}>
                  {label}
                </option>
              ))}
            </select>
          </div>
        </div>
      ) : null}
    </Dialog>
  )
}
