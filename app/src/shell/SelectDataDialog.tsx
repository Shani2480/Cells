import { useState, useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

export const BAD_DATA_RANGE = "The reference isn't valid. References must be to an open worksheet."

/**
 * Chart Design › Select Data (Excel's Select Data Source): the chart's data range, Switch Row/Column, and its series
 * and category labels. Each change applies at once, as in Excel; Cancel takes them all back.
 */
export function SelectDataDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const charts = controller.charts
  const [startDepth] = useState(() => controller.wb.undoDepth)
  const [text, setText] = useState(() => charts.dataRange())
  const [picked, setPicked] = useState(0)
  const chart = charts.active()
  const spec = chart ? charts.liveSpec(chart) : null
  if (!chart || !spec) return null

  const apply = (byRow?: boolean) => {
    if (text === charts.dataRange() && byRow === undefined) return true
    if (!charts.setData(text, byRow)) {
      controller.shell.alert(BAD_DATA_RANGE)
      return false
    }
    setText(charts.dataRange())
    return true
  }
  const cancel = () => {
    // Back to the chart as it was when the dialog opened.
    while (controller.wb.undoDepth > startDepth) controller.wb.undo()
    onClose()
  }
  const switched = () => {
    if (!apply()) return
    charts.switchRowColumn()
    setText(charts.dataRange())
  }

  return (
    <Dialog title="Select Data Source" onCancel={cancel} onAccept={() => apply() && onClose()} width={560}>
      <label className="dialog-label" htmlFor="sd-range">
        Chart <u>d</u>ata range:
      </label>
      <input
        id="sd-range"
        className="dialog-field"
        value={text}
        spellCheck={false}
        onChange={(e) => setText(e.target.value)}
        onBlur={() => apply()}
      />
      <div className="sd-switch">
        <button type="button" className="dialog-btn" onClick={switched} disabled={!chart.refs}>
          S<u>w</u>itch Row/Column
        </button>
      </div>
      <div className="sd-lists">
        <div>
          <div className="dialog-label">
            Legend Entries (<u>S</u>eries)
          </div>
          <div className="sd-actions">
            <button
              type="button"
              className="dialog-btn"
              disabled={spec.series.length <= 1 || !chart.refs}
              onClick={() => {
                charts.removeSeries(picked)
                setPicked(Math.max(0, picked - 1))
                setText(charts.dataRange())
              }}
            >
              <u>R</u>emove
            </button>
          </div>
          <div className="dialog-list sd-list" role="listbox" aria-label="Series">
            {spec.series.map((s, i) => (
              <button
                key={i}
                type="button"
                role="option"
                aria-selected={i === picked}
                className={`dialog-list-item${i === picked ? " is-selected" : ""}`}
                onClick={() => setPicked(i)}
              >
                {s.name}
              </button>
            ))}
          </div>
        </div>
        <div>
          <div className="dialog-label">
            Horizontal (<u>C</u>ategory) Axis Labels
          </div>
          <div className="sd-actions" />
          <div className="dialog-list sd-list" role="listbox" aria-label="Category labels">
            {spec.categories.map((c, i) => (
              <div key={i} className="dialog-list-item" role="option" aria-selected={false}>
                {c}
              </div>
            ))}
          </div>
        </div>
      </div>
    </Dialog>
  )
}
