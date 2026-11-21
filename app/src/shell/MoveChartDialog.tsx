import { nextChartSheetName } from "@cells/core"
import { useId, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * Chart Design › Move Chart: a chart sheet of its own (named, Chart1 to start) or an object on another worksheet, as
 * Excel's dialog offers.
 */
export function MoveChartDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const wb = controller.wb
  const onChartSheet = controller.sheet.chartSheet
  const worksheets = wb.sheets.flatMap((s, i) => (!s.chartSheet && !s.hidden ? [{ name: s.name, index: i }] : []))
  const others = worksheets.filter((w) => w.index !== controller.sheetIndex)
  const [where, setWhere] = useState<"sheet" | "object">(onChartSheet ? "object" : "sheet")
  const [name, setName] = useState(
    onChartSheet ? controller.sheet.name : nextChartSheetName(wb.sheets.map((s) => s.name)),
  )
  const [target, setTarget] = useState(others[0]?.index ?? -1)
  const id = useId()
  const accept = () => {
    const ok = where === "sheet" ? controller.charts.moveToNewSheet(name) : controller.charts.moveToSheet(target)
    if (ok) onClose()
  }
  return (
    <Dialog title="Move Chart" onCancel={onClose} onAccept={accept} width={460}>
      <p className="dialog-label">Choose where you want the chart to be placed:</p>
      <div className="move-chart">
        <label className="dialog-radio">
          <input type="radio" name={id} checked={where === "sheet"} onChange={() => setWhere("sheet")} />
          <span>
            <u>N</u>ew sheet:
          </span>
        </label>
        <input
          className="dialog-field"
          aria-label="New sheet name"
          value={name}
          disabled={where !== "sheet"}
          onChange={(e) => setName(e.target.value)}
          onFocus={() => setWhere("sheet")}
        />
        <label className="dialog-radio">
          <input
            type="radio"
            name={id}
            checked={where === "object"}
            disabled={!others.length}
            onChange={() => setWhere("object")}
          />
          <span>
            <u>O</u>bject in:
          </span>
        </label>
        <select
          className="dialog-field"
          aria-label="Object in"
          value={target}
          disabled={where !== "object" || !others.length}
          onChange={(e) => setTarget(Number(e.target.value))}
        >
          {others.map((w) => (
            <option key={w.index} value={w.index}>
              {w.name}
            </option>
          ))}
        </select>
      </div>
    </Dialog>
  )
}
