import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import type { PivotDestination } from "@/sheet/pivots"
import { parseSheetRange } from "./AdvancedFilterDialog"
import { Dialog } from "./Dialog"

const TITLES = {
  create: "PivotTable from table or range",
  source: "Change PivotTable Data Source",
  move: "Move PivotTable",
  chart: "Create PivotChart",
}

/**
 * Insert › PivotTable (and Table Design › Summarize with PivotTable): Excel 365's "PivotTable from table or range".
 * The table or range to summarize, and a new worksheet or a cell of an existing one. "Add this data to the Data
 * Model" needs Microsoft's engine and isn't offered.
 */
export function CreatePivotDialog({
  controller,
  source: initial,
  mode = "create",
  chartType = "column",
  onClose,
}: {
  controller: SheetController
  /** The Table/Range box's starting text (Summarize with PivotTable gives its table's name). */
  source?: string
  /**
   * "source": Analyze › Change Data Source (the Table/Range box only); "move": Analyze › Move PivotTable (where to
   * only).
   */
  mode?: "create" | "source" | "move" | "chart"
  /** "chart" (Insert › PivotChart outside a PivotTable): the chart type picked. */
  chartType?: string
  onClose: () => void
}) {
  const [source, setSource] = useState(
    () => initial ?? (mode === "source" ? controller.pivots.sourceText() : controller.pivots.defaultSource()),
  )
  const [existing, setExisting] = useState(false)
  const [location, setLocation] = useState("")

  const accept = () => {
    if (mode === "source") {
      const problem = controller.pivots.changeSource(source)
      if (problem) return controller.shell.alert(problem)
      return onClose()
    }
    let dest: PivotDestination = { kind: "newSheet" }
    if (existing) {
      const at = parseSheetRange(controller, location)
      if (!at) return controller.shell.alert("Reference isn't valid.")
      dest = { kind: "existing", sheet: at.sheet, at: { row: at.range.r0, col: at.range.c0 } }
    }
    const problem =
      mode === "move"
        ? controller.pivots.move(dest)
        : mode === "chart"
          ? controller.pivots.createWithChart(source, dest, chartType)
          : controller.pivots.create(source, dest)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }

  return (
    <Dialog title={TITLES[mode]} onCancel={onClose} onAccept={accept} width={400}>
      {mode === "move" ? null : (
        <p className="dialog-label">
          {mode === "create" ? "Select a table or range" : "Choose the data that you want to analyze"}
        </p>
      )}
      <div className="goal-seek-fields" hidden={mode === "move"}>
        <label className="dialog-label" htmlFor="pivot-source">
          Table/Range:
        </label>
        <input
          id="pivot-source"
          className="dialog-field"
          autoFocus
          value={source}
          onChange={(e) => setSource(e.target.value)}
        />
      </div>
      <fieldset className="dialog-group" hidden={mode === "source"}>
        <legend>Choose where you want the {mode === "chart" ? "PivotChart" : "PivotTable"} to be placed</legend>
        <label className="dialog-radio">
          <input type="radio" name="pivot-place" checked={!existing} onChange={() => setExisting(false)} />
          <span>New Worksheet</span>
        </label>
        <label className="dialog-radio">
          <input type="radio" name="pivot-place" checked={existing} onChange={() => setExisting(true)} />
          <span>Existing Worksheet</span>
        </label>
        <div className="goal-seek-fields">
          <label className="dialog-label" htmlFor="pivot-location">
            Location:
          </label>
          <input
            id="pivot-location"
            className="dialog-field"
            disabled={!existing}
            value={location}
            onChange={(e) => setLocation(e.target.value)}
          />
        </div>
      </fieldset>
    </Dialog>
  )
}
