import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { parseSheetRange } from "./AdvancedFilterDialog"
import { Dialog } from "./Dialog"

/**
 * Insert › PivotTable (and Table Design › Summarize with PivotTable): Excel 365's "PivotTable from table or range".
 * The table or range to summarize, and a new worksheet or a cell of an existing one. "Add this data to the Data
 * Model" needs Microsoft's engine and isn't offered.
 */
export function CreatePivotDialog({
  controller,
  source: initial,
  onClose,
}: {
  controller: SheetController
  /** The Table/Range box's starting text (Summarize with PivotTable gives its table's name). */
  source?: string
  onClose: () => void
}) {
  const [source, setSource] = useState(() => initial ?? controller.pivots.defaultSource())
  const [existing, setExisting] = useState(false)
  const [location, setLocation] = useState("")

  const accept = () => {
    let problem: string | null
    if (existing) {
      const at = parseSheetRange(controller, location)
      if (!at) return controller.shell.alert("Reference isn't valid.")
      problem = controller.pivots.create(source, {
        kind: "existing",
        sheet: at.sheet,
        at: { row: at.range.r0, col: at.range.c0 },
      })
    } else problem = controller.pivots.create(source, { kind: "newSheet" })
    if (problem) return controller.shell.alert(problem)
    onClose()
  }

  return (
    <Dialog title="PivotTable from table or range" onCancel={onClose} onAccept={accept} width={400}>
      <p className="dialog-label">Select a table or range</p>
      <div className="goal-seek-fields">
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
      <fieldset className="dialog-group">
        <legend>Choose where you want the PivotTable to be placed</legend>
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
