import { columnName, displayText, type Range } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * Remove Duplicates: the columns to compare (each a check box, named by its header when the data has headers),
 * Select All, Unselect All and My data has headers.
 */
export function RemoveDuplicatesDialog({
  controller,
  range,
  onClose,
}: {
  controller: SheetController
  range: Range
  onClose: () => void
}) {
  const cols = Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i)
  const [header, setHeader] = useState(() => controller.data.target(range).header)
  const [chosen, setChosen] = useState<Set<number>>(() => new Set(cols))
  const sheet = controller.sheet
  const nameOf = (col: number) => {
    if (!header) return `Column ${columnName(col)}`
    const code = controller.wb.styles.get(sheet.styleAt(range.r0, col)).numberFormat
    return displayText(sheet.cells.get(range.r0, col)?.value, code) || `Column ${columnName(col)}`
  }
  const accept = () => {
    if (!chosen.size) return
    onClose()
    controller.data.removeDuplicates(
      range,
      cols.filter((c) => chosen.has(c)),
      header,
    )
  }
  return (
    <Dialog title="Remove Duplicates" width={420} onCancel={onClose} onAccept={accept}>
      <p className="dialog-message">To delete duplicate values, select one or more columns that contain duplicates.</p>
      <div className="rd-bar">
        <button type="button" className="dialog-btn" onClick={() => setChosen(new Set(cols))}>
          Select All
        </button>
        <button type="button" className="dialog-btn" onClick={() => setChosen(new Set())}>
          Unselect All
        </button>
        <label className="dialog-radio rd-header">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          <span>
            My data has <u>h</u>eaders
          </span>
        </label>
      </div>
      <div className="rd-list" role="group" aria-label="Columns">
        <div className="rd-list-head">Columns</div>
        {cols.map((col) => (
          <label key={col} className="dialog-radio rd-item">
            <input
              type="checkbox"
              checked={chosen.has(col)}
              onChange={(e) => {
                const next = new Set(chosen)
                if (e.target.checked) next.add(col)
                else next.delete(col)
                setChosen(next)
              }}
            />
            {nameOf(col)}
          </label>
        ))}
      </div>
    </Dialog>
  )
}
