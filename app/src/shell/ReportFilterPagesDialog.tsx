import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** Show Report Filter Pages: which report filter gets a sheet for each of its items. */
export function ReportFilterPagesDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const filters = controller.pivots.reportFilters()
  const [chosen, setChosen] = useState(0)
  if (!filters.length) return null
  const accept = () => {
    controller.pivots.showReportFilterPages(filters[chosen].field)
    onClose()
  }
  return (
    <Dialog title="Show Report Filter Pages" width={300} onCancel={onClose} onAccept={accept}>
      <span className="dialog-label" id="report-pages-label">
        <u>S</u>how all report filter pages of:
      </span>
      <div
        className="dialog-list report-pages-list"
        role="listbox"
        aria-labelledby="report-pages-label"
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            // Enter on the list is OK, as in Excel (not a click on the focused item).
            e.preventDefault()
            e.stopPropagation()
            accept()
            return
          }
          const step = e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0
          if (!step) return
          e.preventDefault()
          setChosen((i) => Math.min(filters.length - 1, Math.max(0, i + step)))
        }}
      >
        {filters.map((f, i) => (
          <button
            key={f.field}
            type="button"
            role="option"
            aria-selected={i === chosen}
            data-autofocus={i === chosen ? true : undefined}
            className={`dialog-list-item${i === chosen ? " is-selected" : ""}`}
            onClick={() => setChosen(i)}
            onDoubleClick={() => {
              setChosen(i)
              controller.pivots.showReportFilterPages(f.field)
              onClose()
            }}
          >
            {f.name}
          </button>
        ))}
      </div>
    </Dialog>
  )
}
