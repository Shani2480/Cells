import { Eye, Trash2, X } from "lucide-react"
import { useState, useSyncExternalStore } from "react"
import { absoluteRef } from "@/sheet/names"
import type { SheetController } from "@/sheet/controller"
import { valueText } from "./AuditDialogs"

/**
 * Formulas › Watch Window: cells kept in view, with their sheet, name, value and formula, while you work elsewhere in
 * the workbook. Add Watch adds the selected cells; picking a row goes to its cell. Docked beside the sheet (Excel's is
 * a floating window that docks).
 */
export function WatchWindow({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const [chosen, setChosen] = useState<number | null>(null)
  const watches = controller.audit.watches
  const close = () => {
    controller.audit.toggleWatchWindow(false)
    onDone()
  }
  return (
    <aside
      className="comments-pane watch-window"
      aria-label="Watch Window"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Watch Window</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.6} aria-hidden />
        </button>
      </header>
      <div className="clipboard-pane-actions">
        <button type="button" className="comments-pane-new" onClick={() => controller.audit.addWatches()}>
          <Eye size={16} strokeWidth={1.6} aria-hidden />
          Add Watch
        </button>
        <button
          type="button"
          className="comments-pane-new"
          disabled={chosen === null || !watches[chosen]}
          onClick={() => {
            if (chosen !== null) controller.audit.deleteWatch(chosen)
            setChosen(null)
          }}
        >
          <Trash2 size={16} strokeWidth={1.6} aria-hidden />
          Delete Watch
        </button>
      </div>
      <div className="comments-pane-list" role="list">
        {!watches.length ? (
          <p className="comments-pane-empty">Select cells and choose Add Watch to keep an eye on their values here.</p>
        ) : null}
        {watches.map((w, i) => {
          const index = controller.wb.sheets.indexOf(w.sheet)
          if (index < 0) return null
          const cell = absoluteRef({ r0: w.row, c0: w.col, r1: w.row, c1: w.col }).replace(/\$/g, "")
          const name = controller.names.all().find((n) => {
            const t = controller.names.resolve(n.name)
            return (
              t &&
              t.sheet === index &&
              t.range.r0 === w.row &&
              t.range.r1 === w.row &&
              t.range.c0 === w.col &&
              t.range.c1 === w.col
            )
          })?.name
          const input = w.sheet.cells.input(w.row, w.col)
          return (
            <button
              key={`${index}:${w.row}:${w.col}`}
              type="button"
              role="listitem"
              className={`comments-pane-thread watch-row${i === chosen ? " is-selected" : ""}`}
              aria-label={`${w.sheet.name} ${cell}`}
              onClick={() => {
                setChosen(i)
                if (index !== controller.sheetIndex) controller.activate(index)
                controller.goTo({ r0: w.row, c0: w.col, r1: w.row, c1: w.col })
              }}
            >
              <span className="comments-pane-cell">
                {w.sheet.name}!{cell}
                {name ? <span className="watch-name"> · {name}</span> : null}
              </span>
              <span className="watch-value">{valueText(w.sheet.cells.value(w.row, w.col))}</span>
              {input?.startsWith("=") ? <code className="watch-formula">{input}</code> : null}
            </button>
          )
        })}
      </div>
    </aside>
  )
}
