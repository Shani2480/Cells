import { isChart, type SheetObject } from "@cells/core"
import { ChevronDown, ChevronUp, Eye, EyeOff, X } from "lucide-react"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"

/**
 * Home › Find & Select › Selection Pane (and Arrange › Selection Pane): the sheet's objects, the top one first.
 * A click selects one (Ctrl: one more), a double-click renames it, the eye hides or shows it, the arrows move it
 * up or down the order. Show All and Hide All do every object.
 */
export function SelectionPane({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  const objects = controller.objects
  const [renaming, setRenaming] = useState<number | null>(null)
  const list = [...controller.sheet.objects].reverse()
  const close = () => {
    objects.setPane(false)
    onDone()
  }
  const hidden = (o: SheetObject) => !isChart(o) && !!o.hidden
  const step = (dir: 1 | -1) => {
    const id = objects.selected[0]
    const at = controller.sheet.objects.findIndex((o) => o.id === id)
    if (at >= 0) objects.moveTo(id, at + dir)
  }
  return (
    <aside
      className="comments-pane selection-pane"
      aria-label="Selection"
      onKeyDown={(e) => {
        if (e.key === "Escape" && renaming === null) {
          e.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Selection</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <div className="selection-tools">
        <button type="button" className="dialog-btn" onClick={() => objects.setAllHidden(false)}>
          Show All
        </button>
        <button type="button" className="dialog-btn" onClick={() => objects.setAllHidden(true)}>
          Hide All
        </button>
        <span className="selection-gap" />
        <button
          type="button"
          className="comments-pane-close"
          aria-label="Bring Forward"
          title="Bring Forward"
          disabled={objects.selected.length !== 1}
          onClick={() => step(1)}
        >
          <ChevronUp size={16} strokeWidth={1.75} />
        </button>
        <button
          type="button"
          className="comments-pane-close"
          aria-label="Send Backward"
          title="Send Backward"
          disabled={objects.selected.length !== 1}
          onClick={() => step(-1)}
        >
          <ChevronDown size={16} strokeWidth={1.75} />
        </button>
      </div>
      <ul
        className="comments-pane-list selection-list"
        role="listbox"
        aria-label="Shapes on this Sheet"
        aria-multiselectable
      >
        {list.length ? null : (
          <p className="comments-pane-empty">There are no shapes, pictures or charts on this sheet.</p>
        )}
        {list.map((o) => {
          const on = objects.selected.includes(o.id)
          return (
            <li
              key={o.id}
              role="option"
              aria-selected={on}
              className={`selection-row${on ? " is-selected" : ""}${hidden(o) ? " is-hidden" : ""}`}
              onClick={(e) => {
                if (hidden(o)) return
                objects.select(o.id, e.ctrlKey || e.metaKey)
              }}
              onDoubleClick={() => setRenaming(o.id)}
            >
              {renaming === o.id ? (
                <input
                  className="dialog-field"
                  aria-label="Name"
                  defaultValue={o.name}
                  autoFocus
                  onFocus={(e) => e.currentTarget.select()}
                  onKeyDown={(e) => {
                    e.stopPropagation()
                    if (e.key === "Enter") {
                      objects.rename(o.id, e.currentTarget.value)
                      setRenaming(null)
                    } else if (e.key === "Escape") setRenaming(null)
                  }}
                  onBlur={(e) => {
                    objects.rename(o.id, e.currentTarget.value)
                    setRenaming(null)
                  }}
                />
              ) : (
                <span className="selection-name">{o.name}</span>
              )}
              {isChart(o) ? (
                <span className="selection-eye" />
              ) : (
                <button
                  type="button"
                  className="selection-eye"
                  aria-label={o.hidden ? `Show ${o.name}` : `Hide ${o.name}`}
                  title={o.hidden ? "Show" : "Hide"}
                  onClick={(e) => {
                    e.stopPropagation()
                    objects.setHidden(o.id, !o.hidden)
                  }}
                >
                  {o.hidden ? <EyeOff size={15} strokeWidth={1.75} /> : <Eye size={15} strokeWidth={1.75} />}
                </button>
              )}
            </li>
          )
        })}
      </ul>
    </aside>
  )
}
