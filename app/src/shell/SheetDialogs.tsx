import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * Move or Copy (Excel's dialog): where the active sheet goes ("Before sheet", or "(move to end)") and whether it is a
 * copy. Excel's "To book" box lists open workbooks; with one workbook open it shows only this one.
 */
export function MoveCopyDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const sheets = controller.wb.sheets
  const index = controller.sheetIndex
  // Excel selects the sheet after the active one (the active sheet's own place).
  const [before, setBefore] = useState(Math.min(index + 1, sheets.length))
  const [copy, setCopy] = useState(false)
  const accept = () => {
    if (copy) controller.sheets.copy(index, before)
    else controller.sheets.move(index, before)
    onClose()
  }
  const items = [
    ...sheets.map((s, i) => ({ i, name: s.name, hidden: s.hidden })),
    { i: sheets.length, name: "(move to end)", hidden: false },
  ]
  return (
    <Dialog title="Move or Copy" onCancel={onClose} onAccept={accept} width={320}>
      <span className="dialog-label">Move selected sheets</span>
      <label className="dialog-label" htmlFor="movecopy-book">
        <u>T</u>o book:
      </label>
      <select id="movecopy-book" className="dialog-field" disabled>
        <option>Book1</option>
      </select>
      <span className="dialog-label">
        <u>B</u>efore sheet:
      </span>
      <div className="dialog-list sheet-list" role="listbox" aria-label="Before sheet">
        {items
          .filter((item) => !item.hidden)
          .map((item) => (
            <button
              type="button"
              role="option"
              aria-selected={item.i === before}
              key={item.i}
              className={`dialog-list-item${item.i === before ? " is-selected" : ""}`}
              onClick={() => setBefore(item.i)}
              onDoubleClick={() => {
                setBefore(item.i)
                if (copy) controller.sheets.copy(index, item.i)
                else controller.sheets.move(index, item.i)
                onClose()
              }}
            >
              {item.name}
            </button>
          ))}
      </div>
      <label className="fc-check">
        <input type="checkbox" checked={copy} onChange={(e) => setCopy(e.target.checked)} />
        <span>
          <u>C</u>reate a copy
        </span>
      </label>
    </Dialog>
  )
}

/** Unhide: the hidden sheets; the chosen one is shown again and becomes active. */
export function UnhideSheetDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const hidden = controller.sheets.hidden()
  const [chosen, setChosen] = useState(hidden[0] ?? -1)
  const accept = () => {
    if (chosen >= 0) controller.sheets.unhide(chosen)
    onClose()
  }
  return (
    <Dialog title="Unhide" onCancel={onClose} onAccept={accept} width={300}>
      <span className="dialog-label">
        <u>U</u>nhide sheet:
      </span>
      <div className="dialog-list sheet-list" role="listbox" aria-label="Unhide sheet">
        {hidden.map((i) => (
          <button
            type="button"
            role="option"
            aria-selected={i === chosen}
            key={i}
            className={`dialog-list-item${i === chosen ? " is-selected" : ""}`}
            onClick={() => setChosen(i)}
            onDoubleClick={() => {
              controller.sheets.unhide(i)
              onClose()
            }}
          >
            {controller.wb.sheets[i].name}
          </button>
        ))}
      </div>
    </Dialog>
  )
}
