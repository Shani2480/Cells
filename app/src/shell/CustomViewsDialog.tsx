import { useState, useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"
import { VIEW_MESSAGES } from "@/sheet/custom-views"
import { Dialog } from "./Dialog"

/** View › Custom Views: the views, with Show, Close, Add… and Delete (Excel's dialog). */
export function CustomViewsDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const views = controller.customViews.list()
  const [chosen, setChosen] = useState(views[0]?.name ?? "")
  const [adding, setAdding] = useState(false)
  const current = views.find((v) => v.name === chosen) ?? views[0]
  const show = (name: string) => {
    onClose()
    controller.customViews.show(name)
  }
  if (adding) return <AddViewDialog controller={controller} onClose={onClose} />
  return (
    <Dialog
      title="Custom Views"
      onCancel={onClose}
      onAccept={() => current && show(current.name)}
      width={380}
      // Excel's buttons are beside the list, with no row of OK and Cancel.
      actions={<></>}
    >
      <div className="custom-views">
        <div>
          <span className="dialog-label">
            Vie<u>w</u>s:
          </span>
          <div className="dialog-list custom-views-list" role="listbox" aria-label="Views">
            {views.map((v) => (
              <button
                type="button"
                role="option"
                aria-selected={v === current}
                key={v.guid}
                className={`dialog-list-item${v === current ? " is-selected" : ""}`}
                onClick={() => setChosen(v.name)}
                onDoubleClick={() => show(v.name)}
              >
                {v.name}
              </button>
            ))}
          </div>
        </div>
        <div className="custom-views-buttons">
          <button
            type="button"
            className="dialog-btn is-primary"
            disabled={!current}
            onClick={() => current && show(current.name)}
          >
            <u>S</u>how
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Close
          </button>
          <button type="button" className="dialog-btn" onClick={() => setAdding(true)}>
            <u>A</u>dd...
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={!current}
            onClick={() =>
              current &&
              controller.shell.confirm(VIEW_MESSAGES.remove(current.name), () => {
                controller.customViews.remove(current.name)
                setChosen(controller.customViews.list()[0]?.name ?? "")
              })
            }
          >
            <u>D</u>elete
          </button>
        </div>
      </div>
    </Dialog>
  )
}

/** Add View: a name, and whether the view keeps the print settings and the hidden rows, columns and filters. */
export function AddViewDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [name, setName] = useState("")
  const [print, setPrint] = useState(true)
  const [hidden, setHidden] = useState(true)
  const accept = () => {
    const n = name.trim()
    if (!n) return
    const add = () => {
      controller.customViews.add(n, print, hidden)
      onClose()
    }
    if (controller.customViews.find(n)) controller.shell.confirm(VIEW_MESSAGES.replace(n), add)
    else add()
  }
  return (
    <Dialog title="Add View" onCancel={onClose} onAccept={accept} width={340}>
      <label className="dialog-label" htmlFor="view-name">
        <u>N</u>ame:
      </label>
      <input id="view-name" className="dialog-field" autoFocus value={name} onChange={(e) => setName(e.target.value)} />
      <fieldset className="zoom-options">
        <legend className="dialog-label">Include in view</legend>
        <label className="fc-check">
          <input type="checkbox" checked={print} onChange={(e) => setPrint(e.target.checked)} />
          <span>
            <u>P</u>rint settings
          </span>
        </label>
        <label className="fc-check">
          <input type="checkbox" checked={hidden} onChange={(e) => setHidden(e.target.checked)} />
          <span>
            Hidden <u>r</u>ows, columns and filter settings
          </span>
        </label>
      </fieldset>
    </Dialog>
  )
}
