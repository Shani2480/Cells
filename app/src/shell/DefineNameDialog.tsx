import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * New Name (Formulas › Define Name): the name, its scope, a comment and what it refers to (the selection to start
 * with). Names are workbook-wide for now; sheet scope comes with the Name Manager.
 */
export function DefineNameDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [name, setName] = useState(() => suggestName(controller))
  const [comment, setComment] = useState("")
  const [refersTo, setRefersTo] = useState(() => "=" + controller.names.selectionRef())
  const accept = () => {
    const problem = controller.names.define(name, refersTo, comment)
    if (problem) {
      controller.shell.alert(problem)
      return
    }
    onClose()
  }
  return (
    <Dialog title="New Name" onCancel={onClose} onAccept={accept} width={380}>
      <div className="define-name-grid">
        <label className="dialog-label" htmlFor="name-name">
          <u>N</u>ame:
        </label>
        <input id="name-name" className="dialog-field" value={name} onChange={(e) => setName(e.target.value)} />
        <label className="dialog-label" htmlFor="name-scope">
          <u>S</u>cope:
        </label>
        <select id="name-scope" className="dialog-field" disabled title="Sheet scope arrives with the Name Manager.">
          <option>Workbook</option>
        </select>
        <label className="dialog-label" htmlFor="name-comment">
          C<u>o</u>mment:
        </label>
        <textarea
          id="name-comment"
          className="dialog-field define-name-comment"
          value={comment}
          onChange={(e) => setComment(e.target.value)}
        />
        <label className="dialog-label" htmlFor="name-refers">
          <u>R</u>efers to:
        </label>
        <input
          id="name-refers"
          className="dialog-field"
          value={refersTo}
          onChange={(e) => setRefersTo(e.target.value)}
        />
      </div>
    </Dialog>
  )
}

/** Excel suggests the text of the cell above or left of a single selected cell when it makes a valid name. */
function suggestName(c: SheetController): string {
  const r = c.selection.ranges[0]
  const sheet = c.sheet
  const candidates = [
    r.r0 > 0 ? sheet.cells.value(r.r0 - 1, r.c0) : undefined,
    r.c0 > 0 ? sheet.cells.value(r.r0, r.c0 - 1) : undefined,
    sheet.cells.value(r.r0, r.c0),
  ]
  for (const v of candidates) {
    if (typeof v !== "string") continue
    const name = v.trim().replace(/\s+/g, "_")
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !c.names.resolve(name)) return name
  }
  return ""
}
