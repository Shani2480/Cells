import { useMemo, useState } from "react"
import { vbaMacros } from "@cells/io"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * View › Macros: the workbook's VBA macros, as Excel lists them. Cells can't run VBA, so Run stays off with the
 * reason shown; Edit shows the macro's module (read-only). The macros stay in the file when it's saved as .xlsm.
 */
export function MacrosDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const modules = controller.wb.vba
  const macros = useMemo(() => vbaMacros(modules), [modules])
  const [chosen, setChosen] = useState(macros[0]?.name ?? "")
  const [viewing, setViewing] = useState(false)
  const current = macros.find((m) => m.name === chosen) ?? macros[0]
  const code = modules.find((m) => m.name === current?.module)?.code ?? ""
  if (viewing && current)
    return (
      <Dialog
        title={`${current.module} (Code)`}
        onCancel={() => setViewing(false)}
        onAccept={() => setViewing(false)}
        width={620}
      >
        <pre className="macro-code" aria-label="Macro code" tabIndex={0}>
          {code}
        </pre>
      </Dialog>
    )
  return (
    <Dialog title="Macro" onCancel={onClose} onAccept={onClose} width={440} actions={<></>}>
      <div className="custom-views">
        <div>
          <label className="dialog-label" htmlFor="macro-name">
            <u>M</u>acro name:
          </label>
          <input id="macro-name" className="dialog-field" readOnly value={current?.name ?? ""} />
          <div className="dialog-list custom-views-list" role="listbox" aria-label="Macros">
            {macros.map((m) => (
              <button
                type="button"
                role="option"
                aria-selected={m === current}
                key={m.name}
                className={`dialog-list-item${m === current ? " is-selected" : ""}`}
                onClick={() => setChosen(m.name)}
                onDoubleClick={() => setViewing(true)}
              >
                {m.name}
              </button>
            ))}
          </div>
          <p className="dialog-note">
            {macros.length
              ? "Cells shows macros but doesn't run them. They stay in the workbook when you save it as .xlsm."
              : "This workbook has no macros."}
          </p>
        </div>
        <div className="custom-views-buttons">
          <button type="button" className="dialog-btn" disabled title="Cells doesn't run VBA macros">
            <u>R</u>un
          </button>
          <button type="button" className="dialog-btn" disabled={!current} onClick={() => setViewing(true)}>
            <u>E</u>dit
          </button>
          <button type="button" className="dialog-btn is-primary" onClick={onClose}>
            Close
          </button>
        </div>
      </div>
    </Dialog>
  )
}
