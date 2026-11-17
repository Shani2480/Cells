import { useState } from "react"
import type { Arrangement, DocumentWindows } from "@/file/windows"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

const ARRANGEMENTS: { kind: Arrangement; label: [string, string, string] }[] = [
  { kind: "tiled", label: ["", "T", "iled"] },
  { kind: "horizontal", label: ["Hori", "z", "ontal"] },
  { kind: "vertical", label: ["", "V", "ertical"] },
  { kind: "cascade", label: ["", "C", "ascade"] },
]

/** View › Arrange All (Excel's Arrange Windows). */
export function ArrangeWindowsDialog({ windows, onClose }: { windows: DocumentWindows; onClose: () => void }) {
  const [kind, setKind] = useState<Arrangement>(windows.arrangement ?? "tiled")
  const accept = () => {
    windows.arrange(kind)
    onClose()
  }
  return (
    <Dialog title="Arrange Windows" onCancel={onClose} onAccept={accept} width={260}>
      <fieldset className="dialog-group">
        <legend>Arrange</legend>
        {ARRANGEMENTS.map(({ kind: k, label: [before, key, after] }) => (
          <label className="dialog-radio" key={k}>
            <input type="radio" name="arrange-windows" checked={kind === k} onChange={() => setKind(k)} />
            <span>
              {before}
              <u>{key}</u>
              {after}
            </span>
          </label>
        ))}
      </fieldset>
    </Dialog>
  )
}

function WindowList({
  label,
  list,
  title,
  chosen,
  onChoose,
  onPick,
}: {
  label: string
  list: SheetController[]
  title: (c: SheetController) => string
  chosen: SheetController | undefined
  onChoose: (c: SheetController) => void
  onPick: (c: SheetController) => void
}) {
  return (
    <div className="dialog-list sheet-list" role="listbox" aria-label={label}>
      {list.map((c) => (
        <button
          type="button"
          role="option"
          aria-selected={c === chosen}
          key={title(c)}
          className={`dialog-list-item${c === chosen ? " is-selected" : ""}`}
          onClick={() => onChoose(c)}
          onDoubleClick={() => onPick(c)}
        >
          {title(c)}
        </button>
      ))}
    </div>
  )
}

/** View › Unhide: the workbook's hidden windows. */
export function UnhideWindowDialog({
  windows,
  baseTitle,
  onClose,
}: {
  windows: DocumentWindows
  baseTitle: string
  onClose: () => void
}) {
  const hidden = windows.hidden
  const [chosen, setChosen] = useState(hidden[0])
  const pick = (c: SheetController | undefined) => {
    if (c) windows.unhide(c)
    onClose()
  }
  return (
    <Dialog title="Unhide" onCancel={onClose} onAccept={() => pick(chosen)} width={300}>
      <span className="dialog-label">
        <u>U</u>nhide workbook:
      </span>
      <WindowList
        label="Unhide workbook"
        list={hidden}
        title={(c) => windows.title(baseTitle, c)}
        chosen={chosen}
        onChoose={setChosen}
        onPick={pick}
      />
    </Dialog>
  )
}

/** View Side by Side with more than two windows: which one to compare the active window with. */
export function CompareSideBySideDialog({
  windows,
  baseTitle,
  onClose,
}: {
  windows: DocumentWindows
  baseTitle: string
  onClose: () => void
}) {
  const others = windows.all.filter((c) => c !== windows.active)
  const [chosen, setChosen] = useState(others[0])
  const pick = (c: SheetController | undefined) => {
    onClose()
    if (c) windows.startSideBySide(c)
  }
  return (
    <Dialog title="Compare Side by Side" onCancel={onClose} onAccept={() => pick(chosen)} width={320}>
      <span className="dialog-label">
        Compare <strong>{windows.title(baseTitle)}</strong> side by side with:
      </span>
      <WindowList
        label="Compare side by side with"
        list={others}
        title={(c) => windows.title(baseTitle, c)}
        chosen={chosen}
        onChoose={setChosen}
        onPick={pick}
      />
    </Dialog>
  )
}
