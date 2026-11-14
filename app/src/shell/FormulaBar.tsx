import { parseRange } from "@cells/core"
import { Check, X } from "lucide-react"
import { useEffect, useRef, useState, useSyncExternalStore } from "react"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"
import { keyInputFrom } from "@/sheet/keys"

/**
 * The Name Box (shows the active cell; type a reference and Enter to go there), Cancel/Enter buttons while an entry
 * is open, and the formula bar, which edits the same entry as the cell. Insert Function (fx) arrives with the
 * function library in Phase 2.
 */
export function FormulaBar({
  controller,
  dragging,
  onFocusGrid,
}: {
  controller: SheetController
  dragging: boolean
  onFocusGrid: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const platform = usePlatform()
  const [nameDraft, setNameDraft] = useState<string | null>(null)
  const barRef = useRef<HTMLTextAreaElement>(null)
  const editing = controller.edit
  const text = controller.formulaBarText()

  useEffect(() => {
    const bar = barRef.current
    if (bar && bar.value !== text) bar.value = text
  }, [text])

  const goToName = () => {
    const range = nameDraft ? parseRange(nameDraft) : null
    setNameDraft(null)
    if (range) controller.goTo(range)
    onFocusGrid()
  }

  return (
    <div className="formula-bar">
      <input
        className="name-box"
        aria-label="Name Box"
        value={nameDraft ?? controller.nameBoxText(dragging)}
        onFocus={(event) => {
          setNameDraft(controller.nameBoxText())
          requestAnimationFrame(() => event.target.select())
        }}
        onChange={(event) => setNameDraft(event.target.value)}
        onBlur={() => setNameDraft(null)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault()
            goToName()
          } else if (event.key === "Escape") {
            event.preventDefault()
            setNameDraft(null)
            onFocusGrid()
          }
        }}
      />
      <div className="formula-bar-buttons">
        <button
          type="button"
          className="formula-bar-button"
          aria-label="Cancel"
          title="Cancel"
          disabled={!editing}
          onPointerDown={(event) => event.preventDefault()}
          onClick={() => {
            controller.cancelEdit()
            onFocusGrid()
          }}
        >
          <X size={15} strokeWidth={1.75} />
        </button>
        <button
          type="button"
          className="formula-bar-button"
          aria-label="Enter"
          title="Enter"
          disabled={!editing}
          onPointerDown={(event) => event.preventDefault()}
          onClick={() => {
            controller.commitEdit()
            onFocusGrid()
          }}
        >
          <Check size={15} strokeWidth={1.75} />
        </button>
        <span className="formula-bar-fx" aria-hidden>
          fx
        </span>
      </div>
      <textarea
        ref={barRef}
        className="formula-input"
        aria-label="Formula Bar"
        rows={1}
        wrap="off"
        spellCheck={false}
        defaultValue={text}
        onFocus={() => {
          if (controller.edit) controller.setEditSurface("bar")
          else controller.beginEdit("edit", undefined, "bar")
        }}
        onInput={(event) => {
          if (!controller.edit) controller.beginEdit("edit", event.currentTarget.value, "bar")
          else controller.setEditText(event.currentTarget.value)
        }}
        onKeyDown={(event) => {
          const input = keyInputFrom(event, platform.os)
          if (controller.edit && controller.keyEditing(input)) {
            event.preventDefault()
            onFocusGrid()
          }
        }}
      />
    </div>
  )
}
