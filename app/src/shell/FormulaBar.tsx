import { parseRange } from "@cells/core"
import { Check, X } from "lucide-react"
import { useLayoutEffect, useRef, useState, useSyncExternalStore } from "react"
import { FormulaAssist, FormulaHighlight } from "@/grid/FormulaAssist"
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

  const [scrollLeft, setScrollLeft] = useState(0)
  useLayoutEffect(() => {
    const bar = barRef.current
    if (!bar) return
    if (bar.value !== text) bar.value = text
    // Typing here, or pointing at cells from here: keep the caret where the controller put it.
    if (editing?.surface === "bar" && document.activeElement === bar) {
      if (bar.selectionStart !== editing.selStart || bar.selectionEnd !== editing.selEnd)
        bar.setSelectionRange(editing.selStart, editing.selEnd)
    }
    setScrollLeft(bar.scrollLeft)
  }, [text, editing])

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
      <div className="formula-input-wrap">
        <FormulaHighlight controller={controller} className="formula-input-highlight" scrollLeft={scrollLeft} />
        <textarea
          ref={barRef}
          className={
            "formula-input" +
            (editing?.text.startsWith("=") ? " has-highlight" : "") +
            (controller.formulaBarGhost ? " is-ghost" : "")
          }
          aria-label="Formula Bar"
          rows={1}
          wrap="off"
          spellCheck={false}
          defaultValue={text}
          onFocus={() => {
            // A spilled cell has no formula of its own to edit (its anchor has).
            if (controller.formulaBarGhost) return
            if (controller.edit) controller.setEditSurface("bar")
            else controller.beginEdit("edit", undefined, "bar")
          }}
          onInput={(event) => {
            const { value, selectionStart, selectionEnd } = event.currentTarget
            if (!controller.edit) controller.beginEdit("edit", value, "bar")
            else controller.setEditText(value, selectionStart, selectionEnd)
          }}
          onSelect={(event) => {
            // Only when the textarea shows the entry's current text (otherwise the offsets are stale).
            const edit = controller.edit
            if (edit?.surface === "bar" && event.currentTarget.value === edit.text)
              controller.setEditSelection(event.currentTarget.selectionStart, event.currentTarget.selectionEnd)
            setScrollLeft(event.currentTarget.scrollLeft)
          }}
          onScroll={(event) => setScrollLeft(event.currentTarget.scrollLeft)}
          onKeyDown={(event) => {
            const input = keyInputFrom(event, platform.os)
            if (controller.edit && controller.keyEditing(input)) {
              event.preventDefault()
              // Keys that keep the entry open (F4, pointing, the function list) keep the formula bar focused.
              if (!controller.edit) onFocusGrid()
            }
          }}
          readOnly={controller.formulaBarGhost}
        />
        {editing?.surface === "bar" ? (
          <FormulaAssist controller={controller} style={{ left: 0, top: 26 }} onPick={() => barRef.current?.focus()} />
        ) : null}
      </div>
    </div>
  )
}
