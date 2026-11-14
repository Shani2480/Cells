import { Check, ChevronDown, ChevronUp, X } from "lucide-react"
import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from "react"
import { FormulaAssist, FormulaHighlight } from "@/grid/FormulaAssist"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"
import { keyInputFrom } from "@/sheet/keys"

/**
 * The Name Box (shows the active cell or the selection's name; type a reference or name and Enter to go there, or a
 * new name to name the selection; its arrow lists the names), Cancel/Enter while an entry is open, Insert Function
 * (fx), and the formula bar, which edits the same entry as the cell. It expands to several lines (the chevron or
 * Ctrl+Shift+U) and its bottom edge sets how many.
 */
/** The formula bar's line height (its font is the cells' 11 pt). */
const LINE = 17

export function FormulaBar({
  controller,
  dragging,
  onFocusGrid,
  onInsertFunction,
}: {
  controller: SheetController
  dragging: boolean
  onFocusGrid: () => void
  onInsertFunction: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const platform = usePlatform()
  const [nameDraft, setNameDraft] = useState<string | null>(null)
  const barRef = useRef<HTMLTextAreaElement>(null)
  const editing = controller.edit
  const text = controller.formulaBarText()

  const [scrollLeft, setScrollLeft] = useState(0)
  const [scrollTop, setScrollTop] = useState(0)
  const [namesOpen, setNamesOpen] = useState(false)
  const expanded = controller.formulaBarExpanded
  const names = controller.names.all()
  const nameWrapRef = useRef<HTMLDivElement>(null)
  // The names list closes on a click anywhere else, or Esc.
  useEffect(() => {
    if (!namesOpen) return
    const away = (event: PointerEvent) => {
      if (!nameWrapRef.current?.contains(event.target as Node)) setNamesOpen(false)
    }
    const esc = (event: KeyboardEvent) => event.key === "Escape" && setNamesOpen(false)
    window.addEventListener("pointerdown", away, true)
    window.addEventListener("keydown", esc, true)
    return () => {
      window.removeEventListener("pointerdown", away, true)
      window.removeEventListener("keydown", esc, true)
    }
  }, [namesOpen])
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
    setScrollTop(bar.scrollTop)
  }, [text, editing])

  const goToName = () => {
    const draft = nameDraft
    setNameDraft(null)
    if (draft) controller.names.fromNameBox(draft)
    onFocusGrid()
  }

  // Dragging the bottom edge of an expanded bar: whole lines, 17 px each.
  const startResize = (event: React.PointerEvent<HTMLDivElement>) => {
    event.preventDefault()
    const top = barRef.current?.getBoundingClientRect().top ?? event.clientY
    const target = event.currentTarget
    target.setPointerCapture?.(event.pointerId)
    const move = (e: PointerEvent) => controller.setFormulaBarLines((e.clientY - top - 6) / LINE)
    const up = () => {
      target.removeEventListener("pointermove", move)
      target.removeEventListener("pointerup", up)
      target.removeEventListener("pointercancel", up)
    }
    target.addEventListener("pointermove", move)
    target.addEventListener("pointerup", up)
    target.addEventListener("pointercancel", up)
  }

  return (
    <div className={`formula-bar${expanded ? " is-expanded" : ""}`}>
      <div className="name-box-wrap" ref={nameWrapRef}>
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
        <button
          type="button"
          className="name-box-arrow"
          aria-label="Defined names"
          aria-haspopup="listbox"
          aria-expanded={namesOpen}
          disabled={!names.length}
          onPointerDown={(event) => event.preventDefault()}
          onClick={() => setNamesOpen((open) => !open)}
        >
          <ChevronDown size={12} strokeWidth={1.75} />
        </button>
        {namesOpen ? (
          <div className="name-box-list" role="listbox" aria-label="Defined names">
            {names.map((n) => (
              <button
                type="button"
                role="option"
                aria-selected={false}
                key={n.name}
                onPointerDown={(event) => event.preventDefault()}
                onClick={() => {
                  setNamesOpen(false)
                  const target = controller.names.resolve(n.name)
                  if (target) controller.names.goTo(target)
                  onFocusGrid()
                }}
              >
                {n.name}
              </button>
            ))}
          </div>
        ) : null}
      </div>
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
        <button
          type="button"
          className="formula-bar-button formula-bar-fx"
          aria-label="Insert Function"
          title="Insert Function (Shift+F3)"
          onPointerDown={(event) => event.preventDefault()}
          onClick={onInsertFunction}
        >
          fx
        </button>
      </div>
      <div className="formula-input-wrap">
        <FormulaHighlight
          controller={controller}
          className="formula-input-highlight"
          scrollLeft={scrollLeft}
          scrollTop={scrollTop}
        />
        <textarea
          ref={barRef}
          className={
            "formula-input" +
            (editing?.text.startsWith("=") ? " has-highlight" : "") +
            (controller.formulaBarGhost ? " is-ghost" : "")
          }
          aria-label="Formula Bar"
          rows={expanded ? controller.formulaBarLines : 1}
          wrap={expanded ? "soft" : "off"}
          style={expanded ? { height: controller.formulaBarLines * LINE + 6 } : undefined}
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
            setScrollTop(event.currentTarget.scrollTop)
          }}
          onScroll={(event) => {
            setScrollLeft(event.currentTarget.scrollLeft)
            setScrollTop(event.currentTarget.scrollTop)
          }}
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
          <FormulaAssist
            controller={controller}
            style={{ left: 0, top: (expanded ? controller.formulaBarLines * LINE + 6 : 22) + 4 }}
            onPick={() => barRef.current?.focus()}
          />
        ) : null}
      </div>
      <button
        type="button"
        className="formula-bar-button formula-bar-expand"
        aria-label={expanded ? "Collapse Formula Bar" : "Expand Formula Bar"}
        title={`${expanded ? "Collapse" : "Expand"} Formula Bar (Ctrl+Shift+U)`}
        onPointerDown={(event) => event.preventDefault()}
        onClick={() => controller.toggleFormulaBarExpanded()}
      >
        {expanded ? <ChevronUp size={14} strokeWidth={1.75} /> : <ChevronDown size={14} strokeWidth={1.75} />}
      </button>
      {expanded ? (
        <div
          className="formula-bar-resize"
          role="separator"
          aria-orientation="horizontal"
          aria-label="Formula bar height"
          onPointerDown={startResize}
        />
      ) : null}
    </div>
  )
}
