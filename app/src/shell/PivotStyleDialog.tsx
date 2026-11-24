import { useState } from "react"
import {
  duplicateStyleName,
  elementSummary,
  isEmptyElement,
  newCustomStyleName,
  PIVOT_STYLE_ELEMENTS,
  pivotStyleDef,
  type CustomTableStyle,
  type PivotStyleDef,
  type TableStyleElement,
} from "@cells/core"
import { cn } from "@/lib/cn"
import { PivotStyleThumb } from "@/ribbon/pivot-styles"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"
import { ElementFormatDialog } from "./format-cells/ElementFormatDialog"

/** How the dialog was opened: New PivotTable Style, Modify (a custom style), Duplicate (any style). */
export type PivotStyleRequest = { mode: "new" } | { mode: "modify"; name: string } | { mode: "duplicate"; name: string }

const PREVIEW_OPTIONS = { rowHeaders: true, colHeaders: true, rowStripes: true, colStripes: true, lastColumn: true }

/**
 * New PivotTable Style (and Modify / Duplicate from the gallery), as Excel's: a name, the Table Element list with
 * Format and Clear, Stripe Size for the stripes, the element's formatting in words, a preview, and Set as default
 * PivotTable style for this document.
 */
export function PivotStyleDialog({
  controller,
  request,
  onClose,
}: {
  controller: SheetController
  request: PivotStyleRequest
  onClose: () => void
}) {
  const styles = controller.pivotStyles
  const theme = controller.wb.theme
  const [draft, setDraft] = useState<CustomTableStyle>(() => {
    const all = controller.wb.tableStyles.styles
    if (request.mode === "new") return { name: newCustomStyleName(all), table: false, pivot: true, elements: {} }
    const source = styles.get(request.name)
    const elements = structuredClone(
      (source?.elements ?? pivotStyleDef(request.name, theme) ?? {}) as Record<string, TableStyleElement>,
    )
    return request.mode === "modify"
      ? { ...structuredClone(source!), elements }
      : { name: duplicateStyleName(request.name, all), table: false, pivot: true, elements }
  })
  const [isDefault, setIsDefault] = useState(() => request.mode === "modify" && styles.defaultStyle() === request.name)
  const [picked, setPicked] = useState(PIVOT_STYLE_ELEMENTS[0].key)
  const [formatting, setFormatting] = useState(false)
  const element = PIVOT_STYLE_ELEMENTS.find((e) => e.key === picked)!
  const current = draft.elements[picked]
  const setElement = (e: TableStyleElement | null) => {
    const elements = { ...draft.elements }
    if (e && !isEmptyElement(e)) elements[picked] = e
    else delete elements[picked]
    setDraft({ ...draft, elements })
  }
  const accept = () => {
    const problem = styles.save({ style: draft, isDefault }, request.mode === "modify" ? request.name : undefined)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }
  const title = request.mode === "modify" ? "Modify PivotTable Style" : "New PivotTable Style"
  return (
    <>
      <Dialog title={title} onCancel={onClose} onAccept={accept} width={520}>
        <div className="goal-seek-fields">
          <label className="dialog-label" htmlFor="pivot-style-name">
            Name:
          </label>
          <input
            id="pivot-style-name"
            className="dialog-field"
            autoFocus
            value={draft.name}
            onChange={(e) => setDraft({ ...draft, name: e.target.value })}
          />
        </div>
        <div className="pivot-style-editor">
          <div className="pivot-style-elements">
            <span className="dialog-label" id="pivot-style-elements-label">
              Table Element:
            </span>
            <div className="dialog-list pivot-style-list" role="listbox" aria-labelledby="pivot-style-elements-label">
              {PIVOT_STYLE_ELEMENTS.map((e) => (
                <button
                  key={e.key}
                  type="button"
                  role="option"
                  aria-selected={e.key === picked}
                  className={cn(
                    "dialog-list-item",
                    e.key === picked && "is-selected",
                    draft.elements[e.key] && "is-formatted",
                  )}
                  onClick={() => setPicked(e.key)}
                  onDoubleClick={() => setFormatting(true)}
                >
                  {e.label}
                </button>
              ))}
            </div>
            <div className="pivot-style-buttons">
              <button type="button" className="dialog-btn" onClick={() => setFormatting(true)}>
                Format
              </button>
              <button type="button" className="dialog-btn" disabled={!current} onClick={() => setElement(null)}>
                Clear
              </button>
            </div>
          </div>
          <div className="pivot-style-side">
            <span className="dialog-label">Preview</span>
            <div className="pivot-style-preview" data-keep-light>
              <PivotStyleThumb style={null} options={PREVIEW_OPTIONS} def={draft.elements as PivotStyleDef} scale={2} />
            </div>
            {element.stripe ? (
              <label className="pivot-style-stripe">
                <span>Stripe Size:</span>
                <select
                  className="dialog-field"
                  value={current?.stripe ?? 1}
                  onChange={(e) => {
                    const stripe = Number(e.target.value)
                    const next: TableStyleElement = { ...(current ?? {}) }
                    if (stripe > 1) next.stripe = stripe
                    else delete next.stripe
                    setElement(next)
                  }}
                >
                  {Array.from({ length: 9 }, (_, i) => (
                    <option key={i} value={i + 1}>
                      {i + 1}
                    </option>
                  ))}
                </select>
              </label>
            ) : null}
          </div>
        </div>
        <fieldset className="dialog-group">
          <legend>Element Formatting:</legend>
          <p className="pivot-style-summary">{elementSummary(current) || " "}</p>
        </fieldset>
        <label className="dialog-radio">
          <input type="checkbox" checked={isDefault} onChange={(e) => setIsDefault(e.target.checked)} />
          Set as default PivotTable style for this document
        </label>
      </Dialog>
      {formatting ? (
        <ElementFormatDialog
          element={current ?? {}}
          theme={theme}
          onAccept={setElement}
          onClose={() => setFormatting(false)}
        />
      ) : null}
    </>
  )
}
