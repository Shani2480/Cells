import { useState } from "react"
import { columnName, CONSOLIDATE_FUNCTIONS, type ConsolidateSource } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"
import { parseSheetRange } from "./AdvancedFilterDialog"

/** A reference as typed, or a defined name, as the sheet and range it means. */
export function resolveReference(controller: SheetController, text: string): ConsolidateSource | null {
  const direct = parseSheetRange(controller, text)
  if (direct) return direct
  const bare = text.trim().replace(/^=/, "").toLowerCase()
  const name = controller.names.all().find((n) => n.name.toLowerCase() === bare)
  return name ? parseSheetRange(controller, name.ref) : null
}

/** How a reference is listed: with its sheet, absolute, as Excel shows it. */
function tidy(controller: SheetController, text: string): string {
  const r = parseSheetRange(controller, text)
  if (!r) return text.trim().replace(/^=/, "")
  const abs = (row: number, col: number) => `$${columnName(col)}$${row + 1}`
  const { r0, c0, r1, c1 } = r.range
  const area = r0 === r1 && c0 === c1 ? abs(r0, c0) : `${abs(r0, c0)}:${abs(r1, c1)}`
  const name = controller.wb.sheets[r.sheet].name
  const quoted = /^[A-Za-z_][\w.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
  return `${quoted}!${area}`
}

/**
 * Data › Data Tools › Consolidate: summarise ranges from this and other sheets into the selection, by position or by
 * their labels, with one function, as values or as formulas linked to the sources.
 */
export function ConsolidateDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const saved = controller.sheet.consolidate
  const [fn, setFn] = useState(saved?.fn ?? 0)
  const [refs, setRefs] = useState<string[]>(saved?.refs ?? [])
  const [chosen, setChosen] = useState<number | null>(null)
  const [reference, setReference] = useState(saved?.refs[0] ?? "")
  const [top, setTop] = useState(saved?.top ?? false)
  const [left, setLeft] = useState(saved?.left ?? false)
  const [links, setLinks] = useState(saved?.links ?? false)

  const add = () => {
    if (!reference.trim()) return
    if (!resolveReference(controller, reference)) return controller.shell.alert("Reference isn't valid.")
    const text = tidy(controller, reference)
    const at = refs.findIndex((r) => r.toLowerCase() === text.toLowerCase())
    if (at >= 0) return setChosen(at)
    setRefs([...refs, text].sort((a, b) => a.localeCompare(b)))
    setChosen(null)
  }

  const remove = () => {
    if (chosen === null) return
    setRefs(refs.filter((_, i) => i !== chosen))
    setChosen(null)
    setReference("")
  }

  const run = () => {
    // A reference typed but not added is used too, as in Excel.
    const all = [...refs]
    if (reference.trim() && !all.some((r) => r.toLowerCase() === tidy(controller, reference).toLowerCase())) {
      if (!resolveReference(controller, reference)) return controller.shell.alert("Reference isn't valid.")
      all.push(tidy(controller, reference))
    }
    const sources: ConsolidateSource[] = []
    for (const text of all) {
      const source = resolveReference(controller, text)
      if (!source) return controller.shell.alert(`Cannot open consolidate source file ${text}.`)
      sources.push(source)
    }
    const problem = controller.data.consolidate(sources, { fn, refs: all, top, left, links })
    if (problem) return controller.shell.alert(problem)
    onClose()
  }

  return (
    <Dialog
      title="Consolidate"
      onCancel={onClose}
      onAccept={run}
      width={440}
      actions={
        <>
          <button type="button" className="dialog-btn is-primary" onClick={run}>
            OK
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Close
          </button>
        </>
      }
    >
      <label className="dialog-label" htmlFor="consolidate-fn">
        <u>F</u>unction:
      </label>
      <select id="consolidate-fn" className="dialog-field" value={fn} onChange={(e) => setFn(Number(e.target.value))}>
        {CONSOLIDATE_FUNCTIONS.map((f, i) => (
          <option key={f.xml} value={i}>
            {f.label}
          </option>
        ))}
      </select>
      <label className="dialog-label consolidate-gap" htmlFor="consolidate-ref">
        <u>R</u>eference:
      </label>
      <div className="consolidate-row">
        <input
          id="consolidate-ref"
          className="dialog-field"
          autoFocus
          value={reference}
          placeholder="Sheet2!$A$1:$D$10"
          onChange={(e) => setReference(e.target.value)}
        />
      </div>
      <span className="dialog-label consolidate-gap" id="consolidate-list-label">
        All r<u>e</u>ferences:
      </span>
      <div className="consolidate-row">
        <div className="dialog-list consolidate-list" role="listbox" aria-labelledby="consolidate-list-label">
          {refs.map((r, i) => (
            <button
              key={r}
              type="button"
              role="option"
              aria-selected={i === chosen}
              className={`dialog-list-item${i === chosen ? " is-selected" : ""}`}
              onClick={() => {
                setChosen(i)
                setReference(r)
              }}
            >
              {r}
            </button>
          ))}
        </div>
        <div className="consolidate-buttons">
          <button type="button" className="dialog-btn" onClick={add}>
            <u>A</u>dd
          </button>
          <button type="button" className="dialog-btn" disabled={chosen === null} onClick={remove}>
            <u>D</u>elete
          </button>
        </div>
      </div>
      <fieldset className="dialog-group consolidate-gap">
        <legend>Use labels in</legend>
        <div className="consolidate-options">
          <div>
            <label className="dialog-radio">
              <input type="checkbox" checked={top} onChange={(e) => setTop(e.target.checked)} />
              <span>
                <u>T</u>op row
              </span>
            </label>
            <label className="dialog-radio">
              <input type="checkbox" checked={left} onChange={(e) => setLeft(e.target.checked)} />
              <span>
                <u>L</u>eft column
              </span>
            </label>
          </div>
          <label className="dialog-radio">
            <input type="checkbox" checked={links} onChange={(e) => setLinks(e.target.checked)} />
            <span>
              Create li<u>n</u>ks to source data
            </span>
          </label>
        </div>
      </fieldset>
    </Dialog>
  )
}
