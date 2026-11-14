import type { SpecialKind, ValueKinds } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import type { FoundCell, SearchOptions } from "@/sheet/find"
import { cn } from "@/lib/cn"
import { Dialog } from "./Dialog"

const LOOK_IN: { id: SearchOptions["lookIn"]; label: string }[] = [
  { id: "formulas", label: "Formulas" },
  { id: "values", label: "Values" },
  { id: "notes", label: "Notes" },
  { id: "comments", label: "Comments" },
]

/**
 * Find and Replace (Ctrl+F, Ctrl+H): modeless, as Excel's, so the sheet can be clicked while it's open. Options >>
 * shows Within, Search, Look in, Match case and Match entire cell contents; Find All lists every match.
 */
export function FindReplaceDialog({
  controller,
  tab: initialTab,
  onClose,
}: {
  controller: SheetController
  tab: "find" | "replace"
  onClose: () => void
}) {
  const f = controller.find
  const [tab, setTab] = useState(initialTab)
  // It opens with the last search, as Excel's does.
  const [o, setO] = useState<SearchOptions>(() => ({ ...f.last }))
  const [replacement, setReplacement] = useState(f.replaceWith)
  const [options, setOptions] = useState(false)
  const [found, setFound] = useState<FoundCell[] | null>(null)
  const [picked, setPicked] = useState<number | null>(null)
  const set = (patch: Partial<SearchOptions>) => setO((x) => ({ ...x, ...patch }))
  // Replace looks in formulas only.
  const opts: SearchOptions = tab === "replace" ? { ...o, lookIn: "formulas" } : o

  const findAll = () => {
    const list = f.all(opts)
    setFound(list)
    setPicked(list.length ? 0 : null)
    if (list.length) f.goToFound(list[0])
    else controller.shell.alert("We couldn't find what you were looking for. Select Options for more ways to search.")
  }

  const button = (label: string, run: () => void, disabled = false, primary = false) => (
    <button type="button" className={cn("dialog-btn", primary && "is-primary")} disabled={disabled} onClick={run}>
      {label}
    </button>
  )
  const empty = !o.text

  return (
    <Dialog
      title="Find and Replace"
      width={options ? 560 : 520}
      modeless
      onCancel={onClose}
      onAccept={() => !empty && f.next(opts)}
      actions={
        <div className="fr-actions">
          {tab === "replace" ? (
            <>
              {button("Replace All", () => f.replaceAll(opts, replacement), empty)}
              {button("Replace", () => f.replace(opts, replacement), empty)}
            </>
          ) : null}
          <span className="fr-gap" />
          {button("Find All", findAll, empty)}
          {button("Find Next", () => f.next(opts), empty, true)}
          {button("Close", onClose)}
        </div>
      }
    >
      <div className="fr">
        <div className="fr-tabs" role="tablist" aria-label="Find and Replace">
          {(["find", "replace"] as const).map((t) => (
            <button
              key={t}
              type="button"
              role="tab"
              aria-selected={tab === t}
              className={cn("fr-tab", tab === t && "is-current")}
              onClick={() => setTab(t)}
            >
              {t === "find" ? "Find" : "Replace"}
            </button>
          ))}
        </div>
        <div className="fr-panel" role="tabpanel">
          <label className="fr-row">
            <span>
              Fi<u>n</u>d what:
            </span>
            <input
              className="dialog-field"
              value={o.text}
              data-autofocus
              aria-label="Find what"
              onChange={(e) => {
                set({ text: e.target.value })
                setFound(null)
              }}
            />
          </label>
          {tab === "replace" ? (
            <label className="fr-row">
              <span>
                R<u>e</u>place with:
              </span>
              <input
                className="dialog-field"
                value={replacement}
                aria-label="Replace with"
                onChange={(e) => setReplacement(e.target.value)}
              />
            </label>
          ) : null}
          {options ? (
            <div className="fr-options">
              <label className="fr-select">
                <span>
                  Wit<u>h</u>in:
                </span>
                <select
                  className="dialog-field"
                  aria-label="Within"
                  value={o.within}
                  onChange={(e) => set({ within: e.target.value as SearchOptions["within"] })}
                >
                  <option value="sheet">Sheet</option>
                  <option value="workbook">Workbook</option>
                </select>
              </label>
              <label className="dialog-radio">
                <input
                  type="checkbox"
                  checked={Boolean(o.matchCase)}
                  onChange={(e) => set({ matchCase: e.target.checked })}
                />
                <span>
                  Match <u>c</u>ase
                </span>
              </label>
              <label className="fr-select">
                <span>
                  <u>S</u>earch:
                </span>
                <select
                  className="dialog-field"
                  aria-label="Search"
                  value={o.byColumns ? "cols" : "rows"}
                  onChange={(e) => set({ byColumns: e.target.value === "cols" })}
                >
                  <option value="rows">By Rows</option>
                  <option value="cols">By Columns</option>
                </select>
              </label>
              <label className="dialog-radio">
                <input
                  type="checkbox"
                  checked={Boolean(o.entire)}
                  onChange={(e) => set({ entire: e.target.checked })}
                />
                <span>
                  Match entire cell c<u>o</u>ntents
                </span>
              </label>
              <label className="fr-select">
                <span>
                  <u>L</u>ook in:
                </span>
                <select
                  className="dialog-field"
                  aria-label="Look in"
                  value={opts.lookIn}
                  onChange={(e) => set({ lookIn: e.target.value as SearchOptions["lookIn"] })}
                >
                  {(tab === "replace" ? LOOK_IN.slice(0, 1) : LOOK_IN).map((l) => (
                    <option key={l.id} value={l.id}>
                      {l.label}
                    </option>
                  ))}
                </select>
              </label>
            </div>
          ) : null}
          <div className="fr-more">
            <button type="button" className="dialog-btn" onClick={() => setOptions(!options)}>
              Op<u>t</u>ions {options ? "<<" : ">>"}
            </button>
          </div>
        </div>
      </div>
      {found && found.length ? (
        <div className="fr-found">
          <div className="fr-found-scroll">
            <table aria-label="Found cells">
              <thead>
                <tr>
                  <th>Sheet</th>
                  <th>Cell</th>
                  <th>Value</th>
                  <th>Formula</th>
                </tr>
              </thead>
              <tbody>
                {found.map((x, i) => (
                  <tr
                    key={`${x.sheet}:${x.row}:${x.col}`}
                    aria-selected={picked === i}
                    className={picked === i ? "is-selected" : undefined}
                    onClick={() => {
                      setPicked(i)
                      f.goToFound(x)
                    }}
                  >
                    <td>{x.sheetName}</td>
                    <td>{x.cell}</td>
                    <td>{x.value}</td>
                    <td>{x.formula}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="fr-count">
            {found.length} cell{found.length === 1 ? "" : "s"} found
          </p>
        </div>
      ) : null}
    </Dialog>
  )
}

type SpecialChoice = SpecialKind["kind"] | "objects" | "validation"

/** Go To Special (F5 › Special…, or Find & Select): pick cells by what they hold. */
export function GoToSpecialDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [choice, setChoice] = useState<SpecialChoice>("constants")
  const [types, setTypes] = useState<ValueKinds>({ numbers: true, text: true, logicals: true, errors: true })
  const [all, setAll] = useState(false)
  const [same, setSame] = useState(false)

  const accept = () => {
    let kind: SpecialKind | null = null
    switch (choice) {
      case "constants":
      case "formulas":
        kind = { kind: choice, types }
        break
      case "precedents":
      case "dependents":
        kind = { kind: choice, all }
        break
      case "conditionalFormats":
        kind = { kind: choice, same }
        break
      case "objects":
      case "validation":
        break
      default:
        kind = { kind: choice } as SpecialKind
    }
    if (!kind) {
      controller.shell.alert("No cells were found.")
      return
    }
    if (controller.find.special(kind)) onClose()
  }
  const radio = (id: SpecialChoice, label: string, disabled = false) => (
    <label className={cn("dialog-radio", disabled && "is-disabled")} title={disabled ? "Not built yet." : undefined}>
      <input type="radio" name="special" checked={choice === id} disabled={disabled} onChange={() => setChoice(id)} />
      {label}
    </label>
  )
  const typeBox = (key: keyof ValueKinds, label: string) => (
    <label className={cn("dialog-radio gs-sub", choice !== "constants" && choice !== "formulas" && "is-disabled")}>
      <input
        type="checkbox"
        checked={types[key]}
        disabled={choice !== "constants" && choice !== "formulas"}
        onChange={(e) => setTypes({ ...types, [key]: e.target.checked })}
      />
      {label}
    </label>
  )
  const depends = choice === "precedents" || choice === "dependents"
  const cf = choice === "conditionalFormats" || choice === "validation"
  return (
    <Dialog title="Go To Special" width={420} onCancel={onClose} onAccept={accept}>
      <fieldset className="tiw-fieldset gs">
        <legend>Select</legend>
        <div className="gs-col">
          {radio("notes", "Notes")}
          {radio("constants", "Constants")}
          {radio("formulas", "Formulas")}
          {typeBox("numbers", "Numbers")}
          {typeBox("text", "Text")}
          {typeBox("logicals", "Logicals")}
          {typeBox("errors", "Errors")}
          {radio("blanks", "Blanks")}
          {radio("currentRegion", "Current region")}
          {radio("currentArray", "Current array")}
          {radio("objects", "Objects", true)}
        </div>
        <div className="gs-col">
          {radio("rowDifferences", "Row differences")}
          {radio("columnDifferences", "Column differences")}
          {radio("precedents", "Precedents")}
          {radio("dependents", "Dependents")}
          <label className={cn("dialog-radio gs-sub", !depends && "is-disabled")}>
            <input type="radio" name="levels" disabled={!depends} checked={!all} onChange={() => setAll(false)} />
            Direct only
          </label>
          <label className={cn("dialog-radio gs-sub", !depends && "is-disabled")}>
            <input type="radio" name="levels" disabled={!depends} checked={all} onChange={() => setAll(true)} />
            All levels
          </label>
          {radio("lastCell", "Last cell")}
          {radio("visible", "Visible cells only")}
          {radio("conditionalFormats", "Conditional formats")}
          {radio("validation", "Data validation", true)}
          <label className={cn("dialog-radio gs-sub", !cf && "is-disabled")}>
            <input type="radio" name="same" disabled={!cf} checked={!same} onChange={() => setSame(false)} />
            All
          </label>
          <label className={cn("dialog-radio gs-sub", !cf && "is-disabled")}>
            <input type="radio" name="same" disabled={!cf} checked={same} onChange={() => setSame(true)} />
            Same
          </label>
        </div>
      </fieldset>
    </Dialog>
  )
}
