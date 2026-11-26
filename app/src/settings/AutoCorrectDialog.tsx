import { useMemo, useState } from "react"
import { autoCorrectEntries, BUILT_IN_ENTRIES, type AutoCorrectSettings } from "@/proofing/autocorrect"
import { Dialog } from "@/shell/Dialog"

/**
 * Excel's AutoCorrect dialog (File › Options › Proofing › AutoCorrect Options…), its AutoCorrect tab: the four fixes,
 * Replace text as you type, and the list. Typing in Replace finds the entry; Add puts a new one in (Replace when
 * the name is in the list), Delete takes the selected one out. OK hands the result back to Excel Options.
 */
export function AutoCorrectDialog({
  value,
  onDone,
  onCancel,
}: {
  value: AutoCorrectSettings
  onDone: (next: AutoCorrectSettings) => void
  onCancel: () => void
}) {
  const [draft, setDraft] = useState<AutoCorrectSettings>(value)
  const [from, setFrom] = useState("")
  const [to, setTo] = useState("")
  const entries = useMemo(() => [...autoCorrectEntries(draft)].sort(([a], [b]) => a.localeCompare(b)), [draft])
  const found = entries.find(([f]) => f === from)
  const built = new Set(BUILT_IN_ENTRIES.map(([f]) => f))

  const check = (key: "twoInitialCaps" | "sentenceCaps" | "dayNames" | "capsLock" | "replaceText", label: string) => (
    <label className="dialog-radio">
      <input type="checkbox" checked={draft[key]} onChange={(e) => setDraft({ ...draft, [key]: e.target.checked })} />
      <span>{label}</span>
    </label>
  )
  const add = () => {
    const name = from.trim()
    if (!name || !to) return
    setDraft({
      ...draft,
      added: [...draft.added.filter(([f]) => f !== name), [name, to]],
      removed: draft.removed.filter((f) => f !== name),
    })
  }
  const remove = () => {
    if (!found) return
    setDraft({
      ...draft,
      added: draft.added.filter(([f]) => f !== found[0]),
      removed: built.has(found[0]) ? [...draft.removed, found[0]] : draft.removed,
    })
    setFrom("")
    setTo("")
  }

  return (
    <Dialog title="AutoCorrect: English (United States)" width={460} onCancel={onCancel} onAccept={() => onDone(draft)}>
      <div className="autocorrect">
        {check("twoInitialCaps", "Correct TWo INitial CApitals")}
        {check("sentenceCaps", "Capitalize first letter of sentences")}
        {check("dayNames", "Capitalize names of days")}
        {check("capsLock", "Correct accidental use of cAPS LOCK key")}
        {check("replaceText", "Replace text as you type")}
        <div className="autocorrect-fields">
          <label className="dialog-label" htmlFor="autocorrect-from">
            Replace:
          </label>
          <label className="dialog-label" htmlFor="autocorrect-to">
            With:
          </label>
          <input
            id="autocorrect-from"
            className="dialog-field"
            value={from}
            disabled={!draft.replaceText}
            onChange={(e) => {
              setFrom(e.target.value)
              const hit = entries.find(([f]) => f === e.target.value)
              if (hit) setTo(hit[1])
            }}
          />
          <input
            id="autocorrect-to"
            className="dialog-field"
            value={to}
            disabled={!draft.replaceText}
            onChange={(e) => setTo(e.target.value)}
          />
        </div>
        <div className="dialog-list autocorrect-list" role="listbox" aria-label="AutoCorrect entries">
          {entries.map(([f, t]) => (
            <button
              key={f}
              type="button"
              role="option"
              aria-selected={f === from}
              aria-label={`${f}, ${t}`}
              className={`dialog-list-item autocorrect-row${f === from ? " is-selected" : ""}`}
              disabled={!draft.replaceText}
              onClick={() => {
                setFrom(f)
                setTo(t)
              }}
            >
              <span>{f}</span>
              <span>{t}</span>
            </button>
          ))}
        </div>
        <div className="autocorrect-buttons">
          <button
            type="button"
            className="dialog-btn"
            disabled={!draft.replaceText || !from.trim() || !to || (found && found[1] === to)}
            onClick={add}
          >
            {found ? "Replace" : "Add"}
          </button>
          <button type="button" className="dialog-btn" disabled={!draft.replaceText || !found} onClick={remove}>
            Delete
          </button>
        </div>
      </div>
    </Dialog>
  )
}
