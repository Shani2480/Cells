import { useState } from "react"
import type { Target } from "@/sheet/names"
import { Dialog } from "./Dialog"

/**
 * Go To (F5, Ctrl+G): the defined names and the references visited before, and a Reference box (a reference or a
 * name), and Special… for Go To Special.
 */
export function GoToDialog({
  names = [],
  recent,
  resolve,
  onGo,
  onSpecial,
  onClose,
}: {
  names?: readonly string[]
  recent: string[]
  resolve: (text: string) => Target | null
  onGo: (target: Target, text: string) => void
  /** Special…: Go To Special. */
  onSpecial?: () => void
  onClose: () => void
}) {
  const [text, setText] = useState("")
  const [error, setError] = useState(false)
  const accept = () => {
    const target = resolve(text)
    if (!target) {
      setError(true)
      return
    }
    onGo(
      target,
      names.some((n) => n.toUpperCase() === text.trim().toUpperCase()) ? text.trim() : text.trim().toUpperCase(),
    )
  }
  const items = [...names, ...recent.filter((r) => !names.includes(r))]
  return (
    <Dialog
      title="Go To"
      onCancel={onClose}
      onAccept={accept}
      width={300}
      actions={
        <div className="fr-actions">
          {onSpecial ? (
            <button type="button" className="dialog-btn" onClick={onSpecial}>
              <u>S</u>pecial...
            </button>
          ) : null}
          <span className="fr-gap" />
          <button type="button" className="dialog-btn is-primary" onClick={accept}>
            OK
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </div>
      }
    >
      <span className="dialog-label">
        <u>G</u>o to:
      </span>
      <div className="dialog-list" role="listbox" aria-label="Go to">
        {items.map((item) => (
          <button
            type="button"
            role="option"
            aria-selected={item === text}
            key={item}
            className={`dialog-list-item${item === text ? " is-selected" : ""}`}
            onClick={() => setText(item)}
            onDoubleClick={() => {
              const target = resolve(item)
              if (target) onGo(target, item)
            }}
          >
            {item}
          </button>
        ))}
      </div>
      <label className="dialog-label" htmlFor="goto-reference">
        <u>R</u>eference:
      </label>
      <input
        id="goto-reference"
        className="dialog-field"
        value={text}
        aria-invalid={error}
        onChange={(event) => {
          setText(event.target.value)
          setError(false)
        }}
      />
      {error ? (
        <span className="dialog-label" role="alert">
          Reference isn't valid.
        </span>
      ) : null}
    </Dialog>
  )
}
