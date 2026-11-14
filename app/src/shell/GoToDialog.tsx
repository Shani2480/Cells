import { parseRange, type Range } from "@cells/core"
import { useState } from "react"
import { Dialog } from "./Dialog"

/**
 * Go To (F5, Ctrl+G): the references visited before and a Reference box. Excel's Special… button belongs to Go To
 * Special, which arrives with the data tools (Phase 4); it is left out rather than shown dead.
 */
export function GoToDialog({
  recent,
  onGo,
  onClose,
}: {
  recent: string[]
  onGo: (range: Range, text: string) => void
  onClose: () => void
}) {
  const [text, setText] = useState("")
  const [error, setError] = useState(false)
  const accept = () => {
    const range = parseRange(text)
    if (!range) {
      setError(true)
      return
    }
    onGo(range, text.trim().toUpperCase())
  }
  return (
    <Dialog title="Go To" onCancel={onClose} onAccept={accept} width={300}>
      <span className="dialog-label">
        <u>G</u>o to:
      </span>
      <div className="dialog-list" role="listbox" aria-label="Go to">
        {recent.map((item) => (
          <button
            type="button"
            role="option"
            aria-selected={item === text}
            key={item}
            className={`dialog-list-item${item === text ? " is-selected" : ""}`}
            onClick={() => setText(item)}
            onDoubleClick={() => {
              const range = parseRange(item)
              if (range) onGo(range, item)
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
