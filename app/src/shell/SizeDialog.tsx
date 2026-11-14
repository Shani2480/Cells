import { useState } from "react"
import { Dialog } from "./Dialog"

/** Row Height (points) and Column Width (characters), as Excel's. */
export function SizeDialog({
  axis,
  initial,
  onApply,
  onClose,
}: {
  axis: "rows" | "cols"
  initial: number | null
  onApply: (value: number) => void
  onClose: () => void
}) {
  const [text, setText] = useState(initial === null ? "" : String(initial))
  const max = axis === "rows" ? 409 : 255
  const accept = () => {
    const value = Number(text)
    if (!Number.isFinite(value) || value < 0 || value > max) return
    onApply(value)
  }
  const label = axis === "rows" ? "Row height:" : "Column width:"
  return (
    <Dialog title={axis === "rows" ? "Row Height" : "Column Width"} onCancel={onClose} onAccept={accept} width={240}>
      <label className="dialog-label" htmlFor="size-value">
        {label}
      </label>
      <input
        id="size-value"
        className="dialog-field"
        inputMode="decimal"
        value={text}
        onChange={(event) => setText(event.target.value)}
      />
    </Dialog>
  )
}
