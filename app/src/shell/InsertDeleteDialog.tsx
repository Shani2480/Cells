import { useState } from "react"
import { Dialog } from "./Dialog"

type Choice = "right" | "down" | "left" | "up" | "row" | "col"

/** Excel's Insert and Delete dialogs (Ctrl++ / Ctrl+- on a cell range). */
export function InsertDeleteDialog({
  kind,
  onChoose,
  onClose,
}: {
  kind: "insert" | "delete"
  onChoose: (choice: Choice) => void
  onClose: () => void
}) {
  const options: { value: Choice; label: string; key: string }[] =
    kind === "insert"
      ? [
          { value: "right", label: "Shift cells right", key: "i" },
          { value: "down", label: "Shift cells down", key: "d" },
          { value: "row", label: "Entire row", key: "r" },
          { value: "col", label: "Entire column", key: "c" },
        ]
      : [
          { value: "left", label: "Shift cells left", key: "l" },
          { value: "up", label: "Shift cells up", key: "u" },
          { value: "row", label: "Entire row", key: "r" },
          { value: "col", label: "Entire column", key: "c" },
        ]
  const [choice, setChoice] = useState<Choice>(kind === "insert" ? "down" : "up")
  return (
    <Dialog
      title={kind === "insert" ? "Insert" : "Delete"}
      onCancel={onClose}
      onAccept={() => onChoose(choice)}
      width={240}
    >
      <div className="dialog-group" role="radiogroup" aria-label={kind === "insert" ? "Insert" : "Delete"}>
        {options.map((option) => (
          <label key={option.value} className="dialog-radio">
            <input
              type="radio"
              name="insert-delete"
              checked={choice === option.value}
              onChange={() => setChoice(option.value)}
            />
            {underline(option.label, option.key)}
          </label>
        ))}
      </div>
    </Dialog>
  )
}

function underline(label: string, key: string) {
  const index = label.toLowerCase().indexOf(key)
  if (index === -1) return label
  return (
    <span>
      {label.slice(0, index)}
      <u>{label[index]}</u>
      {label.slice(index + 1)}
    </span>
  )
}
