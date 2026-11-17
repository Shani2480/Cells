import { useState } from "react"
import { Dialog } from "./Dialog"

type Fill = "all" | "contents" | "formats"

/** Home › Fill › Across Worksheets (with grouped sheets): what to copy to the same cells of the other sheets. */
export function FillAcrossDialog({ onFill, onClose }: { onFill: (what: Fill) => void; onClose: () => void }) {
  const [what, setWhat] = useState<Fill>("all")
  return (
    <Dialog
      title="Fill Across Worksheets"
      onCancel={onClose}
      onAccept={() => {
        onFill(what)
        onClose()
      }}
      width={240}
    >
      <fieldset className="zoom-options">
        <legend className="dialog-label">Fill</legend>
        {(
          [
            ["all", "All"],
            ["contents", "Contents"],
            ["formats", "Formats"],
          ] as const
        ).map(([value, label]) => (
          <label key={value} className="fc-check">
            <input type="radio" name="fill-across" checked={what === value} onChange={() => setWhat(value)} />
            <span>{label}</span>
          </label>
        ))}
      </fieldset>
    </Dialog>
  )
}
