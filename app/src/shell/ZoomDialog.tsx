import { useState } from "react"
import { MAX_ZOOM, MIN_ZOOM } from "@/sheet/controller"
import { Dialog } from "./Dialog"

const PRESETS = [2, 1, 0.75, 0.5, 0.25]

/** Excel's Zoom dialog: fixed magnifications, Fit selection, or a custom percentage (10–400). */
export function ZoomDialog({
  zoom,
  onZoom,
  onFitSelection,
  onClose,
}: {
  zoom: number
  onZoom: (zoom: number) => void
  onFitSelection: () => void
  onClose: () => void
}) {
  const preset = PRESETS.find((p) => Math.abs(p - zoom) < 0.001)
  const [choice, setChoice] = useState<number | "fit" | "custom">(preset ?? "custom")
  const [custom, setCustom] = useState(String(Math.round(zoom * 100)))
  const [error, setError] = useState<string | null>(null)
  const accept = () => {
    if (choice === "fit") onFitSelection()
    else if (choice === "custom") {
      const n = Number(custom.replace("%", "").trim())
      if (!Number.isFinite(n) || n < MIN_ZOOM * 100 || n > MAX_ZOOM * 100) {
        setError("Enter a number between 10 and 400.")
        return
      }
      onZoom(Math.round(n) / 100)
    } else onZoom(choice)
    onClose()
  }
  return (
    <Dialog title="Zoom" onCancel={onClose} onAccept={accept} width={240}>
      <fieldset className="zoom-options">
        <legend className="dialog-label">Magnification</legend>
        {PRESETS.map((p) => (
          <label key={p} className="fc-check">
            <input type="radio" name="zoom" checked={choice === p} onChange={() => setChoice(p)} />
            <span>{p * 100}%</span>
          </label>
        ))}
        <label className="fc-check">
          <input type="radio" name="zoom" checked={choice === "fit"} onChange={() => setChoice("fit")} />
          <span>Fit selection</span>
        </label>
        <label className="fc-check zoom-custom">
          <input type="radio" name="zoom" checked={choice === "custom"} onChange={() => setChoice("custom")} />
          <span>Custom:</span>
          <input
            className="dialog-field"
            aria-label="Custom zoom"
            value={custom}
            onFocus={() => setChoice("custom")}
            onChange={(e) => {
              setCustom(e.target.value)
              setError(null)
            }}
          />
          <span>%</span>
        </label>
      </fieldset>
      {error ? (
        <span className="dialog-label" role="alert">
          {error}
        </span>
      ) : null}
    </Dialog>
  )
}
