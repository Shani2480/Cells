import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * Alt Text for a chart: how a screen reader describes it to someone who can't see it (one or two sentences on what
 * it shows), or Mark as decorative when it adds nothing to the data beside it.
 */
export function AltTextDialog({
  controller,
  chartId,
  onClose,
}: {
  controller: SheetController
  chartId: number
  onClose: () => void
}) {
  const chart = controller.sheet.charts.find((c) => c.id === chartId)
  const [text, setText] = useState(chart?.altText ?? "")
  const [decorative, setDecorative] = useState(chart?.decorative ?? false)
  if (!chart) return null
  return (
    <Dialog
      title="Alt Text"
      width={380}
      onCancel={onClose}
      onAccept={() => {
        controller.charts.setAltText(chartId, text, decorative)
        onClose()
      }}
    >
      <p className="dialog-note">
        How would you describe this chart and what it shows to someone who is blind or has low vision?
      </p>
      <textarea
        className="dialog-field alt-text-field"
        aria-label={`Alt text for ${chart.name}`}
        data-autofocus
        rows={4}
        value={text}
        disabled={decorative}
        placeholder="For example: Monthly sales rise from 40 in January to 95 in June."
        onChange={(e) => setText(e.target.value)}
      />
      <label className="fc-check">
        <input type="checkbox" checked={decorative} onChange={(e) => setDecorative(e.target.checked)} />
        <span>
          <u>M</u>ark as decorative
        </span>
      </label>
      <p className="dialog-note">
        Decorative charts add visual interest but aren't informative. Screen readers skip them.
      </p>
    </Dialog>
  )
}
