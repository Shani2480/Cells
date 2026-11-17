import { useId, useState } from "react"
import type { SparklineGroup } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import type { SparklineDialogInit } from "@/sheet/sparklines"
import { Dialog } from "./Dialog"

/**
 * Insert › Sparklines (Create Sparklines) and the Sparkline tab's Edit Data dialogs, Hidden & Empty Cell Settings,
 * the vertical axis's custom value and the date axis's range: Excel's wording and layout.
 */
export function SparklineDataDialog({
  controller,
  init,
  onClose,
}: {
  controller: SheetController
  init: SparklineDialogInit
  onClose: () => void
}) {
  const id = useId()
  const [data, setData] = useState(init.data)
  const [location, setLocation] = useState(init.kind === "editSingle" ? "" : init.location)
  const s = controller.sparklines
  const accept = () => {
    const ok =
      init.kind === "create"
        ? s.create(init.type, data, location)
        : init.kind === "editGroup"
          ? s.editGroup(data, location)
          : s.editSingle(data)
    if (ok) onClose()
  }
  const title =
    init.kind === "create" ? "Create Sparklines" : init.kind === "editGroup" ? "Edit Sparklines" : "Edit Sparkline Data"
  if (init.kind === "editSingle")
    return (
      <Dialog title={title} width={340} onCancel={onClose} onAccept={accept}>
        <label className="dialog-label" htmlFor={`${id}-d`}>
          Select a data source for the sparkline <b>{init.cell}</b>:
        </label>
        <input
          id={`${id}-d`}
          className="dialog-field"
          value={data}
          spellCheck={false}
          data-autofocus
          onChange={(e) => setData(e.target.value)}
        />
      </Dialog>
    )
  return (
    <Dialog title={title} width={360} onCancel={onClose} onAccept={accept}>
      <div className="spark-heading">Choose the data that you want</div>
      <label className="spark-row" htmlFor={`${id}-d`}>
        <span>
          <u>D</u>ata Range:
        </span>
        <input
          id={`${id}-d`}
          className="dialog-field"
          value={data}
          spellCheck={false}
          data-autofocus={!init.data || undefined}
          onChange={(e) => setData(e.target.value)}
        />
      </label>
      <div className="spark-heading">Choose where you want the sparklines to be placed</div>
      <label className="spark-row" htmlFor={`${id}-l`}>
        <span>
          <u>L</u>ocation Range:
        </span>
        <input
          id={`${id}-l`}
          className="dialog-field"
          value={location}
          spellCheck={false}
          data-autofocus={!!init.data || undefined}
          onChange={(e) => setLocation(e.target.value)}
        />
      </label>
    </Dialog>
  )
}

/** Edit Data › Hidden & Empty Cells. */
export function SparklineEmptyDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const g = controller.sparklines.active()
  const [empty, setEmpty] = useState<SparklineGroup["emptyAs"]>(g?.emptyAs ?? "gap")
  const [hidden, setHidden] = useState(g?.showHidden ?? false)
  const line = g?.type === "line"
  return (
    <Dialog
      title="Hidden and Empty Cell Settings"
      width={340}
      onCancel={onClose}
      onAccept={() => {
        controller.sparklines.setEmpty(empty, hidden)
        onClose()
      }}
    >
      <fieldset className="tiw-fieldset">
        <legend>Show empty cells as:</legend>
        {(
          [
            ["gap", "Gaps"],
            ["zero", "Zero"],
            ["span", "Connect data points with line"],
          ] as const
        ).map(([value, label]) => (
          <label key={value} className="dialog-radio">
            <input
              type="radio"
              name="spark-empty"
              checked={empty === value}
              disabled={value === "span" && !line}
              onChange={() => setEmpty(value)}
            />
            {label}
          </label>
        ))}
      </fieldset>
      <label className="dialog-radio">
        <input type="checkbox" checked={hidden} onChange={(e) => setHidden(e.target.checked)} />
        Show data in <u>h</u>idden rows and columns
      </label>
    </Dialog>
  )
}

/** Axis › Custom Value… for the vertical axis's minimum or maximum. */
export function SparklineAxisDialog({
  controller,
  which,
  onClose,
}: {
  controller: SheetController
  which: "min" | "max"
  onClose: () => void
}) {
  const g = controller.sparklines.active()
  const [text, setText] = useState(String((which === "min" ? g?.manualMin : g?.manualMax) ?? 0))
  const accept = () => {
    const value = Number(text.trim())
    if (!text.trim() || !Number.isFinite(value)) return controller.shell.alert("Enter a number.")
    controller.sparklines.setScale(which, "custom", value)
    onClose()
  }
  return (
    <Dialog title="Sparkline Vertical Axis Settings" width={320} onCancel={onClose} onAccept={accept}>
      <label className="dialog-label" htmlFor="spark-axis">
        Enter the {which === "min" ? "minimum" : "maximum"} value for the vertical axis:
      </label>
      <input
        id="spark-axis"
        className="dialog-field"
        inputMode="decimal"
        value={text}
        data-autofocus
        onChange={(e) => setText(e.target.value)}
      />
    </Dialog>
  )
}

/** Axis › Date Axis Type…: the range holding each point's date. */
export function SparklineDateDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [text, setText] = useState(controller.sparklines.dateRangeText())
  return (
    <Dialog
      title="Sparkline Date Range"
      width={340}
      onCancel={onClose}
      onAccept={() => {
        if (controller.sparklines.setDateRange(text)) onClose()
      }}
    >
      <label className="dialog-label" htmlFor="spark-dates">
        Select the range that contains the date values for the sparkline data:
      </label>
      <input
        id="spark-dates"
        className="dialog-field"
        value={text}
        spellCheck={false}
        data-autofocus
        onChange={(e) => setText(e.target.value)}
      />
    </Dialog>
  )
}
