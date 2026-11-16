import { useMemo, useState } from "react"
import { SUBTOTAL_FUNCTIONS, type AxisName } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** Data › Group (or Ungroup) when the selection isn't whole rows or columns: which to group, as Excel asks. */
export function GroupDialog({
  controller,
  ungroup,
  onClose,
}: {
  controller: SheetController
  ungroup: boolean
  onClose: () => void
}) {
  const [axis, setAxis] = useState<AxisName>("rows")
  const accept = () => {
    if (ungroup) controller.outline.ungroup(axis)
    else controller.outline.group(axis)
    onClose()
  }
  const radio = (value: AxisName, label: string, letter: string) => (
    <label className="dialog-radio">
      <input type="radio" name="outline-axis" checked={axis === value} onChange={() => setAxis(value)} />
      <span>
        <u>{letter}</u>
        {label.slice(1)}
      </span>
    </label>
  )
  return (
    <Dialog title={ungroup ? "Ungroup" : "Group"} onCancel={onClose} onAccept={accept} width={240}>
      <fieldset className="dialog-group">
        <legend>{ungroup ? "Ungroup" : "Group"}</legend>
        {radio("rows", "Rows", "R")}
        {radio("cols", "Columns", "C")}
      </fieldset>
    </Dialog>
  )
}

/**
 * Data › Outline › Settings: where summary rows and columns sit. (Excel's Automatic styles, Create and Apply Styles,
 * which style the summary lines, aren't here yet.)
 */
export function OutlineSettingsDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [below, setBelow] = useState(controller.sheet.outline.summaryBelow)
  const [right, setRight] = useState(controller.sheet.outline.summaryRight)
  const accept = () => {
    controller.outline.setSettings({ ...controller.sheet.outline, summaryBelow: below, summaryRight: right })
    onClose()
  }
  return (
    <Dialog title="Settings" onCancel={onClose} onAccept={accept} width={300}>
      <fieldset className="dialog-group">
        <legend>Direction</legend>
        <label className="dialog-radio">
          <input type="checkbox" checked={below} onChange={(e) => setBelow(e.target.checked)} />
          <span>
            Summary <u>r</u>ows below detail
          </span>
        </label>
        <label className="dialog-radio">
          <input type="checkbox" checked={right} onChange={(e) => setRight(e.target.checked)} />
          <span>
            Summary <u>c</u>olumns to right of detail
          </span>
        </label>
      </fieldset>
    </Dialog>
  )
}

/** Data › Subtotal: a total at each change in a column, with the outline to show only the totals. */
export function SubtotalDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const range = useMemo(() => controller.outline.subtotalRange(), [controller])
  const columns = useMemo(() => {
    const out: { col: number; label: string }[] = []
    for (let c = range.c0; c <= range.c1; c++) {
      const v = controller.sheet.cells.value(range.r0, c)
      const text = v === undefined || v === null || v === "" ? "" : typeof v === "object" ? "" : String(v)
      out.push({ col: c, label: text || `Column ${columnLetter(c)}` })
    }
    return out
  }, [controller, range])
  const [by, setBy] = useState(range.c0)
  const [fn, setFn] = useState(0)
  const [cols, setCols] = useState<number[]>([range.c1])
  const [replace, setReplace] = useState(true)
  const [pageBreaks, setPageBreaks] = useState(false)
  const [below, setBelow] = useState(true)
  const accept = () => {
    const problem = controller.outline.subtotal(range, { by, fn, cols, replace, pageBreaks, below })
    if (problem) {
      controller.shell.alert(problem)
      return
    }
    onClose()
  }
  return (
    <Dialog
      title="Subtotal"
      onCancel={onClose}
      onAccept={accept}
      width={340}
      extra={{
        label: "Remove All",
        onSelect: () => {
          controller.outline.removeSubtotals(range)
          onClose()
        },
      }}
    >
      <label className="dialog-label" htmlFor="subtotal-by">
        <u>A</u>t each change in:
      </label>
      <select id="subtotal-by" className="dialog-field" value={by} onChange={(e) => setBy(Number(e.target.value))}>
        {columns.map((c) => (
          <option key={c.col} value={c.col}>
            {c.label}
          </option>
        ))}
      </select>
      <label className="dialog-label" htmlFor="subtotal-fn">
        <u>U</u>se function:
      </label>
      <select id="subtotal-fn" className="dialog-field" value={fn} onChange={(e) => setFn(Number(e.target.value))}>
        {SUBTOTAL_FUNCTIONS.map((f, i) => (
          <option key={f.label} value={i}>
            {f.label}
          </option>
        ))}
      </select>
      <span className="dialog-label" id="subtotal-cols">
        A<u>d</u>d subtotal to:
      </span>
      <div className="subtotal-columns" role="group" aria-labelledby="subtotal-cols">
        {columns.map((c) => (
          <label key={c.col} className="dialog-radio">
            <input
              type="checkbox"
              checked={cols.includes(c.col)}
              onChange={(e) =>
                setCols(e.target.checked ? [...cols, c.col].sort((a, b) => a - b) : cols.filter((x) => x !== c.col))
              }
            />
            <span>{c.label}</span>
          </label>
        ))}
      </div>
      <label className="dialog-radio">
        <input type="checkbox" checked={replace} onChange={(e) => setReplace(e.target.checked)} />
        <span>
          Replace <u>c</u>urrent subtotals
        </span>
      </label>
      <label className="dialog-radio">
        <input type="checkbox" checked={pageBreaks} onChange={(e) => setPageBreaks(e.target.checked)} />
        <span>
          <u>P</u>age break between groups
        </span>
      </label>
      <label className="dialog-radio">
        <input type="checkbox" checked={below} onChange={(e) => setBelow(e.target.checked)} />
        <span>
          <u>S</u>ummary below data
        </span>
      </label>
    </Dialog>
  )
}

function columnLetter(col: number): string {
  let s = ""
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}
