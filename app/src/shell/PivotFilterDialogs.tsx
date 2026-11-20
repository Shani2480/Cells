import type { PivotFilterSpec, PivotLabelOp, PivotValueOp } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import type { PivotFilterRequest } from "@/sheet/pivots"
import { Dialog } from "./Dialog"

/** Label Filters' operators, in Excel's order, as its dialog words them. */
export const LABEL_OPS: { op: PivotLabelOp; menu: string; dialog: string }[] = [
  { op: "equal", menu: "Equals...", dialog: "equals" },
  { op: "notEqual", menu: "Does Not Equal...", dialog: "does not equal" },
  { op: "beginsWith", menu: "Begins With...", dialog: "begins with" },
  { op: "notBeginsWith", menu: "Does Not Begin With...", dialog: "does not begin with" },
  { op: "endsWith", menu: "Ends With...", dialog: "ends with" },
  { op: "notEndsWith", menu: "Does Not End With...", dialog: "does not end with" },
  { op: "contains", menu: "Contains...", dialog: "contains" },
  { op: "notContains", menu: "Does Not Contain...", dialog: "does not contain" },
  { op: "greaterThan", menu: "Greater Than...", dialog: "is greater than" },
  { op: "greaterThanOrEqual", menu: "Greater Than Or Equal To...", dialog: "is greater than or equal to" },
  { op: "lessThan", menu: "Less Than...", dialog: "is less than" },
  { op: "lessThanOrEqual", menu: "Less Than Or Equal To...", dialog: "is less than or equal to" },
  { op: "between", menu: "Between...", dialog: "is between" },
  { op: "notBetween", menu: "Not Between...", dialog: "is not between" },
]

/** Value Filters' operators. */
export const VALUE_OPS: { op: PivotValueOp; menu: string; dialog: string }[] = [
  { op: "equal", menu: "Equals...", dialog: "equals" },
  { op: "notEqual", menu: "Does Not Equal...", dialog: "does not equal" },
  { op: "greaterThan", menu: "Greater Than...", dialog: "is greater than" },
  { op: "greaterThanOrEqual", menu: "Greater Than Or Equal To...", dialog: "is greater than or equal to" },
  { op: "lessThan", menu: "Less Than...", dialog: "is less than" },
  { op: "lessThanOrEqual", menu: "Less Than Or Equal To...", dialog: "is less than or equal to" },
  { op: "between", menu: "Between...", dialog: "is between" },
  { op: "notBetween", menu: "Not Between...", dialog: "is not between" },
]

const ranged = (op: string) => op === "between" || op === "notBetween"
const num = (t: string) => (t.trim() !== "" && Number.isFinite(Number(t)) ? Number(t) : null)

/**
 * A PivotTable field's Label Filter, Value Filter or Top 10 Filter dialog, as Excel's: it opens on the field's
 * current filter of that kind (or the operator picked in the submenu) and OK sets it as one undo step.
 */
export function PivotFilterDialog({
  controller,
  request,
  onClose,
}: {
  controller: SheetController
  request: PivotFilterRequest
  onClose: () => void
}) {
  const { row, col, field } = request
  const pivots = controller.pivots.on(request.sheet)
  const choices = pivots.choices(row, col, field)
  const measures = pivots.valueNames(row, col)
  const current = choices?.filter
  const same = current?.kind === request.kind ? current : null
  const [op, setOp] = useState<string>(request.op ?? (same && same.kind !== "top" ? same.op : "equal"))
  const [value, setValue] = useState(same && same.kind !== "top" ? String(same.value) : "")
  const [value2, setValue2] = useState(same && same.kind !== "top" ? String(same.value2 ?? "") : "")
  const [measure, setMeasure] = useState(same && same.kind !== "label" ? same.measure : 0)
  const [top, setTop] = useState(same?.kind === "top" ? same.top : true)
  const [count, setCount] = useState(same?.kind === "top" ? String(same.value) : "10")
  const [by, setBy] = useState<"count" | "percent" | "sum">(same?.kind === "top" ? same.by : "count")
  const name = choices?.name ?? ""

  const spec = ((): PivotFilterSpec | null => {
    if (request.kind === "label") {
      const s: PivotFilterSpec = { kind: "label", op: op as PivotLabelOp, value }
      if (ranged(op)) s.value2 = value2
      return s
    }
    if (request.kind === "value") {
      const a = num(value)
      const b = num(value2)
      if (a === null || (ranged(op) && b === null)) return null
      const s: PivotFilterSpec = { kind: "value", op: op as PivotValueOp, measure, value: a }
      if (ranged(op)) s.value2 = b!
      return s
    }
    const n = num(count)
    if (n === null || n <= 0 || (by === "count" && !Number.isInteger(n)) || (by === "percent" && n > 100)) return null
    return { kind: "top", by, top, measure, value: n }
  })()

  const accept = () => {
    if (!spec) return
    onClose()
    pivots.setFilter(row, col, field, spec)
  }
  const measureSelect = (
    <select
      className="dialog-field"
      aria-label="Value field"
      value={measure}
      onChange={(e) => setMeasure(Number(e.target.value))}
    >
      {measures.map((m, i) => (
        <option key={i} value={i}>
          {m}
        </option>
      ))}
    </select>
  )

  if (request.kind === "top")
    return (
      <Dialog title={`Top 10 Filter (${name})`} width={460} onCancel={onClose} onAccept={accept}>
        <fieldset className="tiw-fieldset pivot-top10">
          <legend>Show</legend>
          <select
            className="dialog-field"
            aria-label="Top or bottom"
            value={top ? "top" : "bottom"}
            onChange={(e) => setTop(e.target.value === "top")}
          >
            <option value="top">Top</option>
            <option value="bottom">Bottom</option>
          </select>
          <input
            className="dialog-field"
            type="number"
            aria-label="How many"
            min={1}
            value={count}
            aria-invalid={!spec}
            onChange={(e) => setCount(e.target.value)}
          />
          <select
            className="dialog-field"
            aria-label="Items, percent or sum"
            value={by}
            onChange={(e) => setBy(e.target.value as typeof by)}
          >
            <option value="count">Items</option>
            <option value="percent">Percent</option>
            <option value="sum">Sum</option>
          </select>
          <span>by</span>
          {measureSelect}
        </fieldset>
      </Dialog>
    )

  const ops = request.kind === "label" ? LABEL_OPS : VALUE_OPS
  return (
    <Dialog
      title={`${request.kind === "label" ? "Label" : "Value"} Filter (${name})`}
      width={request.kind === "label" ? 480 : 600}
      onCancel={onClose}
      onAccept={accept}
    >
      <p className="pivot-filter-lead">
        {request.kind === "label" ? "Show items for which the label" : "Show items for which"}
      </p>
      <div className="pivot-filter-row">
        {request.kind === "value" ? measureSelect : null}
        <select className="dialog-field" aria-label="Comparison" value={op} onChange={(e) => setOp(e.target.value)}>
          {ops.map((o) => (
            <option key={o.op} value={o.op}>
              {o.dialog}
            </option>
          ))}
        </select>
        <input
          className="dialog-field"
          aria-label={ranged(op) ? "From" : "Value"}
          value={value}
          aria-invalid={request.kind === "value" && value.trim() !== "" && num(value) === null}
          onChange={(e) => setValue(e.target.value)}
        />
        {ranged(op) ? (
          <>
            <span>and</span>
            <input
              className="dialog-field"
              aria-label="To"
              value={value2}
              onChange={(e) => setValue2(e.target.value)}
            />
          </>
        ) : null}
      </div>
      {request.kind === "label" ? (
        <p className="pivot-filter-hint">
          Use ? to represent any single character. Use * to represent any series of characters.
        </p>
      ) : null}
    </Dialog>
  )
}
