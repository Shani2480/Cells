import { dateSerial, displayText, filterChoices, type ColumnFilter, type FilterCondition } from "@cells/core"
import { useId, useMemo, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { todaySerial } from "@/sheet/controller"
import type { CustomOperator, CustomPreset } from "@/sheet/data"
import { Dialog } from "./Dialog"

const OPERATORS: { id: CustomOperator; label: string; date?: string }[] = [
  { id: "equals", label: "equals" },
  { id: "notEquals", label: "does not equal" },
  { id: "greater", label: "is greater than", date: "is after" },
  { id: "greaterOrEqual", label: "is greater than or equal to", date: "is after or equal to" },
  { id: "less", label: "is less than", date: "is before" },
  { id: "lessOrEqual", label: "is less than or equal to", date: "is before or equal to" },
  { id: "beginsWith", label: "begins with" },
  { id: "notBeginsWith", label: "does not begin with" },
  { id: "endsWith", label: "ends with" },
  { id: "notEndsWith", label: "does not end with" },
  { id: "contains", label: "contains" },
  { id: "notContains", label: "does not contain" },
]

/** A dialog row as the condition Excel saves: "begins with x" is "= x*", "does not contain x" is "<> *x*". */
export function conditionFor(op: CustomOperator, value: string): FilterCondition {
  switch (op) {
    case "equals":
      return { op: "eq", value }
    case "notEquals":
      return { op: "ne", value }
    case "greater":
      return { op: "gt", value }
    case "greaterOrEqual":
      return { op: "ge", value }
    case "less":
      return { op: "lt", value }
    case "lessOrEqual":
      return { op: "le", value }
    case "beginsWith":
      return { op: "eq", value: `${value}*` }
    case "notBeginsWith":
      return { op: "ne", value: `${value}*` }
    case "endsWith":
      return { op: "eq", value: `*${value}` }
    case "notEndsWith":
      return { op: "ne", value: `*${value}` }
    case "contains":
      return { op: "eq", value: `*${value}*` }
    case "notContains":
      return { op: "ne", value: `*${value}*` }
  }
}

/** A saved condition as the dialog shows it (the reverse of conditionFor). */
export function rowFor(c: FilterCondition): { op: CustomOperator; value: string } {
  const v = c.value
  const lead = v.startsWith("*") && v.length > 1
  const trail = v.endsWith("*") && !v.endsWith("~*") && v.length > 1
  if (c.op === "eq" || c.op === "ne") {
    const not = c.op === "ne"
    if (lead && trail && v.length > 2) return { op: not ? "notContains" : "contains", value: v.slice(1, -1) }
    if (trail && !lead) return { op: not ? "notBeginsWith" : "beginsWith", value: v.slice(0, -1) }
    if (lead && !trail) return { op: not ? "notEndsWith" : "endsWith", value: v.slice(1) }
    return { op: not ? "notEquals" : "equals", value: v }
  }
  const op = ({ gt: "greater", ge: "greaterOrEqual", lt: "less", le: "lessOrEqual" } as const)[c.op]
  return { op, value: v }
}

type Row = { op: CustomOperator | ""; value: string }

/** Custom AutoFilter: "Show rows where" the column meets one or two conditions, joined by And or Or. */
export function CustomFilterDialog({
  controller,
  col,
  preset,
  onClose,
}: {
  controller: SheetController
  col: number
  preset?: CustomPreset
  onClose: () => void
}) {
  const id = useId()
  const filter = controller.data.autoFilter
  const sheet = controller.sheet
  const styles = controller.wb.styles
  const current = filter?.columns[col - (filter?.range.c0 ?? 0)]
  const choices = useMemo(
    () => (filter ? filterChoices(sheet, styles, filter, col, todaySerial()) : null),
    [filter, sheet, styles, col],
  )
  const [rows, setRows] = useState<[Row, Row]>(() => {
    if (preset)
      return [
        { op: preset.first, value: "" },
        { op: preset.second ?? "", value: "" },
      ]
    if (current?.kind === "custom") {
      const [a, b] = current.conditions.map(rowFor)
      return [a ?? { op: "equals", value: "" }, b ?? { op: "", value: "" }]
    }
    return [
      { op: "equals", value: "" },
      { op: "", value: "" },
    ]
  })
  const [and, setAnd] = useState(preset?.and ?? (current?.kind === "custom" ? current.and !== false : true))
  if (!filter || !choices) return null
  const header = displayText(
    sheet.cells.get(filter.range.r0, col)?.value,
    styles.get(sheet.styleAt(filter.range.r0, col)).numberFormat,
  )
  const dates = choices.kind === "date"
  const suggestions = [
    ...choices.values.map((v) => v.text),
    ...choices.dates.map((d) => displayText(dateSerial(d.year, d.month, d.day), "m/d/yyyy")),
  ]

  const accept = () => {
    const conditions = rows
      .filter((r): r is { op: CustomOperator; value: string } => r.op !== "")
      .map((r) => conditionFor(r.op, r.value))
    onClose()
    if (!conditions.length) return
    const column: ColumnFilter = { kind: "custom", conditions }
    if (conditions.length === 2 && !and) column.and = false
    controller.data.filterColumn(col, column)
  }
  const set = (i: 0 | 1, patch: Partial<Row>) =>
    setRows((r) => (i === 0 ? [{ ...r[0], ...patch }, r[1]] : [r[0], { ...r[1], ...patch }]))

  const line = (i: 0 | 1) => (
    <div className="caf-line">
      <select
        className="dialog-field"
        aria-label={i === 0 ? "First condition" : "Second condition"}
        value={rows[i].op}
        onChange={(e) => set(i, { op: e.target.value as CustomOperator | "" })}
        data-autofocus={i === 0 && !preset ? true : undefined}
      >
        {i === 1 ? <option value="" /> : null}
        {OPERATORS.map((o) => (
          <option key={o.id} value={o.id}>
            {dates ? (o.date ?? o.label) : o.label}
          </option>
        ))}
      </select>
      <input
        className="dialog-field"
        aria-label={i === 0 ? "First value" : "Second value"}
        list={`${id}-values`}
        value={rows[i].value}
        onChange={(e) => set(i, { value: e.target.value })}
        data-autofocus={i === 0 && preset ? true : undefined}
      />
    </div>
  )

  return (
    <Dialog title="Custom AutoFilter" width={480} onCancel={onClose} onAccept={accept}>
      <div className="caf">
        <p className="dialog-message">Show rows where:</p>
        <p className="caf-header">{header}</p>
        {line(0)}
        <div className="caf-join" role="radiogroup" aria-label="Join">
          <label className="dialog-radio">
            <input type="radio" checked={and} onChange={() => setAnd(true)} />
            And
          </label>
          <label className="dialog-radio">
            <input type="radio" checked={!and} onChange={() => setAnd(false)} />
            Or
          </label>
        </div>
        {line(1)}
        <p className="caf-hint">
          Use ? to represent any single character
          <br />
          Use * to represent any series of characters
        </p>
        <datalist id={`${id}-values`}>
          {suggestions.map((s, i) => (
            <option key={s + i} value={s} />
          ))}
        </datalist>
      </div>
    </Dialog>
  )
}

/** Top 10 AutoFilter: Show Top or Bottom, how many, Items or Percent. */
export function Top10Dialog({
  controller,
  col,
  onClose,
}: {
  controller: SheetController
  col: number
  onClose: () => void
}) {
  const filter = controller.data.autoFilter
  const current = filter?.columns[col - (filter?.range.c0 ?? 0)]
  const start = current?.kind === "top10" ? current : { top: true, percent: false, count: 10 }
  const [top, setTop] = useState(start.top)
  const [count, setCount] = useState(String(start.count))
  const [percent, setPercent] = useState(start.percent)
  const valid = /^\d+$/.test(count.trim()) && Number(count) >= 1 && Number(count) <= (percent ? 100 : 500)
  const accept = () => {
    if (!valid) return
    onClose()
    controller.data.filterColumn(col, { kind: "top10", top, percent, count: Number(count) })
  }
  return (
    <Dialog title="Top 10 AutoFilter" width={360} onCancel={onClose} onAccept={accept}>
      <fieldset className="tiw-fieldset top10">
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
          max={percent ? 100 : 500}
          value={count}
          aria-invalid={!valid}
          onChange={(e) => setCount(e.target.value)}
        />
        <select
          className="dialog-field"
          aria-label="Items or percent"
          value={percent ? "percent" : "items"}
          onChange={(e) => setPercent(e.target.value === "percent")}
        >
          <option value="items">Items</option>
          <option value="percent">Percent</option>
        </select>
      </fieldset>
    </Dialog>
  )
}
