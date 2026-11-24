import { useState, type ReactNode } from "react"
import type { PivotSort } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/**
 * A PivotTable field's Sort dialog (More Sort Options): kept as arranged (Manual), or Ascending / Descending by the
 * field's own labels or by one of the value fields, with Excel's summary line.
 */
export function PivotSortDialog({
  controller,
  field,
  onClose,
}: {
  controller: SheetController
  field: number
  onClose: () => void
}) {
  const now = controller.pivots.sortDialog(field)
  const [sort, setSort] = useState<PivotSort>(now?.sort ?? "manual")
  // -1: the field's labels; otherwise the value field's index. Each direction keeps its own pick, as Excel's two lists.
  const [by, setBy] = useState<Record<"ascending" | "descending", number>>(() => ({
    ascending: now?.sort === "ascending" ? (now.by ?? -1) : -1,
    descending: now?.sort === "descending" ? (now.by ?? -1) : -1,
  }))
  if (!now) return null
  const pick = sort === "manual" ? null : by[sort] < 0 ? null : by[sort]
  const accept = () => {
    controller.pivots.setSort(field, sort, pick)
    onClose()
  }
  const summary =
    sort === "manual"
      ? `Drag items of the ${now.name} field to show them in any order`
      : `Sort ${now.name} by ${pick === null ? now.name : now.values[pick]} in ${sort} order`
  const direction = (dir: "ascending" | "descending", label: ReactNode) => (
    <>
      <label className="dialog-radio">
        <input
          type="radio"
          name="pivot-sort"
          checked={sort === dir}
          data-autofocus={sort === dir ? true : undefined}
          onChange={() => setSort(dir)}
        />
        <span>{label}</span>
      </label>
      <select
        className="dialog-field pivot-sort-by"
        aria-label={dir === "ascending" ? "Ascending by" : "Descending by"}
        disabled={sort !== dir}
        value={by[dir]}
        onChange={(e) => setBy((old) => ({ ...old, [dir]: Number(e.target.value) }))}
      >
        <option value={-1}>{now.name}</option>
        {now.values.map((v, i) => (
          <option key={i} value={i}>
            {v}
          </option>
        ))}
      </select>
    </>
  )
  return (
    <Dialog title={`Sort (${now.name})`} width={340} onCancel={onClose} onAccept={accept}>
      <fieldset className="dialog-group">
        <legend>Sort options</legend>
        <label className="dialog-radio">
          <input
            type="radio"
            name="pivot-sort"
            checked={sort === "manual"}
            data-autofocus={sort === "manual" ? true : undefined}
            onChange={() => setSort("manual")}
          />
          <span>
            <u>M</u>anual (you can drag items to rearrange them)
          </span>
        </label>
        {direction(
          "ascending",
          <>
            <u>A</u>scending (A to Z) by:
          </>,
        )}
        {direction(
          "descending",
          <>
            <u>D</u>escending (Z to A) by:
          </>,
        )}
      </fieldset>
      <fieldset className="dialog-group">
        <legend>Summary</legend>
        <p className="pivot-sort-summary">{summary}</p>
      </fieldset>
    </Dialog>
  )
}
