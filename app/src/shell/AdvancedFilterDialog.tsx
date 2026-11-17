import { useState } from "react"
import { columnName, currentRegion, parseRange, type Range } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

const abs = (r: Range) => {
  const one = (row: number, col: number) => `$${columnName(col)}$${row + 1}`
  return r.r0 === r.r1 && r.c0 === r.c1 ? one(r.r0, r.c0) : `${one(r.r0, r.c0)}:${one(r.r1, r.c1)}`
}

const quoted = (name: string) => (/^[A-Za-z_][\w.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)

/** "A1:C9", "$A$1:$C$9" or "Sheet2!A1:B2": the sheet (the active one when not named) and the range. */
export function parseSheetRange(controller: SheetController, text: string): { sheet: number; range: Range } | null {
  const bare = text.trim().replace(/^=/, "")
  if (!bare) return null
  const bang = bare.lastIndexOf("!")
  let sheet = controller.sheetIndex
  if (bang >= 0) {
    const name = bare.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'").toLowerCase()
    sheet = controller.wb.sheets.findIndex((s) => s.name.toLowerCase() === name)
    if (sheet < 0) return null
  }
  const range = parseRange(bare.slice(bang + 1).replace(/\$/g, ""))
  return range ? { sheet, range } : null
}

/**
 * Data › Sort & Filter › Advanced: filter a list by a criteria range of conditions, in place or as a copy elsewhere
 * on the sheet, with an option to leave out repeated rows.
 */
export function AdvancedFilterDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const sheet = controller.sheet
  const previous = sheet.advancedFilter
  const [copy, setCopy] = useState(false)
  const [listText, setListText] = useState(() => {
    if (previous) return abs(previous.list)
    const sel = controller.selection.ranges[0]
    const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
    const region = single ? currentRegion(sheet, controller.selection.active) : sel
    const empty = single && region.r0 === region.r1 && region.c0 === region.c1 && !sheet.cells.get(region.r0, region.c0)
    return empty ? "" : abs(region)
  })
  const [criteriaText, setCriteriaText] = useState(() => {
    if (!previous) return ""
    const { sheet: s, range } = previous.criteria
    return s === controller.sheetIndex ? abs(range) : `${quoted(controller.wb.sheets[s].name)}!${abs(range)}`
  })
  const [toText, setToText] = useState("")
  const [unique, setUnique] = useState(previous?.unique ?? false)

  const run = () => {
    const list = parseSheetRange(controller, listText)
    if (!list || list.sheet !== controller.sheetIndex || list.range.r1 <= list.range.r0)
      return controller.shell.alert("The list range is not valid. It needs a row of headings and the data under it.")
    let criteria: { sheet: number; range: Range }
    if (criteriaText.trim()) {
      const c = parseSheetRange(controller, criteriaText)
      if (!c) return controller.shell.alert("Reference isn't valid.")
      criteria = c
    } else {
      // No criteria: every row (and with Unique records only, Excel's way to remove repeats).
      criteria = {
        sheet: controller.sheetIndex,
        range: { r0: list.range.r0, c0: list.range.c0, r1: list.range.r0, c1: list.range.c0 },
      }
    }
    let to: Range | null = null
    if (copy) {
      const t = parseSheetRange(controller, toText)
      if (!t) return controller.shell.alert("Reference isn't valid.")
      if (t.sheet !== controller.sheetIndex)
        return controller.shell.alert("You can only copy filtered data to the active sheet.")
      to = t.range
    }
    const problem = controller.data.advancedFilter({ list: list.range, criteria, unique }, to)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }

  return (
    <Dialog title="Advanced Filter" onCancel={onClose} onAccept={run} width={360}>
      <fieldset className="dialog-group">
        <legend>Action</legend>
        <label className="dialog-radio">
          <input type="radio" name="advanced-action" checked={!copy} onChange={() => setCopy(false)} />
          <span>
            <u>F</u>ilter the list, in-place
          </span>
        </label>
        <label className="dialog-radio">
          <input type="radio" name="advanced-action" checked={copy} onChange={() => setCopy(true)} />
          <span>
            C<u>o</u>py to another location
          </span>
        </label>
      </fieldset>
      <div className="goal-seek-fields advanced-filter-fields">
        <label className="dialog-label" htmlFor="advanced-list">
          <u>L</u>ist range:
        </label>
        <input
          id="advanced-list"
          className="dialog-field"
          autoFocus
          value={listText}
          onChange={(e) => setListText(e.target.value)}
        />
        <label className="dialog-label" htmlFor="advanced-criteria">
          <u>C</u>riteria range:
        </label>
        <input
          id="advanced-criteria"
          className="dialog-field"
          value={criteriaText}
          onChange={(e) => setCriteriaText(e.target.value)}
        />
        <label className="dialog-label" htmlFor="advanced-to">
          Copy <u>t</u>o:
        </label>
        <input
          id="advanced-to"
          className="dialog-field"
          disabled={!copy}
          value={toText}
          onChange={(e) => setToText(e.target.value)}
        />
      </div>
      <label className="dialog-radio advanced-unique">
        <input type="checkbox" checked={unique} onChange={(e) => setUnique(e.target.checked)} />
        <span>
          Unique <u>r</u>ecords only
        </span>
      </label>
    </Dialog>
  )
}
