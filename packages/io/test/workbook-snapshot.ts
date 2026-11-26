import type { Workbook } from "@cells/core"
import type { SheetView } from "../src/xlsx/sheet-read"

/**
 * What two readings of the same workbook should agree on, as plain data: each sheet's cells (value, what was typed,
 * format), column widths and row heights that differ from the default, hidden rows and columns, merges, notes and
 * links, plus the workbook's names and each sheet's view.
 */
export function snapshot(wb: Workbook, views: SheetView[], cols = 12, rows = 12) {
  return {
    names: wb.names.all().map((n) => ({ ...n })),
    sheets: wb.sheets.map((s, i) => {
      const cells: Record<string, unknown> = {}
      s.cells.forEach({ r0: 0, c0: 0, r1: 200, c1: 50 }, (r, c, d) => {
        const style = wb.styles.get(d.style ?? 0)
        if (d.value === undefined && !Object.keys(style).length) return
        cells[`${r},${c}`] = { value: d.value, input: d.input, style }
      })
      const widths: Record<number, number> = {}
      const heights: Record<number, number> = {}
      for (let c = 0; c < cols; c++)
        if (s.cols.size(c) !== s.cols.defaultSize || s.cols.isHidden(c))
          widths[c] = s.cols.isHidden(c) ? -1 : s.cols.size(c)
      for (let r = 0; r < rows; r++)
        if (s.rows.size(r) !== s.rows.defaultSize || s.rows.isHidden(r))
          heights[r] = s.rows.isHidden(r) ? -1 : s.rows.size(r)
      return {
        name: s.name,
        hidden: s.hidden,
        cells,
        widths,
        heights,
        merges: [...s.merges.all],
        notes: s.annotations.all(),
        links: s.hyperlinks.map((l) => ({ ...l })),
        view: { freeze: views[i]?.freeze, showGridlines: views[i]?.showGridlines },
      }
    }),
  }
}
