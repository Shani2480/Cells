import { pivotLayoutOf } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { INDENT_PX } from "./cell-look"
import { PAD, PIVOT_BUTTON } from "./draw-cells"
import { colX, rowY, type GridLayout } from "./layout"

/**
 * The +/− buttons on the outer items of the sheet's PivotTables (PivotTable Options › Show expand/collapse
 * buttons). A click opens or closes the item, as one undo step.
 */
export function PivotButtons({
  controller,
  layout,
  zoom,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
}) {
  const sheet = controller.sheet
  if (!sheet.pivots.length) return null
  const rows = new Set(layout.rows.map((l) => l.index))
  const cols = new Set(layout.cols.map((l) => l.index))
  const buttons = []
  for (const table of sheet.pivots) {
    const cache = controller.wb.pivotCaches.find((c) => c.id === table.cacheId)
    if (!cache) continue
    for (const cell of pivotLayoutOf(table, cache).cells) {
      if (!cell.button || !rows.has(cell.row) || !cols.has(cell.col)) continue
      if (sheet.rows.isHidden(cell.row) || sheet.cols.isHidden(cell.col)) continue
      const h = sheet.rows.size(cell.row)
      const w = sheet.cols.size(cell.col)
      if (h < PIVOT_BUTTON + 2 || w < PIVOT_BUTTON + PAD + 2) continue
      const x = colX(sheet, layout, cell.col) + PAD + (cell.indent ?? 0) * INDENT_PX
      const y = rowY(sheet, layout, cell.row) + Math.round((h - PIVOT_BUTTON) / 2)
      if (x < layout.rowHeaderWidth || y < layout.headerHeight) continue
      const open = cell.button === "collapse"
      const label = `${open ? "Collapse" : "Expand"} ${String(cell.value ?? "")}`.trim()
      buttons.push(
        <button
          key={`${table.name}:${cell.row}:${cell.col}`}
          type="button"
          className="pivot-drill-button"
          style={{ left: x * zoom, top: y * zoom, width: PIVOT_BUTTON * zoom, height: PIVOT_BUTTON * zoom }}
          aria-label={label}
          aria-expanded={open}
          title={label}
          onPointerDown={(event) => event.stopPropagation()}
          onDoubleClick={(event) => event.stopPropagation()}
          onClick={() => controller.pivots.toggle(cell.row, cell.col)}
        >
          <svg width={7 * zoom} height={7 * zoom} viewBox="0 0 7 7" aria-hidden>
            <path d={open ? "M1 3.5h5" : "M1 3.5h5M3.5 1v5"} stroke="currentColor" strokeWidth="1" />
          </svg>
        </button>,
      )
    }
  }
  return <>{buttons}</>
}
