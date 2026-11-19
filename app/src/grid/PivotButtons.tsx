import { pivotFilterTarget, pivotLayoutOf, type PivotFilterTarget, type PivotTable } from "@cells/core"
import { ArrowDown, ArrowUp, Filter } from "lucide-react"
import { useEffect, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { INDENT_PX } from "./cell-look"
import { PAD, PIVOT_BUTTON } from "./draw-cells"
import { colX, rowY, type GridLayout } from "./layout"
import { PivotFilterMenu } from "./PivotFilterMenu"

/** Excel's filter button: 17 px square at 100%, at the bottom right of its cell. */
const FILTER_BUTTON = 17

/**
 * The buttons on the sheet's PivotTables: the +/− buttons on outer items (PivotTable Options › Show expand/collapse
 * buttons), a click opening or closing the item, and the filter buttons on Row Labels / Column Labels, the fields'
 * headers and the report filters, with the dropdown of the one that's open. Each change is one undo step.
 */
export function PivotButtons({
  controller,
  layout,
  zoom,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  /** Back to the sheet (focus the grid's input). */
  onDone?: () => void
}) {
  const [open, setOpen] = useState<{
    row: number
    col: number
    target: PivotFilterTarget
    left: number
    top: number
  } | null>(null)
  const sheet = controller.sheet

  const place = (row: number, col: number) => {
    if (!layout.rows.some((l) => l.index === row) || !layout.cols.some((l) => l.index === col)) return null
    if (sheet.rows.isHidden(row) || sheet.cols.isHidden(col)) return null
    const w = sheet.cols.size(col)
    const h = sheet.rows.size(row)
    const size = Math.min(FILTER_BUTTON, h - 1, w - 1)
    if (size < 6) return null
    const x = colX(sheet, layout, col) + w - size - 1
    const y = rowY(sheet, layout, row) + h - size - 1
    if (x < layout.rowHeaderWidth || y < layout.headerHeight) return null
    return { x, y, size }
  }

  // Alt+Down on a cell with a filter button opens its dropdown.
  useEffect(() => {
    controller.pivots.openMenu = (row, col) => {
      const target = controller.pivots.filterTargetAt(row, col)
      const at = place(row, col)
      const host = document.querySelector<HTMLElement>("[data-testid=grid]")?.getBoundingClientRect()
      if (!target || !at || !host) return
      setOpen({ row, col, target, left: host.left + at.x * zoom, top: host.top + (at.y + at.size) * zoom + 1 })
    }
    return () => {
      controller.pivots.openMenu = null
    }
  })

  if (!sheet.pivots.length) return null
  const rows = new Set(layout.rows.map((l) => l.index))
  const cols = new Set(layout.cols.map((l) => l.index))
  const buttons = []
  const isOpen = (row: number, col: number) => open?.row === row && open.col === col
  for (const table of sheet.pivots) {
    const cache = controller.wb.pivotCaches.find((c) => c.id === table.cacheId)
    if (!cache) continue
    for (const cell of pivotLayoutOf(table, cache).cells) {
      const target = pivotFilterTarget(table, cache, cell)
      const at = target ? place(cell.row, cell.col) : null
      if (target && at) {
        const { filtered, sorted } = filterState(table, target)
        const name = String(cell.value ?? "")
        buttons.push(
          <button
            key={`f:${table.name}:${cell.row}:${cell.col}`}
            type="button"
            className={`filter-button${isOpen(cell.row, cell.col) ? " is-open" : ""}`}
            style={{ left: at.x * zoom, top: at.y * zoom, width: at.size * zoom, height: at.size * zoom }}
            aria-label={`${name} (${filtered ? "filtered" : "showing all"})`.trim()}
            aria-haspopup="menu"
            title={filtered ? `${name}: filtered` : undefined}
            onPointerDown={(event) => event.stopPropagation()}
            onDoubleClick={(event) => event.stopPropagation()}
            onClick={(event) => {
              const rect = event.currentTarget.getBoundingClientRect()
              if (isOpen(cell.row, cell.col)) setOpen(null)
              else setOpen({ row: cell.row, col: cell.col, target, left: rect.left, top: rect.bottom + 1 })
            }}
          >
            {filtered ? <Filter size={8 * zoom} strokeWidth={2} fill="currentColor" /> : null}
            {sorted && !filtered ? (
              sorted === "descending" ? (
                <ArrowDown size={9 * zoom} strokeWidth={2} />
              ) : (
                <ArrowUp size={9 * zoom} strokeWidth={2} />
              )
            ) : null}
            <svg
              width={(filtered || sorted ? 5 : 7) * zoom}
              height={(filtered || sorted ? 3 : 4) * zoom}
              viewBox="0 0 7 4"
              aria-hidden
            >
              <path d="M0 0h7L3.5 4z" fill="currentColor" />
            </svg>
          </button>,
        )
      }
      if (!cell.button || !rows.has(cell.row) || !cols.has(cell.col)) continue
      if (sheet.rows.isHidden(cell.row) || sheet.cols.isHidden(cell.col)) continue
      const h = sheet.rows.size(cell.row)
      const w = sheet.cols.size(cell.col)
      if (h < PIVOT_BUTTON + 2 || w < PIVOT_BUTTON + PAD + 2) continue
      const x = colX(sheet, layout, cell.col) + PAD + (cell.indent ?? 0) * INDENT_PX
      const y = rowY(sheet, layout, cell.row) + Math.round((h - PIVOT_BUTTON) / 2)
      if (x < layout.rowHeaderWidth || y < layout.headerHeight) continue
      const expanded = cell.button === "collapse"
      const label = `${expanded ? "Collapse" : "Expand"} ${String(cell.value ?? "")}`.trim()
      buttons.push(
        <button
          key={`${table.name}:${cell.row}:${cell.col}`}
          type="button"
          className="pivot-drill-button"
          style={{ left: x * zoom, top: y * zoom, width: PIVOT_BUTTON * zoom, height: PIVOT_BUTTON * zoom }}
          aria-label={label}
          aria-expanded={expanded}
          title={label}
          onPointerDown={(event) => event.stopPropagation()}
          onDoubleClick={(event) => event.stopPropagation()}
          onClick={() => controller.pivots.toggle(cell.row, cell.col)}
        >
          <svg width={7 * zoom} height={7 * zoom} viewBox="0 0 7 7" aria-hidden>
            <path d={expanded ? "M1 3.5h5" : "M1 3.5h5M3.5 1v5"} stroke="currentColor" strokeWidth="1" />
          </svg>
        </button>,
      )
    }
  }
  return (
    <>
      {buttons}
      {open ? (
        <PivotFilterMenu
          key={`${open.row}:${open.col}`}
          controller={controller}
          row={open.row}
          col={open.col}
          target={open.target}
          left={open.left}
          top={open.top}
          onClose={() => {
            setOpen(null)
            onDone?.()
          }}
        />
      ) : null}
    </>
  )
}

/** Whether a button's fields have items unticked, a filter, or a report filter's pick, and the sort of a sorted one. */
function filterState(table: PivotTable, target: PivotFilterTarget) {
  let filtered = false
  let sorted: "ascending" | "descending" | null = null
  for (const f of target.fields) {
    const field = table.fields[f]
    if (!field) continue
    if (
      field.items.some((i) => i.hidden) ||
      table.pages.some((p) => p.field === f && p.item !== null) ||
      table.filters?.some((x) => x.field === f)
    )
      filtered = true
    if (target.kind !== "page" && field.sort !== "manual") sorted ??= field.sort
  }
  return { filtered, sorted }
}
