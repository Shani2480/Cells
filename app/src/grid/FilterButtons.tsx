import { ArrowDown, ArrowUp, Filter } from "lucide-react"
import { useEffect, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { FilterMenu } from "./FilterMenu"
import { colX, rowY, type GridLayout } from "./layout"

/** Excel's filter button: 17 px square at 100%, at the bottom right of each header cell. */
const BUTTON = 17

/**
 * The AutoFilter's dropdown buttons on its header row (a funnel on a filtered column, an arrow on the sorted one),
 * and the dropdown of the one that's open.
 */
export function FilterButtons({
  controller,
  layout,
  zoom,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  /** Back to the sheet (focus the grid's input). */
  onDone: () => void
}) {
  const [open, setOpen] = useState<{ col: number; left: number; top: number } | null>(null)
  const sheet = controller.sheet
  const filter = sheet.autoFilter

  const place = (col: number) => {
    if (!filter) return null
    const row = filter.range.r0
    if (!layout.rows.some((l) => l.index === row) || !layout.cols.some((l) => l.index === col)) return null
    const w = sheet.cols.size(col)
    const h = sheet.rows.size(row)
    const size = Math.min(BUTTON, h - 1, w - 1)
    if (size < 6) return null
    const x = colX(sheet, layout, col) + w - size - 1
    const y = rowY(sheet, layout, row) + h - size - 1
    if (x < layout.rowHeaderWidth || y < layout.headerHeight) return null
    return { x, y, size }
  }

  // Alt+Down on a header cell opens its dropdown.
  useEffect(() => {
    controller.data.openMenu = (col) => {
      const at = place(col)
      const host = document.querySelector<HTMLElement>("[data-testid=grid]")?.getBoundingClientRect()
      if (!at || !host) return
      setOpen({ col, left: host.left + at.x * zoom, top: host.top + (at.y + at.size) * zoom + 1 })
    }
    return () => {
      controller.data.openMenu = null
    }
  })

  if (!filter) return null
  const cols: number[] = []
  for (let c = filter.range.c0; c <= filter.range.c1; c++) if (!sheet.cols.isHidden(c)) cols.push(c)
  const close = () => {
    setOpen(null)
    onDone()
  }
  return (
    <>
      {cols.map((col) => {
        const at = place(col)
        if (!at) return null
        const offset = col - filter.range.c0
        const filtered = Boolean(filter.columns[offset])
        const sorted = filter.sort?.col === offset ? filter.sort : null
        const header = sheet.cells.get(filter.range.r0, col)?.value
        return (
          <button
            key={col}
            type="button"
            className={`filter-button${open?.col === col ? " is-open" : ""}`}
            style={{ left: at.x * zoom, top: at.y * zoom, width: at.size * zoom, height: at.size * zoom }}
            aria-label={`${header ?? ""} (${filtered ? "filtered" : "showing all"})`.trim()}
            aria-haspopup="menu"
            title={filtered ? `${header ?? ""}: filtered` : undefined}
            onPointerDown={(event) => event.stopPropagation()}
            onClick={(event) => {
              const rect = event.currentTarget.getBoundingClientRect()
              setOpen(open?.col === col ? null : { col, left: rect.left, top: rect.bottom + 1 })
            }}
          >
            {filtered ? <Filter size={8 * zoom} strokeWidth={2} fill="currentColor" /> : null}
            {sorted && !filtered ? (
              sorted.descending ? (
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
          </button>
        )
      })}
      {open ? (
        <FilterMenu
          key={open.col}
          controller={controller}
          col={open.col}
          left={open.left}
          top={open.top}
          onClose={close}
          onDialog={(request) => controller.shell.filterDialog?.(request)}
        />
      ) : null}
    </>
  )
}
