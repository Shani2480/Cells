import type { AxisName, OutlineGroup } from "@cells/core"
import { isCollapsed } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { colX, rowY, type GridLayout } from "./layout"

/** The width of one outline level's lane, as Excel's (about 16 px at 100%). */
export const LANE = 16

/** How thick the bar beside the row (or above the column) headings is: one lane per level plus one, or 0. */
export function outlineBarSize(controller: SheetController, axis: AxisName): number {
  const s = controller.sheet
  if (!s.outline.showSymbols) return 0
  const depth = (axis === "rows" ? s.rows : s.cols).maxLevel
  return depth ? LANE * (depth + 1) + 4 : 0
}

/**
 * The outline bar for rows (left of the row headings) or columns (above the column headings): level buttons 1…n+1
 * level with the headings, then for each group a bracket along its detail ending in − at its summary line, or + when
 * the group is collapsed. Positions follow the sheet's scroll, frozen panes and zoom.
 */
export function OutlineBar({
  controller,
  layout,
  zoom,
  axis,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  axis: AxisName
  onDone: () => void
}) {
  const size = outlineBarSize(controller, axis)
  if (!size) return null
  const sheet = controller.sheet
  const lines = axis === "rows" ? sheet.rows : sheet.cols
  const depth = lines.maxLevel
  const rows = axis === "rows"
  // Along the axis, in screen pixels.
  const start = (i: number) => (rows ? rowY(sheet, layout, i) : colX(sheet, layout, i)) * zoom
  const length = (i: number) => lines.size(i) * zoom
  const head = (rows ? layout.headerHeight : layout.rowHeaderWidth) * zoom
  const end = (rows ? layout.height : layout.width) * zoom
  const lane = (level: number) => 2 + LANE * (level - 1) + LANE / 2
  const groups = controller.outline.groups(axis)
  const visible = (a: number, b: number) => b >= head && a <= end

  const marks: React.ReactNode[] = []
  groups.forEach((g: OutlineGroup, k) => {
    const collapsed = isCollapsed(lines, g)
    const hasSummary = g.summary >= 0 && g.summary < lines.count
    const sumAt = hasSummary ? start(g.summary) + length(g.summary) / 2 : null
    // A group inside a collapsed parent isn't drawn: its lines are hidden.
    if (hasSummary && lines.isHidden(g.summary)) return
    const c = lane(g.level)
    if (!collapsed) {
      const a = start(g.start)
      const b = start(g.end) + length(g.end)
      if (visible(a, b)) {
        const from = Math.max(a + 2, head)
        const to = Math.min(b - 1, end)
        marks.push(
          <span
            key={`l${k}`}
            className="outline-line"
            style={
              rows
                ? { left: c, top: from, height: Math.max(0, to - from) }
                : { top: c, left: from, width: Math.max(0, to - from) }
            }
          />,
        )
        // The tick at the end away from the summary.
        const tickAt = sumAt !== null && sumAt < a ? b - 1 : a + 2
        if (tickAt >= head && tickAt <= end)
          marks.push(
            <span
              key={`t${k}`}
              className="outline-tick"
              style={rows ? { left: c, top: tickAt - 1 } : { top: c, left: tickAt - 1 }}
            />,
          )
      }
    }
    if (sumAt !== null && sumAt >= head && sumAt <= end) {
      const label = `${collapsed ? "Show" : "Hide"} detail of ${rows ? "rows" : "columns"} ${name(rows, g.start)} to ${name(rows, g.end)}`
      marks.push(
        <button
          key={`b${k}`}
          type="button"
          className="outline-toggle"
          aria-label={label}
          title={label}
          style={rows ? { left: c - 6, top: sumAt - 6 } : { top: c - 6, left: sumAt - 6 }}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            controller.outline.toggle(axis, g, collapsed)
            onDone()
          }}
        >
          {collapsed ? "+" : "−"}
        </button>,
      )
    }
  })

  return (
    <div
      className={`outline-bar is-${axis}`}
      style={rows ? { width: size } : { height: size }}
      role="group"
      aria-label={rows ? "Row outline" : "Column outline"}
    >
      <div className="outline-levels" style={rows ? { top: 0, height: head } : { left: 0, width: head }}>
        {Array.from({ length: depth + 1 }, (_, i) => (
          <button
            key={i}
            type="button"
            className="outline-level"
            aria-label={`Show ${rows ? "row" : "column"} level ${i + 1}`}
            title={`Level ${i + 1}`}
            style={rows ? { left: lane(i + 1) - 7 } : { top: lane(i + 1) - 7 }}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              controller.outline.showLevel(axis, i + 1)
              onDone()
            }}
          >
            {i + 1}
          </button>
        ))}
      </div>
      {marks}
    </div>
  )
}

const name = (rows: boolean, i: number) => (rows ? String(i + 1) : columnLetter(i))

function columnLetter(col: number): string {
  let s = ""
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}
