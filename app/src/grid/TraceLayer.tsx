import { useId } from "react"
import type { SheetController } from "@/sheet/controller"
import { colX, rowY, type GridLayout } from "./layout"

/**
 * Trace Precedents / Dependents' arrows over the sheet, drawn as Excel draws them: a blue line from a dot in the
 * precedent (a blue box around a precedent range) to an arrowhead in the formula's cell; red when the value it carries
 * is an error; a dashed black line from a sheet icon when the value comes from another sheet. They don't take clicks.
 */
export function TraceLayer({
  controller,
  layout,
  zoom,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
}) {
  const id = useId().replace(/:/g, "")
  const arrows = controller.audit.arrows()
  if (!arrows.length) return null
  const sheet = controller.sheet
  const center = (row: number, col: number) => ({
    x: colX(sheet, layout, col) + sheet.cols.size(col) / 2,
    y: rowY(sheet, layout, row) + sheet.rows.size(row) / 2,
  })
  const { x0, y0, x1, y1 } = {
    x0: layout.rowHeaderWidth,
    y0: layout.headerHeight,
    x1: layout.width,
    y1: layout.height,
  }
  return (
    <svg
      className="trace-layer"
      width={layout.width * zoom}
      height={layout.height * zoom}
      viewBox={`0 0 ${layout.width} ${layout.height}`}
      aria-hidden
    >
      <defs>
        <clipPath id={`${id}-clip`}>
          <rect x={x0} y={y0} width={Math.max(0, x1 - x0)} height={Math.max(0, y1 - y0)} />
        </clipPath>
        {(["blue", "red", "black"] as const).map((tone) => (
          <marker
            key={tone}
            id={`${id}-${tone}`}
            viewBox="0 0 10 10"
            refX="9"
            refY="5"
            markerWidth="7"
            markerHeight="7"
            orient="auto-start-reverse"
          >
            <path d="M0,0 L10,5 L0,10 z" className={`trace-head is-${tone}`} />
          </marker>
        ))}
      </defs>
      <g clipPath={`url(#${id}-clip)`}>
        {arrows.map((a, i) => {
          const to = center(a.to.row, a.to.col)
          if (!a.from) {
            // From another sheet: a sheet icon up and to the left of the formula's cell.
            const ix = to.x - 46
            const iy = to.y - 30
            return (
              <g key={i} className="trace is-black">
                <line
                  x1={ix + 7}
                  y1={iy + 6}
                  x2={to.x}
                  y2={to.y}
                  strokeDasharray="4 3"
                  markerEnd={`url(#${id}-black)`}
                />
                <rect className="trace-sheet" x={ix} y={iy} width={14} height={12} />
                <path className="trace-sheet-lines" d={`M${ix},${iy + 4}h14M${ix},${iy + 8}h14M${ix + 5},${iy}v12`} />
              </g>
            )
          }
          const tone = a.error ? "red" : "blue"
          const range = a.from.r0 !== a.from.r1 || a.from.c0 !== a.from.c1
          const from = center(a.from.r0, a.from.c0)
          const box = range
            ? {
                x: colX(sheet, layout, a.from.c0),
                y: rowY(sheet, layout, a.from.r0),
                w: colX(sheet, layout, a.from.c1) + sheet.cols.size(a.from.c1) - colX(sheet, layout, a.from.c0),
                h: rowY(sheet, layout, a.from.r1) + sheet.rows.size(a.from.r1) - rowY(sheet, layout, a.from.r0),
              }
            : null
          return (
            <g key={i} className={`trace is-${tone}`}>
              {box ? (
                <rect className="trace-box" x={box.x + 1} y={box.y + 1} width={box.w - 2} height={box.h - 2} />
              ) : null}
              <line x1={from.x} y1={from.y} x2={to.x} y2={to.y} markerEnd={`url(#${id}-${tone})`} />
              <circle className="trace-dot" cx={from.x} cy={from.y} r={2.5} />
            </g>
          )
        })}
      </g>
    </svg>
  )
}
