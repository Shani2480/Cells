import { renderChart } from "@cells/chart"
import type { SheetChart } from "@cells/core"
import { memo, useEffect, useMemo, useRef, useState } from "react"
import { pngKey, preparePng, readyPng } from "./chart-png"
import { chartBox, type ChartBox } from "@/sheet/charts"
import type { SheetController } from "@/sheet/controller"
import { colX, rowY, type GridLayout } from "./layout"

/** Excel's chart handles: the four corners and the middle of each side. */
const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const
type Handle = (typeof HANDLES)[number]

/** A drag starts after this many pixels, so a click only selects. */
const DRAG_START = 4
/** The smallest chart a drag makes (pixels at 100%). */
const MIN_SIZE = 20

/** One chart drawn: its SVG at its size in sheet pixels, scaled by the view. Redrawn only when it or its data change. */
const ChartPicture = memo(function ChartPicture({ svg, w, h }: { svg: string; w: number; h: number }) {
  return (
    <svg
      className="chart-svg"
      width="100%"
      height="100%"
      viewBox={`0 0 ${w} ${h}`}
      preserveAspectRatio="none"
      aria-hidden
      dangerouslySetInnerHTML={{ __html: svg }}
    />
  )
})

function resized(box: ChartBox, handle: Handle, dx: number, dy: number, keepRatio: boolean): ChartBox {
  let { x, y, w, h } = box
  if (handle.includes("w")) {
    x += dx
    w -= dx
  }
  if (handle.includes("e")) w += dx
  if (handle.includes("n")) {
    y += dy
    h -= dy
  }
  if (handle.includes("s")) h += dy
  // Shift on a corner keeps the chart's proportions, as in Excel.
  if (keepRatio && handle.length === 2) {
    const ratio = box.w / box.h
    if (w / h > ratio) w = h * ratio
    else h = w / ratio
    if (handle.includes("w")) x = box.x + box.w - w
    if (handle.includes("n")) y = box.y + box.h - h
  }
  if (w < MIN_SIZE) {
    if (handle.includes("w")) x -= MIN_SIZE - w
    w = MIN_SIZE
  }
  if (h < MIN_SIZE) {
    if (handle.includes("n")) y -= MIN_SIZE - h
    h = MIN_SIZE
  }
  return { x: Math.max(0, x), y: Math.max(0, y), w, h }
}

/**
 * The sheet's charts over the cells: each drawn from its spec and its ranges, picked with a click, moved by dragging
 * and sized by its handles (Shift keeps the proportions). The cells under a chart can't be clicked through it.
 */
export function ChartLayer({
  controller,
  layout,
  zoom,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  onDone: () => void
}) {
  const sheet = controller.sheet
  const version = controller.version
  const [drag, setDrag] = useState<{ id: number; box: ChartBox } | null>(null)
  const start = useRef<{
    id: number
    handle: Handle | "move"
    x: number
    y: number
    box: ChartBox
    moved: boolean
  } | null>(null)

  // The charts' pictures, drawn again when the workbook changes (their data may have).
  const pictures = useMemo(() => {
    const out = new Map<number, { svg: string; box: ChartBox }>()
    for (const chart of sheet.charts) {
      const box = chartBox(sheet, chart.anchor)
      out.set(chart.id, { svg: picture(controller, chart, box), box })
    }
    return out
  }, [sheet, version, controller])

  // The selected chart's PNG, ready for Copy (which can't wait for one to be drawn).
  const selectedId = controller.charts.selected
  const selectedPicture = selectedId === null ? null : pictures.get(selectedId)
  useEffect(() => {
    controller.charts.picture = (chart, _spec, size) => {
      const svg = chart.id === selectedId && selectedPicture ? selectedPicture.svg : null
      return svg ? readyPng(chart.id, pngKey(svg, size.w, size.h)) : null
    }
    if (selectedId !== null && selectedPicture)
      void preparePng(selectedId, selectedPicture.svg, selectedPicture.box.w, selectedPicture.box.h)
  }, [controller, selectedId, selectedPicture])

  if (!sheet.charts.length) return null
  const left = layout.rowHeaderWidth
  const top = layout.headerHeight
  // Placed from the cell it starts in, so a chart in frozen rows or columns stays put as the rest scrolls.
  const toScreen = (chart: SheetChart, box: ChartBox) => {
    const { from } = chart.anchor
    return {
      left: (colX(sheet, layout, from.col) + box.x - sheet.cols.offset(from.col) - left) * zoom,
      top: (rowY(sheet, layout, from.row) + box.y - sheet.rows.offset(from.row) - top) * zoom,
      width: box.w * zoom,
      height: box.h * zoom,
    }
  }

  const begin = (event: React.PointerEvent, chart: SheetChart, handle: Handle | "move") => {
    if (event.button !== 0) return
    event.stopPropagation()
    event.preventDefault()
    controller.charts.select(chart.id)
    const box = pictures.get(chart.id)!.box
    start.current = { id: chart.id, handle, x: event.clientX, y: event.clientY, box, moved: false }
    ;(event.currentTarget as HTMLElement).setPointerCapture(event.pointerId)
  }
  const move = (event: React.PointerEvent) => {
    const s = start.current
    if (!s) return
    const dx = (event.clientX - s.x) / zoom
    const dy = (event.clientY - s.y) / zoom
    if (!s.moved && Math.hypot(dx, dy) < DRAG_START / zoom) return
    s.moved = true
    const box =
      s.handle === "move"
        ? { ...s.box, x: Math.max(0, s.box.x + dx), y: Math.max(0, s.box.y + dy) }
        : resized(s.box, s.handle, dx, dy, event.shiftKey)
    setDrag({ id: s.id, box })
  }
  const end = () => {
    const s = start.current
    start.current = null
    if (s?.moved && drag) controller.charts.place(s.id, drag.box)
    setDrag(null)
    onDone()
  }

  return (
    <div className="chart-layer" style={{ left: left * zoom, top: top * zoom }}>
      {sheet.charts.map((chart) => {
        const p = pictures.get(chart.id)
        if (!p) return null
        const selected = controller.charts.selected === chart.id
        const dragging = drag?.id === chart.id ? drag.box : null
        return (
          <div
            key={chart.id}
            className={`chart-frame${selected ? " is-selected" : ""}`}
            style={toScreen(chart, p.box)}
            role="img"
            aria-label={`${chart.name}`}
            data-chart={chart.id}
            onPointerDown={(event) => begin(event, chart, "move")}
            onPointerMove={move}
            onPointerUp={end}
            onPointerCancel={() => {
              start.current = null
              setDrag(null)
            }}
          >
            <ChartPicture svg={p.svg} w={p.box.w} h={p.box.h} />
            {selected
              ? HANDLES.map((h) => (
                  <span
                    key={h}
                    className={`chart-handle is-${h}`}
                    onPointerDown={(event) => begin(event, chart, h)}
                    onPointerMove={move}
                    onPointerUp={end}
                  />
                ))
              : null}
            {dragging ? (
              <div
                className="chart-ghost"
                style={{
                  left: (dragging.x - p.box.x) * zoom,
                  top: (dragging.y - p.box.y) * zoom,
                  width: dragging.w * zoom,
                  height: dragging.h * zoom,
                }}
              />
            ) : null}
          </div>
        )
      })}
    </div>
  )
}

function picture(controller: SheetController, chart: SheetChart, box: ChartBox): string {
  try {
    return renderChart(controller.charts.liveSpec(chart), box.w, box.h)
  } catch {
    // A chart the renderer can't draw shows as an empty frame rather than taking the sheet down.
    return ""
  }
}
