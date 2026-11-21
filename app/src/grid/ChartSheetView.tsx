import { renderChart } from "@cells/chart"
import { isChart, type SheetChart } from "@cells/core"
import { useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"

/** Room around the chart, as Excel leaves a margin around a chart sheet's chart. */
const MARGIN = 16

/**
 * A chart sheet: its one chart filling the window (Excel's Zoom to Fit), on the workspace colour. The chart is always
 * selected, so the Chart Design and Format tabs show.
 */
export function ChartSheetView({ controller, active }: { controller: SheetController; active: boolean }) {
  const version = useSyncExternalStore(controller.subscribe, controller.getVersion)
  const host = useRef<HTMLDivElement>(null)
  const [size, setSize] = useState({ w: 0, h: 0 })
  useLayoutEffect(() => {
    const el = host.current
    if (!el) return
    const measure = () => setSize({ w: el.clientWidth, h: el.clientHeight })
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(el)
    return () => observer.disconnect()
  }, [])
  const chart = controller.sheet.objects.find(isChart) as SheetChart | undefined
  const w = Math.max(0, size.w - 2 * MARGIN)
  const h = Math.max(0, size.h - 2 * MARGIN)
  const svg = useMemo(() => {
    if (!chart || w < 40 || h < 40) return ""
    try {
      return renderChart(controller.charts.liveSpec(chart), w, h)
    } catch {
      return ""
    }
    // The version: the chart's data or look changed.
  }, [chart, w, h, version, controller])
  return (
    <div
      ref={host}
      className={`chart-sheet${active ? " is-active" : ""}`}
      role="img"
      aria-label={chart ? `${controller.sheet.name}: ${chart.altText || chart.name}` : controller.sheet.name}
      tabIndex={0}
    >
      {svg ? (
        <div className="chart-sheet-chart" style={{ width: w, height: h }}>
          {/* The renderer draws the chart's inside; the frame is ours, as on a worksheet. */}
          <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} aria-hidden dangerouslySetInnerHTML={{ __html: svg }} />
        </div>
      ) : null}
    </div>
  )
}
