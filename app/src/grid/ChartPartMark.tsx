import { useLayoutEffect, useRef, useState } from "react"

type Box = { x: number; y: number; w: number; h: number }

/** Most boxes drawn for one part (a series of many points marks its first ones). */
const MOST = 80

/** The chart part under a pointer event: the nearest element tagged with one, or the chart area. */
export function chartPartAt(target: EventTarget | null): string {
  const element = target instanceof Element ? target.closest("[data-chart-part]") : null
  return element?.getAttribute("data-chart-part") ?? "chartArea"
}

/**
 * The picked part of a selected chart outlined, as Excel boxes the title, an axis or each bar of a series. Put inside
 * the chart's frame: it finds the part's elements in the frame's SVG and measures them after each draw.
 */
export function ChartPartMark({ part }: { part: string }) {
  const self = useRef<HTMLDivElement>(null)
  const [boxes, setBoxes] = useState<Box[]>([])
  useLayoutEffect(() => {
    const frame = self.current?.parentElement
    let next: Box[] = []
    if (frame && part !== "chartArea") {
      const origin = frame.getBoundingClientRect()
      const elements = [...frame.querySelectorAll(`svg [data-chart-part="${CSS.escape(part)}"]`)].slice(0, MOST)
      next = elements
        .map((e) => e.getBoundingClientRect())
        .filter((r) => r.width > 0 || r.height > 0)
        .map((r) => ({ x: r.left - origin.left, y: r.top - origin.top, w: r.width, h: r.height }))
    }
    // Measured after every draw; set only when it moved, or it would draw for ever.
    setBoxes((was) => (JSON.stringify(was) === JSON.stringify(next) ? was : next))
  })
  return (
    <div ref={self} className="chart-part-marks" aria-hidden>
      {boxes.map((b, i) => (
        <div
          key={i}
          className="chart-part-mark"
          style={{ left: b.x - 2, top: b.y - 2, width: b.w + 4, height: b.h + 4 }}
        />
      ))}
    </div>
  )
}
