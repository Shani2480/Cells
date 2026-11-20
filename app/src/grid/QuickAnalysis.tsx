import { useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent } from "react"
import { createPortal } from "react-dom"
import { renderChart } from "@cells/chart"
import type { SheetController } from "@/sheet/controller"
import { QUICK_TABS, type QuickItem, type QuickShape, type QuickTab } from "@/sheet/quick-analysis"
import { rangeRect, type GridLayout } from "./layout"
import { PivotPreview } from "@/shell/RecommendedPivotsDialog"
import { QuickButtonIcon, QuickIcon } from "./quick-icons"

/** Excel's button: 22 px at 100%, just outside the selection's bottom-right corner. */
const BUTTON = 22
const GAP = 3

/**
 * Quick Analysis: the button at the bottom-right corner of a selection with data, and the gallery it opens (also
 * Ctrl+Q and the cell menu's Quick Analysis).
 */
export function QuickAnalysisLayer({
  controller,
  layout,
  zoom,
  dragging,
  onDone,
}: {
  controller: SheetController
  layout: GridLayout
  zoom: number
  /** A selection being dragged: the button waits for the drag to end, as Excel's. */
  dragging: boolean
  onDone: () => void
}) {
  const [open, setOpen] = useState<{ left: number; top: number } | null>(null)
  const quick = controller.quick
  const shown = !dragging && !controller.touchHandles && quick.available()

  const place = () => {
    const sel = controller.selection.ranges[0]
    if (!sel) return null
    const rect = rangeRect(controller, layout, sel)
    const x = rect.x + rect.w + GAP / zoom
    const y = rect.y + rect.h + GAP / zoom
    const inside =
      x >= layout.rowHeaderWidth && y >= layout.headerHeight && x < layout.width - 4 && y < layout.height - 4
    return { x, y, inside }
  }
  const at = shown ? place() : null

  // Ctrl+Q and the cell menu open the gallery under the button (or at the visible part of the selection).
  useEffect(() => {
    quick.openGallery = () => {
      const host = document.querySelector<HTMLElement>("[data-testid=grid]")?.getBoundingClientRect()
      const p = place()
      if (!host || !p) return
      const x = Math.min(Math.max(p.x, layout.rowHeaderWidth), layout.width - 40)
      const y = Math.min(Math.max(p.y, layout.headerHeight), layout.height - 40)
      setOpen({ left: host.left + x * zoom, top: host.top + (y + BUTTON / zoom) * zoom + 2 })
    }
    return () => {
      quick.openGallery = null
    }
  })

  const close = (refocus = true) => {
    quick.endPreview()
    setOpen(null)
    if (refocus) onDone()
  }

  return (
    <>
      {at?.inside ? (
        <button
          type="button"
          className={`quick-button${open ? " is-open" : ""}`}
          style={{ left: at.x * zoom, top: at.y * zoom, width: BUTTON, height: BUTTON }}
          aria-label="Quick Analysis"
          aria-haspopup="dialog"
          title="Quick Analysis (Ctrl+Q)"
          onPointerDown={(event) => event.stopPropagation()}
          onClick={(event) => {
            if (open) return close()
            const r = event.currentTarget.getBoundingClientRect()
            setOpen({ left: r.left, top: r.bottom + 2 })
          }}
        >
          <QuickButtonIcon size={16} />
        </button>
      ) : null}
      {open ? <QuickGallery controller={controller} left={open.left} top={open.top} onClose={close} /> : null}
    </>
  )
}

/** The gallery: five tabs of choices; hovering or focusing one previews it, a click makes it. */
export function QuickGallery({
  controller,
  left,
  top,
  onClose,
}: {
  controller: SheetController
  left: number
  top: number
  onClose: (refocus?: boolean) => void
}) {
  const quick = controller.quick
  const ref = useRef<HTMLDivElement>(null)
  const [tab, setTab] = useState<QuickTab>("formatting")
  const [hover, setHover] = useState<string | null>(null)
  const [pos, setPos] = useState<{ left: number; top: number; visible: boolean }>({ left, top, visible: false })
  // The data is read once, as the gallery opens (previews change the cells under it).
  const shape = useMemo(() => quick.shape(), [quick])
  const items = useMemo(() => (shape ? itemsFor(controller, shape) : null), [controller, shape])

  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const r = el.getBoundingClientRect()
    const x = Math.max(4, Math.min(left, window.innerWidth - r.width - 4))
    // Not enough room below: above the point instead.
    const y = top + r.height > window.innerHeight - 4 ? Math.max(4, top - r.height - 30) : top
    setPos({ left: x, top: y, visible: true })
    el.querySelector<HTMLButtonElement>(".qa-tab[aria-selected=true]")?.focus({ preventScroll: true })
  }, [left, top])

  useEffect(() => {
    const onDown = (event: PointerEvent) => {
      const target = event.target as Node
      if (!ref.current?.contains(target)) onClose(false)
    }
    const onKey = (event: globalThis.KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault()
        event.stopPropagation()
        onClose()
      }
    }
    window.addEventListener("pointerdown", onDown, true)
    window.addEventListener("keydown", onKey, true)
    return () => {
      window.removeEventListener("pointerdown", onDown, true)
      window.removeEventListener("keydown", onKey, true)
    }
  }, [onClose])

  // A preview never outlives the gallery.
  useEffect(() => () => quick.endPreview(), [quick])

  const enter = (item: QuickItem) => {
    setHover(item.id)
    if (!item.disabled) quick.preview(item.id)
  }
  const leave = (item: QuickItem) => {
    setHover((h) => (h === item.id ? null : h))
    quick.endPreview()
  }
  const pick = (item: QuickItem) => {
    if (item.disabled) return
    onClose()
    quick.apply(item.id)
  }

  const tabs = QUICK_TABS
  const list = items?.[tab] ?? []
  const info = tabs.find((t) => t.id === tab)!
  const hovered = list.find((i) => i.id === hover)
  const chart = hover?.startsWith("chart:") ? quick.chartSpec(hover.slice(6)) : null
  const pivot = hover?.startsWith("pivot:") && shape ? quick.pivots(shape)?.items[Number(hover.slice(6))] : undefined

  // Left and Right move along the tabs or the choices; Down goes to the choices, Up back to the tabs.
  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const el = ref.current
    if (!el) return
    const target = event.target as HTMLElement
    const onTab = target.classList.contains("qa-tab")
    const group = [...el.querySelectorAll<HTMLButtonElement>(onTab ? ".qa-tab" : ".qa-item:not([aria-disabled=true])")]
    const i = group.indexOf(target as HTMLButtonElement)
    if (event.key === "ArrowRight" || event.key === "ArrowLeft") {
      event.preventDefault()
      const next = group[(i + (event.key === "ArrowRight" ? 1 : group.length - 1)) % group.length]
      next?.focus()
      if (onTab && next) setTab(next.dataset.tab as QuickTab)
    } else if (event.key === "ArrowDown" && onTab) {
      event.preventDefault()
      el.querySelector<HTMLButtonElement>(".qa-item:not([aria-disabled=true])")?.focus()
    } else if (event.key === "ArrowUp" && !onTab) {
      event.preventDefault()
      el.querySelector<HTMLButtonElement>(".qa-tab[aria-selected=true]")?.focus()
    }
  }

  return createPortal(
    <div
      ref={ref}
      className="quick-gallery"
      role="dialog"
      aria-label="Quick Analysis"
      style={{ left: pos.left, top: pos.top, visibility: pos.visible ? "visible" : "hidden" }}
      onKeyDown={onKeyDown}
    >
      {chart ? (
        <div className={`qa-chart${pos.top < 214 ? " is-below" : ""}`} data-keep-light aria-hidden>
          <ChartPicture spec={chart} />
        </div>
      ) : null}
      {pivot ? (
        <div className={`qa-chart qa-pivot${pos.top < 214 ? " is-below" : ""}`} aria-hidden>
          <div className="recommend-preview-title">{pivot.recipe.title}</div>
          <PivotPreview table={pivot.table} cache={pivot.cache} rows={12} />
        </div>
      ) : null}
      <div className="qa-tabs" role="tablist">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            className="qa-tab"
            data-tab={t.id}
            aria-selected={t.id === tab}
            tabIndex={t.id === tab ? 0 : -1}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className={`qa-items${tab === "totals" ? " is-totals" : ""}`} role="tabpanel" aria-label={info.label}>
        {list.length ? (
          list.map((item) => (
            <button
              key={item.id}
              type="button"
              className="qa-item"
              aria-disabled={item.disabled ? true : undefined}
              title={item.disabled ?? undefined}
              onMouseEnter={() => enter(item)}
              onMouseLeave={() => leave(item)}
              onFocus={() => enter(item)}
              onBlur={() => leave(item)}
              onClick={() => pick(item)}
            >
              <QuickIcon id={item.id} />
              <span>{item.label}</span>
            </button>
          ))
        ) : (
          <div className="qa-empty">{EMPTY[tab]}</div>
        )}
      </div>
      <div className="qa-text">{hovered && !hovered.disabled ? hovered.tip : (hovered?.disabled ?? info.text)}</div>
    </div>,
    document.body,
  )
}

const EMPTY: Record<QuickTab, string> = {
  formatting: "No formatting for this data.",
  charts: "We didn't find charts to recommend for this data.",
  totals: "No totals for this data.",
  tables: "No tables for this data.",
  sparklines: "No sparklines for this data.",
}

function itemsFor(controller: SheetController, shape: QuickShape): Record<QuickTab, QuickItem[]> {
  const quick = controller.quick
  const charts = quick.charts(shape)
  return {
    formatting: quick.formatting(shape),
    charts: charts.length
      ? [
          ...charts.map((id) => ({ id: `chart:${id}`, label: chartLabel(id), tip: CHART_TIP[id] ?? chartLabel(id) })),
          { id: "more", label: "More Charts", tip: "Click for more charts." },
        ]
      : [],
    totals: quick.totals(shape),
    tables: quick.tables(shape),
    sparklines: quick.sparklines(shape),
  }
}

const CHART_LABEL: Record<string, string> = {
  column: "Clustered Column",
  "column-stacked": "Stacked Column",
  hbar: "Clustered Bar",
  line: "Line",
  pie: "Pie",
  area: "Area",
  "area-stacked": "Stacked Area",
  scatter: "Scatter",
}
const chartLabel = (id: string) => CHART_LABEL[id] ?? id

const CHART_TIP: Record<string, string> = {
  column: "Clustered column charts compare values across categories.",
  "column-stacked": "Stacked column charts show how parts add up to a whole.",
  hbar: "Clustered bar charts compare values across categories.",
  line: "Line charts show trends over time.",
  pie: "Pie charts show proportions of a whole.",
  area: "Area charts show trends and the size of change.",
  "area-stacked": "Stacked area charts show how parts change over time.",
  scatter: "Scatter charts show the relationship between sets of values.",
}

/** The chart the pointer is over, drawn at a readable size above the gallery (Excel's chart preview). */
function ChartPicture({ spec }: { spec: Parameters<typeof renderChart>[0] }) {
  const w = 320
  const h = 200
  let svg: string
  try {
    svg = renderChart(spec, w, h)
  } catch {
    svg = ""
  }
  return <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} dangerouslySetInnerHTML={{ __html: svg }} />
}
