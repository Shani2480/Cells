import { cellName, formatGeneral, formatValue, selectionStats } from "@cells/core"
import { useMemo, useSyncExternalStore } from "react"
import { FileText, Minus, Plus, Rows3, Table2 } from "lucide-react"
import { MAX_ZOOM, MIN_ZOOM, type SheetController, type SheetViewMode } from "@/sheet/controller"

/** Page Layout view isn't there yet: its button stays greyed out. */
const PAGE_LAYOUT_BUILT = false

const MODE_LABEL = { ready: "Ready", enter: "Enter", edit: "Edit", point: "Point" } as const

/**
 * Mode (Ready / Enter / Edit), End Mode, and Average, Count and Sum of a multi-cell selection, as Excel's status
 * bar shows them. The zoom slider comes with View › Zoom.
 */
/** Excel's slider: 10%–100% on the left half, 100%–400% on the right (0–1000 on the input). */
export function sliderOf(zoom: number): number {
  return zoom <= 1 ? ((zoom - MIN_ZOOM) / (1 - MIN_ZOOM)) * 500 : 500 + ((zoom - 1) / (MAX_ZOOM - 1)) * 500
}

export function zoomOfSlider(value: number): number {
  const z = value <= 500 ? MIN_ZOOM + (value / 500) * (1 - MIN_ZOOM) : 1 + ((value - 500) / 500) * (MAX_ZOOM - 1)
  return Math.round(z * 100) / 100
}

/** The minus and plus buttons step by 10% (to the next multiple of ten). */
export function stepZoom(zoom: number, dir: 1 | -1): number {
  const pct = Math.round(zoom * 100)
  const next = dir > 0 ? Math.floor(pct / 10) * 10 + 10 : Math.ceil(pct / 10) * 10 - 10
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, next / 100))
}

export function StatusBar({ controller, onZoomDialog }: { controller: SheetController; onZoomDialog: () => void }) {
  const version = useSyncExternalStore(controller.subscribe, controller.getVersion)
  const selection = controller.selection
  const stats = useMemo(() => {
    const single =
      selection.ranges.length === 1 &&
      selection.ranges[0].r0 === selection.ranges[0].r1 &&
      selection.ranges[0].c0 === selection.ranges[0].c1
    return single ? null : selectionStats(controller.sheet, selection.ranges)
  }, [controller, selection, version])
  // Excel shows the figures in the active cell's number format; in General that is about 11 characters
  // ("44.56333333"), like a standard-width cell.
  const { active } = selection
  const numberFormat = controller.wb.styles.get(controller.sheet.cells.style(active.row, active.col)).numberFormat
  const show = (value: number) => formatValue(value, numberFormat, 11) ?? formatGeneral(value, 16) ?? String(value)
  const calc = controller.wb.calcStatus
  const records = controller.data.records()
  // Excel names the cell when the circular reference is on the sheet in view.
  const circular = calc.circular
    ? calc.circular.sheet === controller.sheetIndex
      ? `Circular References: ${cellName(calc.circular)}`
      : "Circular References"
    : null
  return (
    <footer className="status-bar" aria-label="Status bar">
      <span>{MODE_LABEL[controller.mode]}</span>
      {records && controller.mode === "ready" ? (
        <span>
          {records.shown} of {records.total} records found
        </span>
      ) : null}
      {controller.endMode ? <span>End Mode</span> : null}
      {circular ? <span>{circular}</span> : null}
      {calc.mode === "manual" && calc.pending ? <span>Calculate</span> : null}
      <span className="status-bar-spacer" />
      {stats && stats.count >= 2 ? (
        <div className="status-bar-figures">
          {stats.numericCount > 0 && stats.average !== null ? <span>Average: {show(stats.average)}</span> : null}
          <span>Count: {stats.count}</span>
          {stats.numericCount > 0 ? <span>Sum: {show(stats.sum)}</span> : null}
        </div>
      ) : null}
      <StatusViews controller={controller} />
      <StatusZoom controller={controller} onZoomDialog={onZoomDialog} />
    </footer>
  )
}

const VIEWS: { view: SheetViewMode; label: string; icon: React.ReactNode }[] = [
  { view: "normal", label: "Normal", icon: <Table2 size={14} strokeWidth={1.6} /> },
  { view: "pageLayout", label: "Page Layout", icon: <FileText size={14} strokeWidth={1.6} /> },
  { view: "pageBreakPreview", label: "Page Break Preview", icon: <Rows3 size={14} strokeWidth={1.6} /> },
]

/** The view buttons left of the zoom (Excel's Normal, Page Layout and Page Break Preview). */
function StatusViews({ controller }: { controller: SheetController }) {
  return (
    <div className="status-views" role="group" aria-label="Views">
      {VIEWS.map((v) => {
        const built = v.view !== "pageLayout" || PAGE_LAYOUT_BUILT
        return (
          <button
            key={v.view}
            type="button"
            className="status-view"
            aria-label={v.label}
            title={built ? v.label : `${v.label}: not built yet.`}
            aria-pressed={controller.view === v.view}
            disabled={!built}
            onClick={() => controller.setView(v.view)}
          >
            {v.icon}
          </button>
        )
      })}
    </div>
  )
}

/** The zoom controls also follow a pinch, which zooms without re-rendering the rest of the window. */
function StatusZoom({ controller, onZoomDialog }: { controller: SheetController; onZoomDialog: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  useSyncExternalStore(controller.subscribeScroll, controller.getScrollVersion)
  return (
    <div className="status-zoom">
      <button
        type="button"
        className="status-zoom-step"
        aria-label="Zoom Out"
        title="Zoom Out"
        onClick={() => controller.setZoom(stepZoom(controller.zoom, -1))}
      >
        <Minus size={13} strokeWidth={1.75} />
      </button>
      <input
        type="range"
        className="status-zoom-slider"
        aria-label="Zoom"
        min={0}
        max={1000}
        value={Math.round(sliderOf(controller.zoom))}
        onChange={(e) => {
          const v = Number(e.target.value)
          // The slider snaps to 100% near its middle, as Excel's.
          controller.setZoom(Math.abs(v - 500) < 12 ? 1 : zoomOfSlider(v))
        }}
      />
      <button
        type="button"
        className="status-zoom-step"
        aria-label="Zoom In"
        title="Zoom In"
        onClick={() => controller.setZoom(stepZoom(controller.zoom, 1))}
      >
        <Plus size={13} strokeWidth={1.75} />
      </button>
      <button
        type="button"
        className="status-zoom-value"
        title="Zoom level. Click to open the Zoom dialog."
        onClick={onZoomDialog}
      >
        {Math.round(controller.zoom * 100)}%
      </button>
    </div>
  )
}
