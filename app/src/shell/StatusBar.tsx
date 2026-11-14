import { formatGeneral, formatValue, selectionStats } from "@cells/core"
import { useMemo, useSyncExternalStore } from "react"
import type { SheetController } from "@/sheet/controller"

const MODE_LABEL = { ready: "Ready", enter: "Enter", edit: "Edit" } as const

/**
 * Mode (Ready / Enter / Edit), End Mode, and Average, Count and Sum of a multi-cell selection, as Excel's status
 * bar shows them. Sheet tabs and the zoom slider arrive with the ribbon in Phase 3.
 */
export function StatusBar({ controller }: { controller: SheetController }) {
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
  return (
    <footer className="status-bar" aria-label="Status bar">
      <span>{MODE_LABEL[controller.mode]}</span>
      {controller.endMode ? <span>End Mode</span> : null}
      {stats && stats.count >= 2 ? (
        <div className="status-bar-figures">
          {stats.numericCount > 0 && stats.average !== null ? <span>Average: {show(stats.average)}</span> : null}
          <span>Count: {stats.count}</span>
          {stats.numericCount > 0 ? <span>Sum: {show(stats.sum)}</span> : null}
        </div>
      ) : null}
    </footer>
  )
}
