import type { SheetController } from "@/sheet/controller"
import { workbookStatistics, type SheetCounts } from "@/sheet/statistics"
import { Dialog } from "./Dialog"

const ROWS: [keyof SheetCounts, string][] = [
  ["cells", "Cells with data"],
  ["tables", "Tables"],
  ["formulas", "Formulas"],
  ["charts", "Charts"],
  ["images", "Images"],
  ["comments", "Comments"],
  ["notes", "Notes"],
]

/** Review › Workbook Statistics: counts for the sheet shown and for the whole workbook (Excel's dialog). */
export function WorkbookStatisticsDialog({
  controller,
  onClose,
}: {
  controller: SheetController
  onClose: () => void
}) {
  const stats = workbookStatistics(controller.wb, controller.sheetIndex)
  const n = (v: number) => v.toLocaleString()
  return (
    <Dialog title="Workbook Statistics" onCancel={onClose} onAccept={onClose} cancelLabel={null} width={340}>
      <dl className="stats">
        <dt className="stats-head">Current sheet: {stats.sheet.name}</dt>
        <dd />
        <dt>End of sheet</dt>
        <dd>{stats.sheet.end}</dd>
        {ROWS.map(([key, label]) => (
          <StatRow key={key} label={label} value={n(stats.sheet[key])} />
        ))}
        <dt className="stats-head">Workbook</dt>
        <dd />
        <dt>Sheets</dt>
        <dd>{n(stats.workbook.sheets)}</dd>
        {ROWS.map(([key, label]) => (
          <StatRow key={key} label={label} value={n(stats.workbook[key])} />
        ))}
      </dl>
    </Dialog>
  )
}

function StatRow({ label, value }: { label: string; value: string }) {
  return (
    <>
      <dt>{label}</dt>
      <dd>{value}</dd>
    </>
  )
}
