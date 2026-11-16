import { useState } from "react"
import { goalSeek, parseCell, type GoalSeekResult } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** A cell typed in a What-If field: "B4", "$B$4" or "Sheet1!B4" on the active sheet. */
function cellOnSheet(controller: SheetController, text: string) {
  const bare = text.trim().replace(/^=/, "")
  const bang = bare.lastIndexOf("!")
  if (bang >= 0) {
    const name = bare.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'")
    if (name.toLowerCase() !== controller.sheet.name.toLowerCase()) return null
  }
  const at = parseCell(bare.slice(bang + 1).replace(/\$/g, ""))
  return at ? { sheet: controller.sheetIndex, row: at.row, col: at.col } : null
}

/**
 * Data › What-If Analysis › Goal Seek: the value for one cell that makes a formula give the result you want. The
 * answer shows in Goal Seek Status first; OK keeps it (one Undo takes it back), Cancel leaves the sheet as it was.
 */
export function GoalSeekDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const active = controller.selection.active
  const [setText, setSetText] = useState(`${columnLetter(active.col)}${active.row + 1}`)
  const [target, setTarget] = useState("")
  const [changing, setChanging] = useState("")
  const [status, setStatus] = useState<{
    result: GoalSeekResult
    set: string
    changingAt: { row: number; col: number }
  } | null>(null)

  const run = () => {
    const set = cellOnSheet(controller, setText)
    const change = cellOnSheet(controller, changing)
    const goal = Number(target.trim().replace(",", "."))
    const cells = controller.sheet.cells
    if (!set) return controller.shell.alert("Reference isn't valid.")
    if (!cells.input(set.row, set.col)?.startsWith("=")) return controller.shell.alert("Cell must contain a formula.")
    if (!target.trim() || !Number.isFinite(goal)) return controller.shell.alert("Invalid value.")
    if (!change) return controller.shell.alert("Reference isn't valid.")
    if (cells.input(change.row, change.col)?.startsWith("="))
      return controller.shell.alert("Cell must contain a value.")
    const result = goalSeek(controller.wb, set, goal, change)
    setStatus({ result, set: setText.toUpperCase().replace(/\$/g, ""), changingAt: change })
  }

  if (status) {
    const { result } = status
    const keep = () => {
      const at = status.changingAt
      const data = { ...(controller.sheet.cells.get(at.row, at.col) ?? {}), value: result.value }
      delete data.input
      const sel = controller.selection
      controller.transact(
        [{ kind: "setCells", sheet: controller.sheetIndex, cells: [{ row: at.row, col: at.col, data }] }],
        {
          label: "Goal Seek",
          before: sel,
          after: sel,
        },
      )
      onClose()
    }
    return (
      <Dialog title="Goal Seek Status" onCancel={onClose} onAccept={keep} width={320}>
        <p className="dialog-label">
          Goal Seeking with Cell {status.set} {result.found ? "found a solution." : "may not have found a solution."}
        </p>
        <div className="goal-seek-status">
          <span>Target value:</span>
          <span>{target}</span>
          <span>Current value:</span>
          <span>{result.result === null ? "#VALUE!" : formatNumber(result.result)}</span>
        </div>
      </Dialog>
    )
  }

  return (
    <Dialog title="Goal Seek" onCancel={onClose} onAccept={run} width={300}>
      <div className="goal-seek-fields">
        <label className="dialog-label" htmlFor="goal-set">
          S<u>e</u>t cell:
        </label>
        <input id="goal-set" className="dialog-field" value={setText} onChange={(e) => setSetText(e.target.value)} />
        <label className="dialog-label" htmlFor="goal-to">
          To <u>v</u>alue:
        </label>
        <input id="goal-to" className="dialog-field" value={target} onChange={(e) => setTarget(e.target.value)} />
        <label className="dialog-label" htmlFor="goal-by">
          By <u>c</u>hanging cell:
        </label>
        <input id="goal-by" className="dialog-field" value={changing} onChange={(e) => setChanging(e.target.value)} />
      </div>
    </Dialog>
  )
}

const formatNumber = (n: number) =>
  Math.abs(n) >= 1e-4 || n === 0 ? String(Math.round(n * 1e6) / 1e6) : n.toExponential(4)

function columnLetter(col: number): string {
  let s = ""
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s
  return s
}
