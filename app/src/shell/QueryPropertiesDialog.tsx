import { allQueryTables, nameProblem } from "@cells/core"
import { useId, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { sourceName } from "@/sheet/queries"
import { Dialog } from "./Dialog"

/**
 * Connection Properties for a link to a text file (Data › Properties, or Properties in Queries & Connections): its
 * connection's name and description, the name of its data range, the file, and whether it refreshes when the
 * workbook opens.
 */
export function QueryPropertiesDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const target = controller.data.queryTarget
  const sheet = target ? controller.wb.sheets[target.sheet] : undefined
  const table = sheet?.queryTables.find((q) => q.name === target?.name)
  const [connection, setConnection] = useState(table?.connection ?? "")
  const [name, setName] = useState(table?.name ?? "")
  const [description, setDescription] = useState(table?.description ?? "")
  const [onLoad, setOnLoad] = useState(table?.refreshOnLoad ?? false)
  const [problem, setProblem] = useState<string | null>(null)
  const id = useId()
  if (!target || !sheet || !table) {
    return (
      <Dialog title="Connection Properties" onCancel={onClose} onAccept={onClose} cancelLabel={null}>
        <p>The connection is no longer in the workbook.</p>
      </Dialog>
    )
  }
  const others = allQueryTables(controller.wb.sheets).filter((q) => q.table !== table)
  const accept = () => {
    const conn = connection.trim()
    if (!conn) return setProblem("Enter a name for the connection.")
    if (others.some((q) => q.table.connection.toLowerCase() === conn.toLowerCase()))
      return setProblem("A connection with this name already exists. Enter a unique name.")
    const range = name.trim()
    const bad = nameProblem(
      range,
      sheet.queryTables.filter((q) => q !== table).map((q) => ({ name: q.name, ref: "" })),
    )
    if (bad) return setProblem(bad)
    controller.data.setQueryTable(target.sheet, table.name, {
      ...table,
      connection: conn,
      name: range,
      description,
      refreshOnLoad: onLoad,
    })
    onClose()
  }
  return (
    <Dialog title="Connection Properties" onCancel={onClose} onAccept={accept} width={420}>
      <div className="query-props">
        <label htmlFor={`${id}-conn`}>Connection name:</label>
        <input
          id={`${id}-conn`}
          className="dialog-field"
          value={connection}
          onChange={(e) => (setConnection(e.target.value), setProblem(null))}
        />
        <label htmlFor={`${id}-desc`}>Description:</label>
        <textarea
          id={`${id}-desc`}
          className="dialog-field"
          rows={2}
          value={description}
          onChange={(e) => setDescription(e.target.value)}
        />
        <label htmlFor={`${id}-name`}>Data range name:</label>
        <input
          id={`${id}-name`}
          className="dialog-field"
          value={name}
          onChange={(e) => (setName(e.target.value), setProblem(null))}
        />
        <span>File:</span>
        <span className="query-props-file" title={table.source}>
          {sourceName(table.source)}
        </span>
      </div>
      <fieldset className="dialog-group">
        <legend>Refresh control</legend>
        <label className="dialog-radio">
          <input type="checkbox" checked={onLoad} onChange={(e) => setOnLoad(e.target.checked)} />
          Refresh data when opening the file
        </label>
      </fieldset>
      {problem ? (
        <p className="dialog-hint query-props-problem" role="alert">
          {problem}
        </p>
      ) : null}
    </Dialog>
  )
}
