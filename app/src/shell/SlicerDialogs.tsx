import { useId, useState } from "react"
import type { SheetSlicer } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** Insert › Slicer: a box per table column; one slicer for each ticked. */
export function InsertSlicersDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const columns = controller.slicers.insertColumns()
  const [picked, setPicked] = useState<string[]>([])
  return (
    <Dialog
      title="Insert Slicers"
      width={280}
      onCancel={onClose}
      onAccept={() => {
        // In the table's order, as Excel adds them.
        controller.slicers.insert(columns.filter((c) => picked.includes(c)))
        onClose()
      }}
    >
      <div className="slicer-columns" role="group" aria-label="Table columns">
        {columns.map((name, i) => (
          <label key={name} className="dialog-radio">
            <input
              type="checkbox"
              data-autofocus={i === 0 || undefined}
              checked={picked.includes(name)}
              onChange={(e) => setPicked((p) => (e.target.checked ? [...p, name] : p.filter((x) => x !== name)))}
            />
            {name}
          </label>
        ))}
      </div>
    </Dialog>
  )
}

/** Slicer › Slicer Settings: name, caption and header, and how the items are sorted and shown. */
export function SlicerSettingsDialog({
  controller,
  slicer,
  onClose,
}: {
  controller: SheetController
  slicer: SheetSlicer
  onClose: () => void
}) {
  const id = useId()
  const [name, setName] = useState(slicer.name)
  const [caption, setCaption] = useState(slicer.caption)
  const [header, setHeader] = useState(slicer.header)
  const [sort, setSort] = useState(slicer.sort)
  const [customLists, setCustomLists] = useState(slicer.customLists)
  const [hide, setHide] = useState(Boolean(slicer.hideNoData))
  const [indicate, setIndicate] = useState(slicer.crossFilter !== "none")
  const [last, setLast] = useState(slicer.crossFilter !== "showItemsWithNoData")
  const accept = () => {
    const message = controller.slicers.applySettings(slicer, {
      name,
      caption,
      header,
      sort,
      customLists,
      hideNoData: hide,
      crossFilter: !indicate ? "none" : last ? "showItemsWithDataAtTop" : "showItemsWithNoData",
    })
    if (message) controller.shell.alert(message)
    else onClose()
  }
  return (
    <Dialog title="Slicer Settings" width={560} onCancel={onClose} onAccept={accept}>
      <p className="dialog-note">
        Source Name: {slicer.column}
        <br />
        Name to use in formulas: {slicer.cache ?? `Slicer_${slicer.column.replace(/[^\p{L}\p{N}_.]/gu, "_")}`}
      </p>
      <label className="spark-row" htmlFor={`${id}-n`}>
        <span>
          <u>N</u>ame:
        </span>
        <input
          id={`${id}-n`}
          className="dialog-field"
          value={name}
          data-autofocus
          onChange={(e) => setName(e.target.value)}
        />
      </label>
      <div className="slicer-settings">
        <fieldset className="tiw-fieldset">
          <legend>Header</legend>
          <label className="dialog-radio">
            <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
            <span>
              <u>D</u>isplay header
            </span>
          </label>
          <label className="spark-row" htmlFor={`${id}-c`}>
            <span>
              <u>C</u>aption:
            </span>
            <input
              id={`${id}-c`}
              className="dialog-field"
              value={caption}
              disabled={!header}
              onChange={(e) => setCaption(e.target.value)}
            />
          </label>
        </fieldset>
        <fieldset className="tiw-fieldset">
          <legend>Item Sorting and Filtering</legend>
          <label className="dialog-radio">
            <input
              type="radio"
              name={`${id}-sort`}
              checked={sort === "ascending"}
              onChange={() => setSort("ascending")}
            />
            <span>
              <u>A</u>scending (A to Z)
            </span>
          </label>
          <label className="dialog-radio">
            <input
              type="radio"
              name={`${id}-sort`}
              checked={sort === "descending"}
              onChange={() => setSort("descending")}
            />
            <span>
              D<u>e</u>scending (Z to A)
            </span>
          </label>
          <label className="dialog-radio">
            <input type="checkbox" checked={customLists} onChange={(e) => setCustomLists(e.target.checked)} />
            <span>
              Use Custom <u>L</u>ists when sorting
            </span>
          </label>
          <label className="dialog-radio">
            <input type="checkbox" checked={hide} onChange={(e) => setHide(e.target.checked)} />
            <span>
              <u>H</u>ide items with no data
            </span>
          </label>
          <label className="dialog-radio">
            <input type="checkbox" checked={indicate} disabled={hide} onChange={(e) => setIndicate(e.target.checked)} />
            <span>
              <u>V</u>isually indicate items with no data
            </span>
          </label>
          <label className="dialog-radio">
            <input
              type="checkbox"
              checked={last}
              disabled={hide || !indicate}
              onChange={(e) => setLast(e.target.checked)}
            />
            <span>
              <u>S</u>how items with no data last
            </span>
          </label>
        </fieldset>
      </div>
    </Dialog>
  )
}
