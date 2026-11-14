import { useState } from "react"
import { parseRange, rangeName, type Range } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import type { CreateTableInit } from "@/sheet/tables"
import { Dialog } from "./Dialog"

export const BAD_REFERENCE = "The reference isn't valid."

/** "=$A$1:$C$4", as Excel's range boxes show a range. */
export function absoluteRange(r: Range): string {
  return "=" + rangeName(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")
}

/** A range typed in the box ("=$A$1:$C$4", "A1:C4"), or null. */
export function typedRange(text: string): Range | null {
  return parseRange(text.trim().replace(/^=/, "").replace(/\$/g, ""))
}

/** Insert › Table (Ctrl+T): where the data is, and whether its first row holds the headers. */
export function CreateTableDialog({
  controller,
  init,
  onClose,
}: {
  controller: SheetController
  init: CreateTableInit
  onClose: () => void
}) {
  const [text, setText] = useState(absoluteRange(init.range))
  const [header, setHeader] = useState(init.header)
  const accept = () => {
    const range = typedRange(text)
    if (!range) return controller.shell.alert(BAD_REFERENCE)
    if (controller.tables.create(range, header, init.style)) onClose()
  }
  return (
    <Dialog title="Create Table" onCancel={onClose} onAccept={accept} width={300}>
      <label className="dialog-label" htmlFor="ct-range">
        <u>W</u>here is the data for your table?
      </label>
      <input
        id="ct-range"
        className="dialog-field"
        value={text}
        spellCheck={false}
        onChange={(e) => setText(e.target.value)}
      />
      <label className="fc-check ct-check">
        <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
        <span>
          <u>M</u>y table has headers
        </span>
      </label>
    </Dialog>
  )
}

/** Table Design › Resize Table. */
export function ResizeTableDialog({
  controller,
  range,
  onClose,
}: {
  controller: SheetController
  range: Range
  onClose: () => void
}) {
  const [text, setText] = useState(absoluteRange(range))
  const accept = () => {
    const next = typedRange(text)
    if (!next) return controller.shell.alert(BAD_REFERENCE)
    if (controller.tables.resize(next)) onClose()
  }
  return (
    <Dialog title="Resize Table" onCancel={onClose} onAccept={accept} width={320}>
      <label className="dialog-label" htmlFor="rt-range">
        <u>S</u>elect the new data range for your table:
      </label>
      <input
        id="rt-range"
        className="dialog-field"
        value={text}
        spellCheck={false}
        onChange={(e) => setText(e.target.value)}
      />
      <p className="rt-note">
        Note: The headers must remain in the same row, and the resulting table range must overlap the original table
        range.
      </p>
    </Dialog>
  )
}
