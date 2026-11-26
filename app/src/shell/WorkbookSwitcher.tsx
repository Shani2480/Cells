import { useState } from "react"
import { Check, ChevronDown, FilePlus2, FileSpreadsheet, FolderOpen, X } from "lucide-react"
import type { FileCommands } from "@/file/commands"
import { BottomSheet } from "@/phone/BottomSheet"
import { Popover } from "@/ui/popover"

/**
 * The workbooks open in this window (Android and a browser keep several in one window): each with its name and an
 * unsaved-changes dot, the one shown ticked, a Close button on each; New Workbook and Open below. Choosing one shows it.
 */
export function WorkbookList({ files, onDone }: { files: FileCommands; onDone: () => void }) {
  const run = (action: () => Promise<boolean> | void) => {
    onDone()
    void action()
  }
  return (
    <div className="wb-switch" role="menu" aria-label="Open workbooks">
      {files.docs.map((doc) => {
        const current = doc === files.doc
        return (
          <div key={doc.id} className={`wb-switch-row${current ? " is-current" : ""}`}>
            <button
              type="button"
              role="menuitemradio"
              aria-checked={current}
              className="wb-switch-item"
              onClick={() => run(() => files.switchTo(doc))}
            >
              <span className="wb-switch-check" aria-hidden>
                {current ? <Check size={14} /> : null}
              </span>
              <FileSpreadsheet size={18} strokeWidth={1.6} aria-hidden className="wb-switch-icon" />
              <span className="wb-switch-name">{doc.title}</span>
              {doc.dirty ? (
                <span className="wb-switch-dirty" title="Unsaved changes" aria-label="Unsaved changes" />
              ) : null}
            </button>
            <button
              type="button"
              className="wb-switch-close"
              aria-label={`Close ${doc.title}`}
              title="Close"
              onClick={() => run(() => files.closeDocument(doc))}
            >
              <X size={14} aria-hidden />
            </button>
          </div>
        )
      })}
      <div className="wb-switch-sep" role="separator" />
      <button type="button" role="menuitem" className="wb-switch-action" onClick={() => run(() => files.newWorkbook())}>
        <FilePlus2 size={16} strokeWidth={1.6} aria-hidden />
        New Workbook
      </button>
      <button type="button" role="menuitem" className="wb-switch-action" onClick={() => run(() => files.open())}>
        <FolderOpen size={16} strokeWidth={1.6} aria-hidden />
        Open...
      </button>
    </div>
  )
}

/** The title bar's workbook name as a button that opens the list (with how many are open, when more than one). */
export function TitleSwitcher({ files, title }: { files: FileCommands; title: string }) {
  const [open, setOpen] = useState(false)
  const count = files.docs.length
  return (
    <Popover
      open={open}
      onOpenChange={setOpen}
      trigger={
        <button type="button" className="titlebar-title titlebar-switch" aria-label="Switch workbooks" title={title}>
          <span className="titlebar-switch-name">{title}</span>
          {count > 1 ? <span className="wb-switch-count">{count}</span> : null}
          <ChevronDown size={14} aria-hidden />
        </button>
      }
    >
      <WorkbookList files={files} onDone={() => setOpen(false)} />
    </Popover>
  )
}

/** The phone's version: the list in a sheet from the bottom of the screen. */
export function WorkbookSheet({ files, onClose }: { files: FileCommands; onClose: () => void }) {
  return (
    <BottomSheet label="Workbooks" onClose={onClose} header={<span className="phone-sheet-title">Workbooks</span>}>
      <WorkbookList files={files} onDone={onClose} />
    </BottomSheet>
  )
}
