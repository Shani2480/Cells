import { FileSpreadsheet, FolderOpen, RefreshCw, Unlink, X } from "lucide-react"
import { useState, useSyncExternalStore } from "react"
import type { FileCommands } from "@/file/commands"
import type { SheetController } from "@/sheet/controller"
import { cachedLinkRows } from "@/sheet/links"

/**
 * Data › Workbook Links (Excel's Edit Links): the other workbooks this one's formulas use, each with where it is,
 * Update Values, Change Source and Break Link, and Update All on top. Formulas calculate from the values last read
 * from each file until it is updated. Docked beside the sheet; on a phone it takes the whole screen.
 */
const STATUS_HELP: Record<string, string> = {
  Updated: "Values read from the file in this session.",
  "Using saved values": "Formulas use the values saved with the link. Update Values reads the file again.",
  "Not read yet": "A formula names this workbook, but its values haven't been read. Choose Update Values.",
}

export function WorkbookLinksPane({
  controller,
  files,
  onDone,
}: {
  controller: SheetController
  files?: FileCommands
  onDone: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const [busy, setBusy] = useState<string | null>(null)
  const [updated, setUpdated] = useState<Set<string>>(() => new Set())
  const rows = cachedLinkRows(controller.wb)
  const close = () => {
    controller.data.toggleLinks(false)
    onDone()
  }
  const run = async (key: string, work: () => Promise<boolean>, names: string[] = []) => {
    setBusy(key)
    try {
      const ok = await work()
      if (ok && names.length) setUpdated((u) => new Set([...u, ...names.map((n) => n.toUpperCase())]))
    } finally {
      setBusy(null)
    }
  }
  const updateAll = async () => {
    let any = false
    for (const r of rows) {
      const ok = await files!.updateLink(r.name)
      if (ok) setUpdated((u) => new Set([...u, r.name.toUpperCase()]))
      any ||= ok
    }
    return any
  }
  return (
    <aside
      className="comments-pane queries-pane"
      aria-label="Workbook Links"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Workbook Links</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.6} aria-hidden />
        </button>
      </header>
      <div className="clipboard-pane-actions">
        <button
          type="button"
          className="comments-pane-new"
          disabled={!files || !rows.length || busy !== null}
          onClick={() => void run("*", updateAll)}
        >
          <RefreshCw size={16} strokeWidth={1.6} aria-hidden className={busy === "*" ? "queries-spin" : undefined} />
          Update All
        </button>
      </div>
      <div className="comments-pane-list" role="list">
        {!rows.length ? (
          <p className="comments-pane-empty">
            This workbook doesn&apos;t use other workbooks. A formula such as =[Prices.xlsx]Sheet1!A1 links to one, and
            it shows up here.
          </p>
        ) : null}
        {rows.map((r) => {
          const key = r.name.toUpperCase()
          const status = updated.has(key) ? "Updated" : r.known ? "Using saved values" : "Not read yet"
          return (
            <div key={key} role="listitem" className="queries-item">
              <div className="queries-go links-row">
                <FileSpreadsheet size={18} strokeWidth={1.5} aria-hidden className="queries-icon" />
                <span className="queries-text">
                  <strong>{r.name}</strong>
                  <span className="queries-meta" title={r.location}>
                    {r.location}
                  </span>
                  <span className="queries-meta" title={STATUS_HELP[status]}>
                    {status}
                  </span>
                  <span className="queries-meta">
                    Used by {r.cells.toLocaleString()} {r.cells === 1 ? "formula" : "formulas"}
                  </span>
                </span>
              </div>
              <span className="queries-actions">
                <button
                  type="button"
                  aria-label={`Update values from ${r.name}`}
                  title="Update Values"
                  disabled={!files || busy !== null}
                  onClick={() => void run(key, () => files!.updateLink(r.name), [r.name])}
                >
                  <RefreshCw
                    size={15}
                    strokeWidth={1.6}
                    aria-hidden
                    className={busy === key ? "queries-spin" : undefined}
                  />
                </button>
                <button
                  type="button"
                  aria-label={`Change source of ${r.name}`}
                  title="Change Source"
                  disabled={!files || busy !== null}
                  onClick={() => void run(key, () => files!.changeLinkSource(r.name))}
                >
                  <FolderOpen size={15} strokeWidth={1.6} aria-hidden />
                </button>
                <button
                  type="button"
                  aria-label={`Break link to ${r.name}`}
                  title="Break Link (formulas become their values)"
                  disabled={!files || busy !== null}
                  onClick={() => void run(key, () => files!.breakLink(r.name))}
                >
                  <Unlink size={15} strokeWidth={1.6} aria-hidden />
                </button>
              </span>
            </div>
          )
        })}
      </div>
    </aside>
  )
}
