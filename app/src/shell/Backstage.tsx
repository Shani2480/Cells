import { ArrowLeft, FileSpreadsheet, FolderOpen, Pin, PinOff, X } from "lucide-react"
import { useEffect, useState } from "react"
import { CURRENT_VERSION, RELEASE_NOTES } from "@/about/release-notes"
import type { FileCommands } from "@/file/commands"
import { folderOf, type RecentFile } from "@/file/recent"
import { cn } from "@/lib/cn"
import type { SheetController } from "@/sheet/controller"
import type { PageSetupTab } from "./PageSetupDialog"
import { ExportPage, PrintPage } from "./PrintPage"

/**
 * File (the backstage view), as Excel's: a pane of pages on the left in the theme colour, the page on the right.
 * Opens opaque at once with only its pane sliding in, and closes at once (a fading full-window view ghosts in a
 * desktop webview). Save and Close act at once, as in Excel; Options opens Excel Options over the sheet; Share comes later.
 */
export type BackstagePage = "home" | "new" | "open" | "info" | "saveAs" | "print" | "export" | "account" | "about"

type NavItem = {
  id: string
  label: string
  page?: BackstagePage
  action?: "save" | "close" | "options"
  gapBefore?: boolean
}

const PAGES: NavItem[] = [
  { id: "home", label: "Home", page: "home" },
  { id: "new", label: "New", page: "new" },
  { id: "open", label: "Open", page: "open" },
  { id: "info", label: "Info", page: "info", gapBefore: true },
  { id: "save", label: "Save", action: "save" },
  { id: "saveAs", label: "Save As", page: "saveAs" },
  { id: "print", label: "Print", page: "print" },
  { id: "share", label: "Share" },
  { id: "export", label: "Export", page: "export" },
  { id: "close", label: "Close", action: "close" },
  { id: "account", label: "Account", page: "account", gapBefore: true },
  { id: "options", label: "Options", action: "options" },
]

export type BackstageFile = {
  title: string
  fileName: string
  /** The folder it's in, for Info (empty for a new workbook or a browser file). */
  folder: string
  saved: boolean
  hasFile: boolean
}

export function Backstage({
  page: initial,
  onClose,
  files,
  file,
  recent = [],
  controller,
  onPageSetup,
  onOptions,
}: {
  page?: string
  onClose: () => void
  files?: FileCommands | null
  file?: BackstageFile
  recent?: RecentFile[]
  /** The workbook shown (Print and Export draw its pages). */
  controller?: SheetController
  /** Print's Page Setup link: the dialog opens over the sheet. */
  onPageSetup?: (tab: PageSetupTab) => void
  /** Options: the File view closes and Excel Options opens over the sheet. */
  onOptions?: () => void
}) {
  const start = initial === "about" ? "account" : (initial ?? "home")
  const [page, setPage] = useState<string>(files || start === "account" ? start : "account")

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault()
        event.stopPropagation()
        onClose()
      }
    }
    window.addEventListener("keydown", onKey, true)
    return () => window.removeEventListener("keydown", onKey, true)
  }, [onClose])

  /** Runs a file command; the File view closes when it went through. */
  const run = (command: () => Promise<boolean>) => {
    void command().then((done) => {
      if (done) onClose()
    })
  }

  const select = (item: NavItem) => {
    if (item.action === "options") {
      onOptions?.()
      return
    }
    if (!files) return
    if (item.action === "save") {
      // A workbook never saved goes to Save As, as Excel does.
      if (!file?.hasFile) setPage("saveAs")
      else run(() => files.save())
    } else if (item.action === "close") run(() => files.close())
    else if (item.page) setPage(item.page)
  }

  const enabled = (item: NavItem) =>
    item.page === "account"
      ? true
      : item.action === "options"
        ? Boolean(onOptions)
        : Boolean(files && (item.page || item.action))

  return (
    <div className="backstage" role="dialog" aria-label="File">
      <nav className="backstage-nav">
        <button type="button" className="backstage-back" aria-label="Back" title="Back (Esc)" onClick={onClose}>
          <ArrowLeft size={18} strokeWidth={1.75} />
        </button>
        {PAGES.map((p) => (
          <button
            key={p.id}
            type="button"
            className={cn("backstage-item", p.gapBefore && "has-gap", page === p.id && "is-current")}
            aria-current={page === p.id ? "page" : undefined}
            disabled={!enabled(p)}
            title={enabled(p) ? undefined : "Not built yet."}
            onClick={() => (p.page === "account" ? setPage("account") : select(p))}
          >
            {p.label}
          </button>
        ))}
      </nav>
      <main className="backstage-page">
        {page === "account" ? <AccountPage /> : null}
        {files && (page === "home" || page === "new") ? (
          <HomePage
            title={page === "home" ? greeting() : "New"}
            showRecent={page === "home"}
            recent={recent}
            files={files}
            run={run}
          />
        ) : null}
        {files && page === "open" ? <OpenPage recent={recent} files={files} run={run} /> : null}
        {files && page === "saveAs" ? <SaveAsPage file={file} run={run} files={files} /> : null}
        {page === "info" && file ? <InfoPage file={file} /> : null}
        {page === "print" && controller && file ? (
          <PrintPage
            controller={controller}
            file={file}
            onPageSetup={(tab) => {
              onClose()
              onPageSetup?.(tab)
            }}
            onDone={onClose}
          />
        ) : null}
        {page === "export" && controller && file && files ? (
          <ExportPage controller={controller} file={file} files={files} run={run} />
        ) : null}
      </main>
    </div>
  )
}

function HomePage({
  title,
  showRecent,
  recent,
  files,
  run,
}: {
  title: string
  showRecent: boolean
  recent: RecentFile[]
  files: FileCommands
  run: (command: () => Promise<boolean>) => void
}) {
  return (
    <div className="backstage-content">
      <h1>{title}</h1>
      <div className="bs-templates">
        <button type="button" className="bs-template" onClick={() => run(() => files.newWorkbook())}>
          <span className="bs-template-thumb" data-keep-light aria-hidden>
            <BlankSheetThumb />
          </span>
          <span>Blank workbook</span>
        </button>
      </div>
      {showRecent ? (
        <>
          <h2 className="bs-section">Recent</h2>
          <RecentList recent={recent.slice(0, 12)} files={files} run={run} />
        </>
      ) : null}
    </div>
  )
}

function OpenPage({
  recent,
  files,
  run,
}: {
  recent: RecentFile[]
  files: FileCommands
  run: (command: () => Promise<boolean>) => void
}) {
  return (
    <div className="backstage-content">
      <h1>Open</h1>
      <div className="bs-split">
        <div className="bs-places">
          <button type="button" className="bs-place is-current">
            <FileSpreadsheet size={20} strokeWidth={1.5} />
            Recent
          </button>
          <button type="button" className="bs-place" onClick={() => run(() => files.open())}>
            <FolderOpen size={20} strokeWidth={1.5} />
            Browse
          </button>
        </div>
        <div className="bs-split-main">
          <RecentList recent={recent} files={files} run={run} />
        </div>
      </div>
    </div>
  )
}

function SaveAsPage({
  file,
  files,
  run,
}: {
  file?: BackstageFile
  files: FileCommands
  run: (command: () => Promise<boolean>) => void
}) {
  return (
    <div className="backstage-content">
      <h1>Save As</h1>
      <div className="bs-split">
        <div className="bs-places">
          <button type="button" className="bs-place is-current" onClick={() => run(() => files.saveAs())}>
            <FolderOpen size={20} strokeWidth={1.5} />
            Browse
          </button>
        </div>
        <div className="bs-split-main">
          <p className="bs-hint">
            Choose where to save {file ? <strong>{file.fileName}</strong> : "the workbook"} and in which format.
          </p>
          <button type="button" className="bs-button" onClick={() => run(() => files.saveAs())}>
            Browse…
          </button>
        </div>
      </div>
    </div>
  )
}

function InfoPage({ file }: { file: BackstageFile }) {
  return (
    <div className="backstage-content">
      <h1>Info</h1>
      <h2 className="bs-file-title">{file.title}</h2>
      {file.folder ? <p className="bs-hint">{file.folder}</p> : null}
      <dl className="bs-props">
        <dt>File</dt>
        <dd>{file.hasFile ? file.fileName : "Not saved yet"}</dd>
        <dt>Changes</dt>
        <dd>{file.saved ? "All changes saved" : "Unsaved changes"}</dd>
      </dl>
    </div>
  )
}

function RecentList({
  recent,
  files,
  run,
}: {
  recent: RecentFile[]
  files: FileCommands
  run: (command: () => Promise<boolean>) => void
}) {
  if (!recent.length)
    return <p className="bs-hint">You haven't opened any files recently. Click Browse to look for a file.</p>
  return (
    <ul className="bs-recent" aria-label="Recent files">
      {recent.map((entry) => (
        <li key={(entry.path ?? entry.name) + entry.opened}>
          <button
            type="button"
            className="bs-recent-open"
            title={entry.path && !entry.path.startsWith("content://") ? entry.path : entry.name}
            onClick={() => run(() => files.openRecent(entry))}
          >
            <FileSpreadsheet size={22} strokeWidth={1.5} className="bs-recent-icon" />
            <span className="bs-recent-text">
              <span className="bs-recent-name">{entry.name}</span>
              <span className="bs-recent-folder">{folderOf(entry.path)}</span>
            </span>
            <span className="bs-recent-date">{new Date(entry.opened).toLocaleDateString()}</span>
          </button>
          <button
            type="button"
            className="bs-recent-action"
            aria-label={entry.pinned ? "Unpin from list" : "Pin to list"}
            title={entry.pinned ? "Unpin from list" : "Pin to list"}
            onClick={() => files.pinRecent(entry, !entry.pinned)}
          >
            {entry.pinned ? <PinOff size={14} strokeWidth={1.75} /> : <Pin size={14} strokeWidth={1.75} />}
          </button>
          <button
            type="button"
            className="bs-recent-action"
            aria-label="Remove from list"
            title="Remove from list"
            onClick={() => files.removeRecent(entry)}
          >
            <X size={14} strokeWidth={1.75} />
          </button>
        </li>
      ))}
    </ul>
  )
}

/** Excel's Home page greets by the time of day. */
function greeting(hour = new Date().getHours()) {
  return hour < 12 ? "Good morning" : hour < 18 ? "Good afternoon" : "Good evening"
}

/** A blank sheet's thumbnail: a small grid, like Excel's Blank workbook tile. */
function BlankSheetThumb() {
  const cols = [0, 1, 2, 3, 4]
  const rows = [0, 1, 2, 3, 4, 5, 6, 7]
  return (
    <svg viewBox="0 0 120 150" width="100%" height="100%">
      <rect x="0.5" y="0.5" width="119" height="149" fill="#ffffff" stroke="#d4d4d4" />
      {cols.map((c) => (
        <line key={`c${c}`} x1={12 + c * 24} y1="12" x2={12 + c * 24} y2="150" stroke="#e1e1e1" />
      ))}
      {rows.map((r) => (
        <line key={`r${r}`} x1="0" y1={12 + r * 18} x2="120" y2={12 + r * 18} stroke="#e1e1e1" />
      ))}
      <rect x="12" y="12" width="24" height="18" fill="none" stroke="#217346" strokeWidth="2" />
    </svg>
  )
}

function AccountPage() {
  return (
    <div className="backstage-content">
      <h1>Account</h1>
      <section className="about-card">
        <img src="/logo-mark.svg" alt="" width={48} height={48} />
        <div>
          <h2>Cells</h2>
          <p className="about-version">Version {CURRENT_VERSION}</p>
          <p className="about-text">A spreadsheet that works offline, on your own machine.</p>
        </div>
      </section>
      <h2 className="whats-new-title">What's New</h2>
      <div className="whats-new">
        {RELEASE_NOTES.map((note) => (
          <section key={note.version}>
            <h3>
              Version {note.version} <span>{note.date}</span>
            </h3>
            <ul>
              {note.notes.map((line, i) => (
                <li key={i}>{line}</li>
              ))}
            </ul>
          </section>
        ))}
      </div>
    </div>
  )
}
