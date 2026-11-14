import { ArrowLeft } from "lucide-react"
import { useEffect, useState } from "react"
import { CURRENT_VERSION, RELEASE_NOTES } from "@/about/release-notes"
import { cn } from "@/lib/cn"

/**
 * File (the backstage view), as Excel's: a pane of pages on the left in the theme colour, the page on the right.
 * Opens opaque at once with only its pane sliding in, and closes at once (a fading full-window view ghosts in a
 * desktop webview). Pages that need files (Open, Save, Print…) come with file support and show as not built yet.
 */
export type BackstagePage = "home" | "info" | "account" | "about"

const PAGES: { id: string; label: string; built?: BackstagePage; gapBefore?: boolean }[] = [
  { id: "home", label: "Home" },
  { id: "new", label: "New" },
  { id: "open", label: "Open" },
  { id: "info", label: "Info", gapBefore: true },
  { id: "save", label: "Save" },
  { id: "saveAs", label: "Save As" },
  { id: "print", label: "Print" },
  { id: "share", label: "Share" },
  { id: "export", label: "Export" },
  { id: "close", label: "Close" },
  { id: "account", label: "Account", built: "account", gapBefore: true },
  { id: "options", label: "Options" },
]

export function Backstage({ page: initial, onClose }: { page?: string; onClose: () => void }) {
  const [page, setPage] = useState<string>(initial === "about" ? "account" : (initial ?? "account"))

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
            disabled={!p.built}
            title={p.built ? undefined : "Not built yet."}
            onClick={() => setPage(p.id)}
          >
            {p.label}
          </button>
        ))}
      </nav>
      <main className="backstage-page">{page === "account" ? <AccountPage /> : null}</main>
    </div>
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
