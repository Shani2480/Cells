import { columnName } from "@cells/core"
import { MessageSquarePlus, X } from "lucide-react"
import { useSyncExternalStore } from "react"
import { initialsOf } from "@/settings/shared"
import type { SheetController } from "@/sheet/controller"

const when = (time: number) => new Date(time).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" })

/**
 * Review › Show Comments: the sheet's comment threads in a pane beside it, in reading order, as Excel's Comments pane
 * lists them. Picking a thread goes to its cell and opens it there (replies, edits and resolving happen in its card);
 * New starts a thread on the active cell. On a phone it takes the whole screen.
 */
export function CommentsPane({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const a = controller.annotations
  const threads = a.threads()
  const close = () => {
    a.togglePane(false)
    onDone()
  }
  return (
    <aside
      className="comments-pane"
      aria-label="Comments"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Comments</h2>
        <button
          type="button"
          className="comments-pane-new"
          onClick={() => {
            a.newComment()
          }}
        >
          <MessageSquarePlus size={16} strokeWidth={1.6} aria-hidden />
          New
        </button>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.6} aria-hidden />
        </button>
      </header>
      <div className="comments-pane-list">
        {threads.length === 0 ? (
          <p className="comments-pane-empty">There are no comments on this sheet. Select a cell and choose New.</p>
        ) : null}
        {threads.map(({ row, col, thread }) => {
          const address = `${columnName(col)}${row + 1}`
          const open = a.open?.row === row && a.open.col === col
          return (
            <button
              key={address}
              type="button"
              className={`comments-pane-thread${open ? " is-open" : ""}${thread.resolved ? " is-resolved" : ""}`}
              aria-label={`Comment on ${address}`}
              aria-current={open || undefined}
              onClick={() => {
                controller.goTo({ r0: row, c0: col, r1: row, c1: col })
                a.openThread(row, col)
              }}
            >
              <span className="comments-pane-cell">
                {address}
                {thread.resolved ? <span className="comments-pane-resolved">Resolved</span> : null}
              </span>
              {thread.posts.map((post, i) => (
                <span key={i} className="comments-pane-post">
                  <span className="comments-pane-avatar" aria-hidden>
                    {initialsOf(post.author)}
                  </span>
                  <span className="comments-pane-body">
                    <span className="comments-pane-meta">
                      <strong>{post.author}</strong> {when(post.time)}
                    </span>
                    <span className="comments-pane-text">{post.text}</span>
                  </span>
                </span>
              ))}
            </button>
          )
        })}
      </div>
    </aside>
  )
}
