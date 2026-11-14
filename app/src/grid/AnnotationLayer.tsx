import { NOTE_HEIGHT, NOTE_WIDTH, type CellComment, type CellNote } from "@cells/core"
import { Check, MoreHorizontal, RotateCcw, SendHorizontal } from "lucide-react"
import { useEffect, useLayoutEffect, useRef, useState } from "react"
import type { CellAt } from "@/sheet/annotations"
import type { SheetController } from "@/sheet/controller"
import { rangeRect, type GridLayout } from "./layout"

type Props = {
  controller: SheetController
  layout: GridLayout
  zoom: number
  /** The cell under the pointer when it has a note or comment. */
  hover: CellAt | null
  /** The pointer is over a hover card: keep it. */
  onKeepHover: () => void
  onLeaveHover: () => void
  /** Back to the sheet (focus the grid's input). */
  onDone: () => void
}

/** Where a note box or comment card sits: right of the cell, a little above it (Excel's default placement). */
function anchor(controller: SheetController, layout: GridLayout, at: CellAt, zoom: number) {
  const sheet = controller.sheet
  const range = sheet.merges.at(at.row, at.col) ?? { r0: at.row, c0: at.col, r1: at.row, c1: at.col }
  const shown = layout.rows.some((l) => l.index === range.r0) && layout.cols.some((l) => l.index === range.c1)
  if (!shown) return null
  const rect = rangeRect(controller, layout, range)
  const cornerX = (rect.x + rect.w) * zoom
  const cornerY = rect.y * zoom
  return { cornerX, cornerY, left: cornerX + 14, top: Math.max(layout.headerHeight * zoom + 2, cornerY - 8) }
}

/**
 * Notes and comment threads over the sheet: the yellow note boxes (shown, hovered or being typed in) with the line
 * to their cell, and the comment card (hovered or opened).
 */
export function AnnotationLayer({ controller, layout, zoom, hover, onKeepHover, onLeaveHover, onDone }: Props) {
  const a = controller.annotations
  const sheet = controller.sheet
  const notes: { at: CellAt; note: CellNote; editing: boolean }[] = []
  for (const item of sheet.annotations.all()) {
    if (item.value.kind !== "note") continue
    const editing = a.editing?.row === item.row && a.editing.col === item.col
    const hovered = hover?.row === item.row && hover.col === item.col
    if (editing || hovered || item.value.shown || a.showAll)
      notes.push({ at: { row: item.row, col: item.col }, note: item.value, editing })
  }
  const card = a.open ?? (hover && a.commentAt(hover.row, hover.col) ? { ...hover, draft: false } : null)
  const pinned = Boolean(a.open)
  const cardPlace = card ? anchor(controller, layout, card, zoom) : null
  return (
    <>
      {notes.map(({ at, note, editing }) => {
        const place = anchor(controller, layout, at, zoom)
        if (!place) return null
        return (
          <NoteBox
            key={`${at.row}:${at.col}`}
            note={note}
            editing={editing}
            place={place}
            zoom={zoom}
            onFinish={(text, size) => {
              a.finishNote(text, size)
              onDone()
            }}
          />
        )
      })}
      {card && cardPlace ? (
        <CommentCard
          key={`${card.row}:${card.col}:${pinned}`}
          controller={controller}
          at={card}
          thread={a.commentAt(card.row, card.col)}
          draft={card.draft}
          pinned={pinned}
          place={cardPlace}
          onPointerEnter={onKeepHover}
          onPointerLeave={onLeaveHover}
          onDone={onDone}
        />
      ) : null}
    </>
  )
}

type Place = { cornerX: number; cornerY: number; left: number; top: number }

/** The connector from the cell's corner to the box, as Excel draws it. */
function Connector({ place }: { place: Place }) {
  const x0 = place.cornerX
  const y0 = place.cornerY
  const x1 = place.left
  const y1 = place.top + 4
  const left = Math.min(x0, x1)
  const top = Math.min(y0, y1)
  return (
    <svg
      className="note-connector"
      style={{ left, top, width: Math.abs(x1 - x0) + 1, height: Math.abs(y1 - y0) + 1 }}
      aria-hidden
    >
      <line x1={x0 - left} y1={y0 - top} x2={x1 - left} y2={y1 - top} />
    </svg>
  )
}

function NoteBox({
  note,
  editing,
  place,
  zoom,
  onFinish,
}: {
  note: CellNote
  editing: boolean
  place: Place
  zoom: number
  onFinish: (text: string, size?: { width: number; height: number }) => void
}) {
  const width = note.width ?? NOTE_WIDTH
  const height = note.height ?? NOTE_HEIGHT
  const boxRef = useRef<HTMLDivElement>(null)
  const textRef = useRef<HTMLTextAreaElement>(null)
  const done = useRef(false)
  useLayoutEffect(() => {
    if (!editing) return
    done.current = false
    const t = textRef.current
    if (t) {
      t.focus()
      t.setSelectionRange(t.value.length, t.value.length)
    }
  }, [editing])
  const finish = () => {
    if (done.current) return
    done.current = true
    const box = boxRef.current
    const size = box ? { width: Math.round(box.offsetWidth), height: Math.round(box.offsetHeight) } : undefined
    const changed = size && (size.width !== width || size.height !== height)
    onFinish(textRef.current?.value ?? note.text, changed ? size : undefined)
  }
  return (
    <>
      <Connector place={place} />
      <div
        ref={boxRef}
        className={`note-box${editing ? " is-editing" : ""}`}
        data-keep-light
        role={editing ? "dialog" : "note"}
        aria-label={`Note by ${note.author}`}
        style={{
          left: place.left,
          top: place.top,
          width,
          height,
          transform: zoom !== 1 ? `scale(${zoom})` : undefined,
          transformOrigin: "0 0",
        }}
      >
        <strong className="note-author">{note.author}:</strong>
        {editing ? (
          <textarea
            ref={textRef}
            className="note-text"
            aria-label="Note text"
            defaultValue={note.text}
            data-own-escape
            onBlur={finish}
            onKeyDown={(event) => {
              event.stopPropagation()
              if (event.key === "Escape") {
                event.preventDefault()
                finish()
              }
            }}
          />
        ) : (
          <div className="note-text">{note.text}</div>
        )}
      </div>
    </>
  )
}

const initials = (name: string) =>
  name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w[0]!.toUpperCase())
    .join("") || "?"

/** A colour per author, from a small palette (as Office colours people's avatars). */
const AVATAR = ["#8764b8", "#038387", "#ca5010", "#4f6bed", "#986f0b", "#c239b3"]
const avatarColor = (name: string) => AVATAR[[...name].reduce((h, ch) => h + ch.charCodeAt(0), 0) % AVATAR.length]

const when = (time: number) => new Date(time).toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" })

function CommentCard({
  controller,
  at,
  thread,
  draft,
  pinned,
  place,
  onPointerEnter,
  onPointerLeave,
  onDone,
}: {
  controller: SheetController
  at: CellAt
  thread: CellComment | undefined
  draft: boolean
  pinned: boolean
  place: Place
  onPointerEnter: () => void
  onPointerLeave: () => void
  onDone: () => void
}) {
  const a = controller.annotations
  const [text, setText] = useState("")
  const [editing, setEditing] = useState<number | null>(null)
  const [editText, setEditText] = useState("")
  const [menu, setMenu] = useState(false)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  useEffect(() => {
    if (pinned && (draft || !thread?.resolved)) inputRef.current?.focus()
  }, [pinned, draft, thread?.resolved])
  const close = () => {
    a.closeThread()
    onDone()
  }
  const post = () => {
    if (!text.trim()) return
    a.post(text)
    setText("")
  }
  const keys = (event: React.KeyboardEvent, send: () => void) => {
    event.stopPropagation()
    if (event.key === "Escape") {
      event.preventDefault()
      close()
    } else if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
      event.preventDefault()
      send()
    }
  }
  const posts = thread?.posts ?? []
  return (
    <>
      <Connector place={place} />
      <div
        className={`comment-card${pinned ? " is-pinned" : ""}${thread?.resolved ? " is-resolved" : ""}`}
        role="dialog"
        aria-label="Comment"
        style={{ left: place.left, top: place.top }}
        onPointerEnter={onPointerEnter}
        onPointerLeave={onPointerLeave}
        onPointerDown={() => {
          if (!pinned && thread) a.openThread(at.row, at.col)
        }}
      >
        {thread ? (
          <div className="comment-head">
            {thread.resolved ? <span className="comment-resolved">Resolved</span> : <span />}
            <div className="comment-head-actions">
              {thread.resolved ? (
                <button
                  type="button"
                  className="comment-icon-button"
                  title="Reopen thread"
                  aria-label="Reopen thread"
                  onClick={() => a.resolve(at.row, at.col, false)}
                >
                  <RotateCcw size={14} strokeWidth={1.6} />
                </button>
              ) : null}
              <button
                type="button"
                className="comment-icon-button"
                title="More thread actions"
                aria-label="More thread actions"
                aria-expanded={menu}
                onClick={() => setMenu((m) => !m)}
              >
                <MoreHorizontal size={16} strokeWidth={1.6} />
              </button>
            </div>
            {menu ? (
              <div className="comment-menu" role="menu">
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => {
                    setMenu(false)
                    a.deleteThread(at.row, at.col)
                    onDone()
                  }}
                >
                  Delete thread
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => {
                    setMenu(false)
                    a.resolve(at.row, at.col, !thread.resolved)
                  }}
                >
                  {thread.resolved ? "Reopen thread" : "Resolve thread"}
                </button>
              </div>
            ) : null}
          </div>
        ) : null}
        {posts.map((p, i) => (
          <div className="comment-post" key={i}>
            <span className="comment-avatar" style={{ background: avatarColor(p.author) }} aria-hidden>
              {initials(p.author)}
            </span>
            <div className="comment-body">
              <div className="comment-meta">
                <strong>{p.author}</strong>
                <span>
                  {when(p.time)}
                  {p.edited ? " (edited)" : ""}
                </span>
              </div>
              {editing === i ? (
                <div className="comment-edit">
                  <textarea
                    className="comment-input"
                    aria-label="Edit comment"
                    value={editText}
                    autoFocus
                    onChange={(e) => setEditText(e.target.value)}
                    onKeyDown={(e) =>
                      keys(e, () => {
                        a.editPost(at.row, at.col, i, editText)
                        setEditing(null)
                      })
                    }
                  />
                  <div className="comment-actions">
                    <button type="button" className="comment-text-button" onClick={() => setEditing(null)}>
                      Cancel
                    </button>
                    <button
                      type="button"
                      className="comment-text-button is-primary"
                      onClick={() => {
                        a.editPost(at.row, at.col, i, editText)
                        setEditing(null)
                      }}
                    >
                      Save
                    </button>
                  </div>
                </div>
              ) : (
                <div className="comment-text">{p.text}</div>
              )}
              {pinned && editing !== i && !thread?.resolved ? (
                <div className="comment-post-actions">
                  <button
                    type="button"
                    className="comment-link"
                    onClick={() => {
                      setEditing(i)
                      setEditText(p.text)
                    }}
                  >
                    Edit
                  </button>
                  <button type="button" className="comment-link" onClick={() => a.deletePost(at.row, at.col, i)}>
                    Delete
                  </button>
                  {i === 0 ? (
                    <button
                      type="button"
                      className="comment-link"
                      title="Resolve thread"
                      onClick={() => a.resolve(at.row, at.col, true)}
                    >
                      <Check size={12} strokeWidth={2} /> Resolve
                    </button>
                  ) : null}
                </div>
              ) : null}
            </div>
          </div>
        ))}
        {!thread?.resolved ? (
          <div className="comment-reply">
            {draft ? (
              <div className="comment-meta">
                <span className="comment-avatar" style={{ background: avatarColor(a.author) }} aria-hidden>
                  {initials(a.author)}
                </span>
                <strong>{a.author}</strong>
              </div>
            ) : null}
            <div className="comment-reply-row">
              <textarea
                ref={inputRef}
                className="comment-input"
                aria-label={draft ? "Start a conversation" : "Reply"}
                placeholder={draft ? "Start a conversation" : "Reply…"}
                rows={1}
                value={text}
                onFocus={() => {
                  if (!pinned) a.openThread(at.row, at.col)
                }}
                onChange={(e) => setText(e.target.value)}
                onKeyDown={(e) => keys(e, post)}
              />
              <button
                type="button"
                className="comment-icon-button is-send"
                title="Post (Ctrl+Enter)"
                aria-label="Post"
                disabled={!text.trim()}
                onClick={post}
              >
                <SendHorizontal size={16} strokeWidth={1.6} />
              </button>
            </div>
            {draft ? (
              <div className="comment-actions">
                <button type="button" className="comment-text-button" onClick={close}>
                  Cancel
                </button>
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
    </>
  )
}
