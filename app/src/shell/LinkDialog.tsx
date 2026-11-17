import { useMemo, useState } from "react"
import { FilePlus2, Globe, Mail, Table2 } from "lucide-react"
import {
  completeAddress,
  hyperlinkKind,
  linkFromText,
  linkText,
  mailAddress,
  mailParts,
  quoteSheet,
  type Hyperlink,
} from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { usePlatform } from "@/platform/context"
import { Dialog } from "./Dialog"

type Page = "file" | "place" | "new" | "mail"

/** The place list: cell references by sheet, then the workbook's names. */
type Place = { kind: "sheet"; sheet: string } | { kind: "name"; name: string }

/**
 * Insert › Link (Ctrl+K) and the cell menu's Edit Link: Excel's Insert Hyperlink dialog. "Link to" picks a web page
 * or file, a place in this workbook, or an e-mail address; Text to display and the ScreenTip are shared. Editing
 * adds Remove Link.
 */
export function LinkDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const platform = usePlatform()
  const editing = useMemo(() => controller.hyperlinks.activeLink(), [controller])
  const shown = useMemo(() => controller.hyperlinks.displayText(), [controller])
  const start = editing ? hyperlinkKind(editing) : "web"
  const [page, setPage] = useState<Page>(start === "mail" ? "mail" : start === "place" ? "place" : "file")
  const [text, setText] = useState(shown ?? "")
  const [tooltip, setTooltip] = useState(editing?.tooltip ?? "")
  const [tipOpen, setTipOpen] = useState(false)
  const [address, setAddress] = useState(() =>
    editing && (start === "web" || start === "file")
      ? editing.address! + (editing.location ? "#" + editing.location : "")
      : "",
  )
  const sheets = controller.wb.sheets.filter((s) => !s.hidden).map((s) => s.name)
  const startPlace = placeFromLocation(editing && start === "place" ? editing.location! : "", sheets[0] ?? "")
  const [cellRef, setCellRef] = useState(startPlace.ref)
  const [place, setPlace] = useState<Place>(startPlace.place)
  const mail = editing && start === "mail" ? mailParts(editing.address!) : { to: "", subject: "" }
  const [mailTo, setMailTo] = useState(mail.to)
  const [subject, setSubject] = useState(mail.subject)
  // Text to display follows the target until it's typed in (Excel fills it from the address as you type).
  const [textTouched, setTextTouched] = useState(Boolean(shown))
  const names = controller.names.all().map((n) => n.name)

  const target = (): Omit<Hyperlink, "range"> | null => {
    if (page === "file") return address.trim() ? linkFromText(address) : null
    if (page === "mail") return mailTo.trim() ? { address: mailAddress(mailTo, subject) } : null
    if (page === "place") {
      if (place.kind === "name") return { location: place.name }
      const ref = cellRef.trim() || "A1"
      return { location: `${quoteSheet(place.sheet)}!${ref}` }
    }
    return null
  }

  const follow = (next: Omit<Hyperlink, "range"> | null) => {
    if (!textTouched && shown !== null) setText(next ? linkText(next) : "")
  }

  const accept = () => {
    const link = target()
    if (!link) {
      controller.shell.alert(
        page === "mail" ? "Type an e-mail address." : page === "file" ? "Type an address." : "Choose a place.",
      )
      return
    }
    if (link.location && !link.address && !controller.names.resolve(link.location)) {
      controller.shell.alert("Reference isn't valid.")
      return
    }
    const display = shown === null ? null : text || linkText(link)
    controller.hyperlinks.set({ ...link, tooltip, text: display })
    onClose()
  }

  const remove = () => {
    controller.hyperlinks.remove(true)
    onClose()
  }

  const browse = async () => {
    const path = await platform.files.pickPath?.("Link to File")
    if (!path) return
    setAddress(path)
    follow({ address: path })
  }

  const pages: { id: Page; label: string; icon: React.ReactNode; title?: string }[] = [
    { id: "file", label: "Existing File or Web Page", icon: <Globe size={22} strokeWidth={1.5} /> },
    { id: "place", label: "Place in This Document", icon: <Table2 size={22} strokeWidth={1.5} /> },
    {
      id: "new",
      label: "Create New Document",
      icon: <FilePlus2 size={22} strokeWidth={1.5} />,
      title: "Cells shows one workbook per window: save the new workbook first, then link to it.",
    },
    { id: "mail", label: "E-mail Address", icon: <Mail size={22} strokeWidth={1.5} /> },
  ]

  return (
    <Dialog
      title={editing ? "Edit Hyperlink" : "Insert Hyperlink"}
      onCancel={onClose}
      onAccept={accept}
      width={600}
      actions={
        <>
          {editing ? (
            <button type="button" className="dialog-btn link-remove" onClick={remove}>
              <u>R</u>emove Link
            </button>
          ) : null}
          <button type="button" className="dialog-btn is-primary" onClick={accept}>
            OK
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <div className="link-dialog">
        <div className="link-to">
          <div className="dialog-label">Link to:</div>
          <div className="link-pages" role="tablist" aria-label="Link to">
            {pages.map((p) => (
              <button
                key={p.id}
                type="button"
                role="tab"
                aria-selected={page === p.id}
                className={page === p.id ? "link-page is-selected" : "link-page"}
                disabled={p.id === "new"}
                title={p.title}
                onClick={() => {
                  setPage(p.id)
                  if (p.id === "place")
                    follow({
                      location: `${quoteSheet(place.kind === "sheet" ? place.sheet : (sheets[0] ?? ""))}!${cellRef}`,
                    })
                }}
              >
                {p.icon}
                <span>{p.label}</span>
              </button>
            ))}
          </div>
        </div>
        <div className="link-main">
          <div className="link-row">
            <label className="dialog-label" htmlFor="link-text">
              <u>T</u>ext to display:
            </label>
            <input
              id="link-text"
              className="dialog-field"
              value={shown === null ? "<<Selection in Document>>" : text}
              disabled={shown === null}
              onChange={(e) => {
                setText(e.target.value)
                setTextTouched(true)
              }}
            />
            <button type="button" className="dialog-btn" onClick={() => setTipOpen(true)}>
              ScreenTip...
            </button>
          </div>

          {page === "file" ? (
            <div className="link-page-body">
              <label className="dialog-label" htmlFor="link-address">
                Addr<u>e</u>ss:
              </label>
              <div className="link-row">
                <input
                  id="link-address"
                  className="dialog-field"
                  data-autofocus
                  value={address}
                  placeholder="https://… or a file's path"
                  onChange={(e) => {
                    setAddress(e.target.value)
                    follow(e.target.value.trim() ? linkFromText(e.target.value) : null)
                  }}
                  onBlur={() => setAddress((a) => (a.trim() && !a.includes("#") ? completeAddress(a) : a))}
                />
                {platform.files.pickPath ? (
                  <button type="button" className="dialog-btn" onClick={() => void browse()}>
                    Browse...
                  </button>
                ) : null}
              </div>
              <p className="dialog-hint">
                A web page, or a file. To go to a place in a workbook, add it after #: Prices.xlsx#Sheet1!A1.
              </p>
            </div>
          ) : null}

          {page === "place" ? (
            <div className="link-page-body">
              <label className="dialog-label" htmlFor="link-ref">
                Type the cell refere<u>n</u>ce:
              </label>
              <input
                id="link-ref"
                className="dialog-field"
                data-autofocus
                value={cellRef}
                disabled={place.kind === "name"}
                onChange={(e) => {
                  setCellRef(e.target.value)
                  if (place.kind === "sheet") follow({ location: `${quoteSheet(place.sheet)}!${e.target.value}` })
                }}
              />
              <div className="dialog-label">
                Or select a <u>p</u>lace in this document:
              </div>
              <div className="dialog-list link-places" role="listbox" aria-label="Place in this document">
                <div className="link-places-head">Cell Reference</div>
                {sheets.map((name) => {
                  const selected = place.kind === "sheet" && place.sheet === name
                  return (
                    <button
                      key={"s:" + name}
                      type="button"
                      role="option"
                      aria-selected={selected}
                      className={selected ? "dialog-list-item is-selected" : "dialog-list-item"}
                      onClick={() => {
                        setPlace({ kind: "sheet", sheet: name })
                        follow({ location: `${quoteSheet(name)}!${cellRef}` })
                      }}
                    >
                      {name}
                    </button>
                  )
                })}
                {names.length ? <div className="link-places-head">Defined Names</div> : null}
                {names.map((name) => {
                  const selected = place.kind === "name" && place.name === name
                  return (
                    <button
                      key={"n:" + name}
                      type="button"
                      role="option"
                      aria-selected={selected}
                      className={selected ? "dialog-list-item is-selected" : "dialog-list-item"}
                      onClick={() => {
                        setPlace({ kind: "name", name })
                        follow({ location: name })
                      }}
                    >
                      {name}
                    </button>
                  )
                })}
              </div>
            </div>
          ) : null}

          {page === "mail" ? (
            <div className="link-page-body">
              <label className="dialog-label" htmlFor="link-mail">
                E-<u>m</u>ail address:
              </label>
              <input
                id="link-mail"
                className="dialog-field"
                data-autofocus
                value={mailTo}
                onChange={(e) => {
                  setMailTo(e.target.value)
                  follow(e.target.value.trim() ? { address: mailAddress(e.target.value, subject) } : null)
                }}
              />
              <label className="dialog-label" htmlFor="link-subject">
                S<u>u</u>bject:
              </label>
              <input
                id="link-subject"
                className="dialog-field"
                value={subject}
                onChange={(e) => {
                  setSubject(e.target.value)
                  if (mailTo.trim()) follow({ address: mailAddress(mailTo, e.target.value) })
                }}
              />
            </div>
          ) : null}
        </div>
      </div>
      {tipOpen ? <ScreenTipDialog value={tooltip} onClose={() => setTipOpen(false)} onSet={setTooltip} /> : null}
    </Dialog>
  )
}

/** ScreenTip...: the text shown when the pointer rests on the link (instead of its address). */
function ScreenTipDialog({
  value,
  onSet,
  onClose,
}: {
  value: string
  onSet: (text: string) => void
  onClose: () => void
}) {
  const [text, setText] = useState(value)
  return (
    <Dialog
      title="Set Hyperlink ScreenTip"
      onCancel={onClose}
      onAccept={() => {
        onSet(text)
        onClose()
      }}
      width={360}
    >
      <label className="dialog-label" htmlFor="link-tip">
        ScreenTip text:
      </label>
      <input id="link-tip" className="dialog-field" value={text} onChange={(e) => setText(e.target.value)} />
      <p className="dialog-hint">Shown when the pointer rests on the link.</p>
    </Dialog>
  )
}

/** The place list's choice and the cell reference for a link's location ("'Q1 Sales'!B4", a name). */
function placeFromLocation(location: string, firstSheet: string): { place: Place; ref: string } {
  const bang = location.lastIndexOf("!")
  if (bang < 0)
    return location
      ? { place: { kind: "name", name: location }, ref: "A1" }
      : { place: { kind: "sheet", sheet: firstSheet }, ref: "A1" }
  const sheet = location
    .slice(0, bang)
    .replace(/^'(.*)'$/, "$1")
    .replace(/''/g, "'")
  return { place: { kind: "sheet", sheet }, ref: location.slice(bang + 1) || "A1" }
}
