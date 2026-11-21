import { useState } from "react"
import { CURRENT_VERSION } from "@/about/release-notes"
import { usePlatform } from "@/platform/context"
import { Dialog } from "@/shell/Dialog"
import { issueUrl, KIND_LABELS, systemLine, type FeedbackKind } from "./feedback"

/**
 * Help › Feedback ("support": Contact Support): what it's about, a title and details, then Open on GitHub, which opens
 * a new issue with this filled in. Nothing is sent until the person submits it there.
 */
export function FeedbackDialog({ mode, onClose }: { mode: "feedback" | "support"; onClose: () => void }) {
  const platform = usePlatform()
  const kinds: FeedbackKind[] = mode === "support" ? ["problem", "question"] : ["like", "dislike", "idea"]
  const [kind, setKind] = useState<FeedbackKind>(kinds[0])
  const [title, setTitle] = useState("")
  const [details, setDetails] = useState("")
  const [withSystem, setWithSystem] = useState(true)
  const [failed, setFailed] = useState(false)
  const system = systemLine(CURRENT_VERSION, platform.os, "__TAURI_INTERNALS__" in window)

  const send = () => {
    if (!title.trim()) return
    const url = issueUrl({ kind, title, details, system: withSystem ? system : undefined })
    platform.opens.external(url).then(onClose, () => setFailed(true))
  }

  return (
    <Dialog
      title={mode === "support" ? "Contact Support" : "Feedback"}
      width={460}
      onCancel={onClose}
      onAccept={send}
      actions={
        <>
          <button type="button" className="dialog-btn is-primary" disabled={!title.trim()} onClick={send}>
            Open on GitHub
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <fieldset className="feedback-kinds">
        <legend className="dialog-label">{mode === "support" ? "What do you need?" : "What's it about?"}</legend>
        {kinds.map((k) => (
          <label key={k} className="dialog-radio">
            <input type="radio" name="feedback-kind" checked={kind === k} onChange={() => setKind(k)} />
            {KIND_LABELS[k]}
          </label>
        ))}
      </fieldset>
      <label className="dialog-label" htmlFor="feedback-title">
        In a few words
      </label>
      <input
        id="feedback-title"
        className="dialog-field"
        data-autofocus
        value={title}
        placeholder={mode === "support" ? "e.g. Sorting by colour leaves a row out" : "e.g. Let me pin sheets"}
        onChange={(e) => setTitle(e.target.value)}
      />
      <label className="dialog-label" htmlFor="feedback-details">
        Tell us more
      </label>
      <textarea
        id="feedback-details"
        className="dialog-field feedback-details"
        rows={5}
        value={details}
        onChange={(e) => setDetails(e.target.value)}
        onKeyDown={(e) => {
          // Enter is a new line here, not OK.
          if (e.key === "Enter") e.stopPropagation()
        }}
      />
      <label className="fc-check">
        <input type="checkbox" checked={withSystem} onChange={(e) => setWithSystem(e.target.checked)} />
        <span>Add the version and system ({system})</span>
      </label>
      <p className="dialog-note">
        Opens a new issue on GitHub with this filled in. You can still change it there, and nothing is sent until you
        submit it on GitHub (you need a GitHub account). Your workbook and file names aren't included, so don't paste
        anything private.
      </p>
      {failed ? (
        <p className="dialog-note is-error" role="alert">
          Couldn't open the browser. The issues are at github.com/Shani2480/Cells/issues.
        </p>
      ) : null}
    </Dialog>
  )
}
