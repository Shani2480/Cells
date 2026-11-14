import { CircleAlert } from "lucide-react"
import { useState } from "react"
import { Dialog } from "@/shell/Dialog"
import type { FileProblem } from "./document"
import type { RecoveredFile } from "./recovery"

/** Excel's "Want to save your changes to Book1?" with Save, Don't Save and Cancel. */
export function SaveChangesDialog({
  title,
  onChoose,
}: {
  title: string
  onChoose: (choice: "save" | "discard" | "cancel") => void
}) {
  return (
    <Dialog
      title="Cells"
      width={380}
      acceptLabel="Save"
      extra={{ label: "Don't Save", onSelect: () => onChoose("discard") }}
      onAccept={() => onChoose("save")}
      onCancel={() => onChoose("cancel")}
    >
      <p className="dialog-message">Want to save your changes to '{title}'?</p>
    </Dialog>
  )
}

/** A file that couldn't be opened or saved: what happened in plain words, the raw error under Details. */
export function FileProblemDialog({ problem, onClose }: { problem: FileProblem; onClose: () => void }) {
  const [copied, setCopied] = useState(false)
  return (
    <Dialog title="Cells" width={440} cancelLabel={null} onAccept={onClose} onCancel={onClose}>
      <div className="file-problem">
        <CircleAlert className="file-problem-icon" size={28} strokeWidth={1.75} aria-hidden />
        <p className="dialog-message">{problem.message}</p>
      </div>
      {problem.detail ? (
        <details className="file-details">
          <summary>Details</summary>
          <pre>{problem.detail}</pre>
          <button
            type="button"
            className="dialog-btn"
            onClick={() =>
              void navigator.clipboard
                ?.writeText(problem.detail)
                .then(() => setCopied(true))
                .catch(() => undefined)
            }
          >
            {copied ? "Copied" : "Copy"}
          </button>
        </details>
      ) : null}
    </Dialog>
  )
}

/** What Cells left out or changed while opening a file. */
export function FileNoticeDialog({
  fileName,
  lines,
  onClose,
}: {
  fileName: string
  lines: string[]
  onClose: () => void
}) {
  return (
    <Dialog title="Cells" width={480} cancelLabel={null} onAccept={onClose} onCancel={onClose}>
      <p className="dialog-message">'{fileName}' opened, with these differences:</p>
      <ul className="file-list">
        {lines.map((line) => (
          <li key={line}>{line}</li>
        ))}
      </ul>
    </Dialog>
  )
}

/** Document Recovery: copies AutoRecover kept from a session that ended unexpectedly. */
export function RecoveryDialog({
  files,
  onOpen,
  onDiscard,
  onClose,
}: {
  files: RecoveredFile[]
  onOpen: (file: RecoveredFile) => void
  onDiscard: (file: RecoveredFile) => void
  onClose: () => void
}) {
  return (
    <Dialog
      title="Document Recovery"
      width={460}
      acceptLabel="Close"
      cancelLabel={null}
      onAccept={onClose}
      onCancel={onClose}
    >
      <p className="dialog-message">
        Cells closed before these workbooks were saved. Open a recovered copy to keep your work, then save it.
      </p>
      <ul className="recovery-list">
        {files.map((file) => (
          <li key={file.id}>
            <span className="recovery-name">
              {file.title} (Recovered)
              <small>AutoRecover copy from {new Date(file.saved).toLocaleString()}</small>
            </span>
            <button type="button" className="dialog-btn is-primary" onClick={() => onOpen(file)}>
              Open
            </button>
            <button type="button" className="dialog-btn" onClick={() => onDiscard(file)}>
              Delete
            </button>
          </li>
        ))}
      </ul>
    </Dialog>
  )
}

/** "Opening Budget.xlsx" with a bar; `done` null: no measure (saving). */
export function FileProgress({ label, done }: { label: string; done: number | null }) {
  return (
    <div className="file-progress" role="status" aria-live="polite">
      <span>{label}…</span>
      <div
        className={done === null ? "file-progress-bar is-indeterminate" : "file-progress-bar"}
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={done === null ? undefined : Math.round(done * 100)}
      >
        <span style={done === null ? undefined : { width: `${Math.round(done * 100)}%` }} />
      </div>
    </div>
  )
}
