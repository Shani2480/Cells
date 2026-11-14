import { DEFAULT_PROTECTION, type ProtectionPermission } from "@cells/core"
import { useState } from "react"
import { SHEET_MESSAGES } from "@/sheet/sheets"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** Excel's "Allow all users of this worksheet to:" list, in its order. */
export const PERMISSIONS: { key: ProtectionPermission; label: string }[] = [
  { key: "selectLocked", label: "Select locked cells" },
  { key: "selectUnlocked", label: "Select unlocked cells" },
  { key: "formatCells", label: "Format cells" },
  { key: "formatColumns", label: "Format columns" },
  { key: "formatRows", label: "Format rows" },
  { key: "insertColumns", label: "Insert columns" },
  { key: "insertRows", label: "Insert rows" },
  { key: "insertHyperlinks", label: "Insert hyperlinks" },
  { key: "deleteColumns", label: "Delete columns" },
  { key: "deleteRows", label: "Delete rows" },
  { key: "sort", label: "Sort" },
  { key: "autoFilter", label: "Use AutoFilter" },
  { key: "pivotTables", label: "Use PivotTable & PivotChart" },
  { key: "editObjects", label: "Edit objects" },
  { key: "editScenarios", label: "Edit scenarios" },
]

export const CONFIRM_CAUTION =
  "Caution: If you lose or forget the password, it cannot be recovered. It is advisable to keep a list of passwords and their corresponding workbook and sheet names in a safe place. (Remember that passwords are case-sensitive.)"

/**
 * Protect Sheet: the options, then (with a password) Confirm Password, as Excel asks. A mismatch shows Excel's
 * message and goes back to the first dialog.
 */
export function ProtectSheetDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [on, setOn] = useState(true)
  const [password, setPassword] = useState("")
  const [options, setOptions] = useState(() => {
    const { password: _hash, ...rest } = { ...DEFAULT_PROTECTION }
    return rest
  })
  const [confirming, setConfirming] = useState(false)
  if (confirming)
    return (
      <ConfirmPassword
        password={password}
        onBack={() => setConfirming(false)}
        onMismatch={() => controller.shell.alert(SHEET_MESSAGES.passwordMismatch)}
        onConfirmed={() => {
          controller.sheets.protect(options, password)
          onClose()
        }}
      />
    )
  return (
    <Dialog
      title="Protect Sheet"
      width={340}
      onCancel={onClose}
      onAccept={() => {
        if (!on) return
        if (password) {
          setConfirming(true)
          return
        }
        controller.sheets.protect(options, "")
        onClose()
      }}
    >
      <label className="fc-check">
        <input type="checkbox" checked={on} onChange={(e) => setOn(e.target.checked)} />
        <span>
          Pro<u>t</u>ect worksheet and contents of locked cells
        </span>
      </label>
      <label className="dialog-label" htmlFor="protect-password">
        <u>P</u>assword to unprotect sheet:
      </label>
      <input
        id="protect-password"
        className="dialog-field"
        type="password"
        data-autofocus
        value={password}
        onChange={(e) => setPassword(e.target.value)}
      />
      <span className="dialog-label">
        <u>A</u>llow all users of this worksheet to:
      </span>
      <div className="dialog-list protect-list" role="group" aria-label="Allow all users of this worksheet to">
        {PERMISSIONS.map((p) => (
          <label className="fc-check" key={p.key}>
            <input
              type="checkbox"
              checked={options[p.key]}
              onChange={(e) => setOptions((o) => ({ ...o, [p.key]: e.target.checked }))}
            />
            <span>{p.label}</span>
          </label>
        ))}
      </div>
    </Dialog>
  )
}

/** Unprotect Sheet with a password. */
export function UnprotectSheetDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [password, setPassword] = useState("")
  return (
    <Dialog
      title="Unprotect Sheet"
      width={300}
      onCancel={onClose}
      onAccept={() => {
        controller.sheets.unprotect(password)
        onClose()
      }}
    >
      <label className="dialog-label" htmlFor="unprotect-password">
        <u>P</u>assword:
      </label>
      <input
        id="unprotect-password"
        className="dialog-field"
        type="password"
        value={password}
        onChange={(e) => setPassword(e.target.value)}
      />
    </Dialog>
  )
}

/** Confirm Password: the password again; a mismatch shows Excel's message and goes back to the first dialog. */
function ConfirmPassword({
  password,
  onBack,
  onMismatch,
  onConfirmed,
}: {
  password: string
  onBack: () => void
  onMismatch: () => void
  onConfirmed: () => void
}) {
  const [again, setAgain] = useState("")
  return (
    <Dialog
      title="Confirm Password"
      width={360}
      onCancel={onBack}
      onAccept={() => {
        if (again !== password) {
          onMismatch()
          onBack()
          return
        }
        onConfirmed()
      }}
    >
      <label className="dialog-label" htmlFor="protect-again">
        <u>R</u>eenter password to proceed.
      </label>
      <input
        id="protect-again"
        className="dialog-field"
        type="password"
        value={again}
        onChange={(e) => setAgain(e.target.value)}
      />
      <p className="dialog-note">{CONFIRM_CAUTION}</p>
    </Dialog>
  )
}

/**
 * Review › Protect Workbook: Excel's "Protect Structure and Windows". Windows is greyed out, as in Excel 365 (it only
 * applied to the old multiple-document window).
 */
export function ProtectWorkbookDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [structure, setStructure] = useState(true)
  const [password, setPassword] = useState("")
  const [confirming, setConfirming] = useState(false)
  if (confirming)
    return (
      <ConfirmPassword
        password={password}
        onBack={() => setConfirming(false)}
        onMismatch={() => controller.shell.alert(SHEET_MESSAGES.passwordMismatch)}
        onConfirmed={() => {
          controller.sheets.protectWorkbook(password)
          onClose()
        }}
      />
    )
  return (
    <Dialog
      title="Protect Structure and Windows"
      width={320}
      onCancel={onClose}
      onAccept={() => {
        if (!structure) return onClose()
        if (password) return setConfirming(true)
        controller.sheets.protectWorkbook("")
        onClose()
      }}
    >
      <span className="dialog-label">Protect workbook for</span>
      <label className="fc-check">
        <input type="checkbox" checked={structure} onChange={(e) => setStructure(e.target.checked)} />
        <span>
          <u>S</u>tructure
        </span>
      </label>
      <label className="fc-check is-disabled">
        <input type="checkbox" checked={false} disabled readOnly />
        <span>
          <u>W</u>indows
        </span>
      </label>
      <label className="dialog-label" htmlFor="protect-workbook-password">
        <u>P</u>assword (optional):
      </label>
      <input
        id="protect-workbook-password"
        className="dialog-field"
        type="password"
        data-autofocus
        value={password}
        onChange={(e) => setPassword(e.target.value)}
      />
    </Dialog>
  )
}

/** Unprotect Workbook with a password. */
export function UnprotectWorkbookDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [password, setPassword] = useState("")
  return (
    <Dialog
      title="Unprotect Workbook"
      width={300}
      onCancel={onClose}
      onAccept={() => {
        controller.sheets.unprotectWorkbook(password)
        onClose()
      }}
    >
      <label className="dialog-label" htmlFor="unprotect-workbook-password">
        <u>P</u>assword:
      </label>
      <input
        id="unprotect-workbook-password"
        className="dialog-field"
        type="password"
        data-autofocus
        value={password}
        onChange={(e) => setPassword(e.target.value)}
      />
    </Dialog>
  )
}
