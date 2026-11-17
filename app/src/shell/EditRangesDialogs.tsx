import { copyEditRanges, hashPassword, nextEditRangeName, type EditRange } from "@cells/core"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { EDIT_RANGE_MESSAGES, parseRefersTo, refersToText } from "@/sheet/edit-ranges"
import { SHEET_MESSAGES } from "@/sheet/sheets"
import { Dialog } from "./Dialog"
import { ConfirmPassword } from "./ProtectDialogs"

/**
 * Review › Allow Edit Ranges (Excel's "Allow Users to Edit Ranges"): the ranges of this sheet that stay editable once
 * it's protected, with New, Modify and Delete, and Protect Sheet to go straight on. Changes apply on OK or Apply.
 */
export function AllowEditRangesDialog({
  controller,
  onClose,
  onProtect,
}: {
  controller: SheetController
  onClose: () => void
  onProtect: () => void
}) {
  const [list, setList] = useState<EditRange[]>(() => copyEditRanges(controller.editRanges.list()))
  const [chosen, setChosen] = useState(0)
  const [editing, setEditing] = useState<{ index: number | null } | null>(null)
  const current = list[chosen]
  const apply = () => controller.editRanges.save(list)

  if (editing)
    return (
      <EditRangeDialog
        controller={controller}
        list={list}
        index={editing.index}
        onCancel={() => setEditing(null)}
        onDone={(range) => {
          const next = [...list]
          if (editing.index === null) next.push(range)
          else next[editing.index] = range
          setList(next)
          setChosen(editing.index ?? next.length - 1)
          setEditing(null)
        }}
      />
    )

  return (
    <Dialog
      title="Allow Users to Edit Ranges"
      width={460}
      onCancel={onClose}
      onAccept={() => {
        apply()
        onClose()
      }}
      extra={{ label: "Apply", onSelect: apply }}
    >
      <span className="dialog-label">Ranges unlocked by a password when sheet is protected:</span>
      <div className="edit-ranges">
        <div className="dialog-list edit-ranges-list" role="listbox" aria-label="Ranges">
          <div className="edit-ranges-head" aria-hidden="true">
            <span>Title</span>
            <span>Refers to cells</span>
          </div>
          {list.map((e, i) => (
            <button
              type="button"
              role="option"
              aria-selected={i === chosen}
              key={i}
              className={`dialog-list-item edit-ranges-row${i === chosen ? " is-selected" : ""}`}
              onClick={() => setChosen(i)}
              onDoubleClick={() => setEditing({ index: i })}
            >
              <span>{e.name}</span>
              <span>{refersToText(e.ranges)}</span>
            </button>
          ))}
        </div>
        <div className="custom-views-buttons">
          <button type="button" className="dialog-btn" onClick={() => setEditing({ index: null })}>
            <u>N</u>ew...
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={!current}
            onClick={() => setEditing({ index: chosen })}
          >
            <u>M</u>odify...
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={!current}
            onClick={() => {
              setList(list.filter((_, i) => i !== chosen))
              setChosen(Math.max(0, chosen - 1))
            }}
          >
            <u>D</u>elete
          </button>
        </div>
      </div>
      <div className="edit-ranges-protect">
        <button
          type="button"
          className="dialog-btn"
          onClick={() => {
            apply()
            onProtect()
          }}
        >
          <u>P</u>rotect Sheet...
        </button>
      </div>
    </Dialog>
  )
}

/** New Range / Modify Range: a title, the cells, and an optional password (asked twice, as Excel does). */
function EditRangeDialog({
  controller,
  list,
  index,
  onCancel,
  onDone,
}: {
  controller: SheetController
  list: readonly EditRange[]
  index: number | null
  onCancel: () => void
  onDone: (range: EditRange) => void
}) {
  const old = index === null ? null : list[index]
  const [name, setName] = useState(old?.name ?? nextEditRangeName(list))
  const [refers, setRefers] = useState(refersToText(old?.ranges ?? controller.selection.ranges))
  // An existing password shows as a placeholder; typing replaces it, clearing it removes it.
  const kept = old && (old.password || old.hash) ? "********" : ""
  const [password, setPassword] = useState(kept)
  const [confirming, setConfirming] = useState(false)

  const finish = () => {
    const ranges = parseRefersTo(refers)!
    const range: EditRange = { name: name.trim(), ranges }
    if (old?.securityDescriptor) range.securityDescriptor = old.securityDescriptor
    if (password && password === kept && old) {
      if (old.password) range.password = old.password
      if (old.hash) range.hash = old.hash
    } else if (password) range.password = hashPassword(password)
    onDone(range)
  }

  if (confirming)
    return (
      <ConfirmPassword
        password={password}
        onBack={() => setConfirming(false)}
        onMismatch={() => controller.shell.alert(SHEET_MESSAGES.passwordMismatch)}
        onConfirmed={finish}
      />
    )

  return (
    <Dialog
      title={old ? "Modify Range" : "New Range"}
      width={360}
      onCancel={onCancel}
      onAccept={() => {
        const title = name.trim()
        if (!title) return controller.shell.alert(EDIT_RANGE_MESSAGES.noTitle)
        if (list.some((e, i) => i !== index && e.name.toLowerCase() === title.toLowerCase()))
          return controller.shell.alert(EDIT_RANGE_MESSAGES.takenTitle)
        if (!parseRefersTo(refers)) return controller.shell.alert(EDIT_RANGE_MESSAGES.badReference)
        if (password && password !== kept) setConfirming(true)
        else finish()
      }}
    >
      <label className="dialog-label" htmlFor="edit-range-title">
        <u>T</u>itle:
      </label>
      <input
        id="edit-range-title"
        className="dialog-field"
        data-autofocus
        value={name}
        onChange={(e) => setName(e.target.value)}
      />
      <label className="dialog-label" htmlFor="edit-range-refers">
        <u>R</u>efers to cells:
      </label>
      <input
        id="edit-range-refers"
        className="dialog-field"
        value={refers}
        onChange={(e) => setRefers(e.target.value)}
      />
      <label className="dialog-label" htmlFor="edit-range-password">
        Range <u>p</u>assword:
      </label>
      <input
        id="edit-range-password"
        className="dialog-field"
        type="password"
        value={password}
        onFocus={(e) => e.target.select()}
        onChange={(e) => setPassword(e.target.value)}
      />
    </Dialog>
  )
}

/** Unlock Range: a cell of a range with a password is about to change. */
export function UnlockRangeDialog({
  controller,
  name,
  onUnlocked,
  onClose,
}: {
  controller: SheetController
  name: string
  onUnlocked: () => void
  onClose: () => void
}) {
  const [password, setPassword] = useState("")
  return (
    <Dialog
      title="Unlock Range"
      width={360}
      onCancel={onClose}
      onAccept={() => {
        if (!controller.editRanges.unlock(name, password)) {
          controller.shell.alert(EDIT_RANGE_MESSAGES.wrongPassword)
          return
        }
        onClose()
        onUnlocked()
      }}
    >
      <p className="dialog-note">A cell you are trying to change is password protected.</p>
      <label className="dialog-label" htmlFor="unlock-range-password">
        <u>E</u>nter the password to change this cell:
      </label>
      <input
        id="unlock-range-password"
        className="dialog-field"
        type="password"
        data-autofocus
        value={password}
        onChange={(e) => setPassword(e.target.value)}
      />
    </Dialog>
  )
}
