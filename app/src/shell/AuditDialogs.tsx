import { useEffect, useMemo, useState, useSyncExternalStore } from "react"
import {
  formatValue,
  isError,
  nextPart,
  parseFormula,
  quoteSheet,
  replacePart,
  stepText,
  valueNode,
  type CellValue,
  type DefinedName,
  type Node,
} from "@cells/core"
import { ERROR_HELP } from "@/sheet/audit"
import type { SheetController } from "@/sheet/controller"
import { absoluteRef, type CreateEdges } from "@/sheet/names"
import { DefineNameDialog } from "./DefineNameDialog"
import { Dialog } from "./Dialog"

/** A value as a dialog lists it: text in quotes, as Excel's Name Manager and Watch Window show values. */
export function valueText(v: CellValue | undefined | null): string {
  if (v === undefined || v === null) return ""
  if (isError(v)) return v.error
  if (typeof v === "string") return `"${v}"`
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE"
  return String(v)
}

/** What a name stands for now: the cell's value, or the values of its range in braces, as Excel shows them. */
function nameValue(c: SheetController, name: DefinedName): string {
  const target = c.names.resolve(name.name)
  if (!target) {
    const parsed = parseFormula(name.ref)
    if (parsed.ok && (parsed.ast.type === "number" || parsed.ast.type === "string" || parsed.ast.type === "bool"))
      return name.ref
    return ""
  }
  const sheet = c.wb.sheets[target.sheet]
  const { r0, c0 } = target.range
  const r1 = Math.min(target.range.r1, r0 + 49)
  const c1 = Math.min(target.range.c1, c0 + 49)
  if (r0 === r1 && c0 === c1) return valueText(sheet.cells.value(r0, c0))
  const rows: string[] = []
  for (let r = r0; r <= r1; r++) {
    const row: string[] = []
    for (let col = c0; col <= c1; col++) row.push(valueText(sheet.cells.value(r, col)))
    rows.push(row.join(","))
  }
  return `{${rows.join(";")}}`
}

/**
 * Formulas › Name Manager: the workbook's names with their values, what they refer to, scope and comment; New, Edit
 * and Delete, and the selected name's Refers To changed in place.
 */
export function NameManagerDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const names = controller.names.all()
  const [chosen, setChosen] = useState<string | null>(names[0]?.name ?? null)
  const [editor, setEditor] = useState<"new" | DefinedName | null>(null)
  const current = names.find((n) => n.name === chosen) ?? null
  const [refersTo, setRefersTo] = useState<string | null>(null)
  const shownRef = refersTo ?? (current ? "=" + current.ref : "")
  const commitRef = () => {
    if (!current || refersTo === null) return
    const problem = controller.names.define(current.name, refersTo, current.comment ?? "", current.name)
    if (problem) controller.shell.alert(problem)
    setRefersTo(null)
  }
  return (
    <>
      <Dialog
        title="Name Manager"
        onCancel={onClose}
        onAccept={onClose}
        width={640}
        acceptLabel="Close"
        cancelLabel={null}
      >
        <div className="cf-manage-toolbar">
          <button type="button" className="dialog-btn" onClick={() => setEditor("new")}>
            <u>N</u>ew...
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={!current}
            onClick={() => current && setEditor(current)}
          >
            <u>E</u>dit...
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={!current}
            onClick={() =>
              current &&
              controller.shell.confirm(`Are you sure you want to delete the name ${current.name}?`, () => {
                controller.names.remove([current.name])
                setChosen(null)
              })
            }
          >
            <u>D</u>elete
          </button>
        </div>
        <div className="cf-manage-table name-manager-table" role="grid" aria-label="Names">
          <div className="cf-manage-head" role="row">
            <span role="columnheader">Name</span>
            <span role="columnheader">Value</span>
            <span role="columnheader">Refers To</span>
            <span role="columnheader">Scope</span>
            <span role="columnheader">Comment</span>
          </div>
          {names.map((n) => (
            <div
              key={n.name}
              role="row"
              aria-selected={n.name === chosen}
              className={`cf-manage-row${n.name === chosen ? " is-selected" : ""}`}
              onClick={() => {
                setChosen(n.name)
                setRefersTo(null)
              }}
              onDoubleClick={() => setEditor(n)}
            >
              <span role="gridcell">{n.name}</span>
              <span role="gridcell" className="name-manager-value">
                {nameValue(controller, n)}
              </span>
              <span role="gridcell">={n.ref}</span>
              <span role="gridcell">Workbook</span>
              <span role="gridcell">{n.comment ?? ""}</span>
            </div>
          ))}
          {!names.length ? <p className="dialog-label name-manager-empty">This workbook has no names yet.</p> : null}
        </div>
        <label className="dialog-label" htmlFor="name-manager-ref">
          <u>R</u>efers to:
        </label>
        <input
          id="name-manager-ref"
          className="dialog-field"
          disabled={!current}
          value={shownRef}
          onChange={(e) => setRefersTo(e.target.value)}
          onBlur={commitRef}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault()
              e.stopPropagation()
              commitRef()
            }
          }}
        />
      </Dialog>
      {editor ? (
        <DefineNameDialog
          controller={controller}
          editing={editor === "new" ? undefined : editor}
          onClose={() => setEditor(null)}
        />
      ) : null}
    </>
  )
}

/** Formulas › Create from Selection: names for the selection's rows and columns, from the labels on its edges. */
export function CreateNamesDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [edges, setEdges] = useState<CreateEdges>(() => controller.names.guessEdges())
  const accept = () => {
    const problem = controller.names.createFromSelection(edges)
    if (problem) {
      controller.shell.alert(problem)
      return
    }
    onClose()
  }
  const box = (key: keyof CreateEdges, label: string, letter: string) => (
    <label className="dialog-radio">
      <input type="checkbox" checked={edges[key]} onChange={(e) => setEdges({ ...edges, [key]: e.target.checked })} />
      <span>
        {label.split(letter)[0]}
        <u>{letter}</u>
        {label.split(letter).slice(1).join(letter)}
      </span>
    </label>
  )
  return (
    <Dialog title="Create Names from Selection" onCancel={onClose} onAccept={accept} width={300}>
      <fieldset className="dialog-group">
        <legend>Create names from values in the:</legend>
        {box("top", "Top row", "T")}
        {box("left", "Left column", "L")}
        {box("bottom", "Bottom row", "B")}
        {box("right", "Right column", "R")}
      </fieldset>
    </Dialog>
  )
}

/** The cell's address as the auditing dialogs show it: `Sheet1!$B$5`. */
function cellRef(c: SheetController, row: number, col: number) {
  return `${quoteSheet(c.sheet.name)}!${absoluteRef({ r0: row, c0: col, r1: row, c1: col })}`
}

/**
 * Formulas › Evaluate Formula: the active cell's formula worked out one part at a time. The part evaluated next is
 * underlined; Evaluate replaces it with its value, until only the result is left, then Restart. (Excel's Step In and
 * Step Out, which open a referenced cell's own formula, aren't here yet.)
 */
export function EvaluateFormulaDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const at = controller.selection.active
  const input = controller.sheet.cells.input(at.row, at.col)
  const start = useMemo<Node | null>(() => {
    if (!input?.startsWith("=")) return null
    const parsed = parseFormula(input.slice(1))
    if (!parsed.ok) return null
    // The window's own engine is idle while a worker calculates: bring it up to date with the values shown.
    controller.wb.calc.rebuild(true)
    return parsed.ast
  }, [controller, input])
  const [root, setRoot] = useState<Node | null>(start)
  if (!start || !root) {
    return (
      <Dialog title="Evaluate Formula" onCancel={onClose} onAccept={onClose} cancelLabel={null} acceptLabel="Close">
        <p className="dialog-label">The active cell doesn't contain a formula.</p>
      </Dialog>
    )
  }
  const part = nextPart(root)
  const { text, start: s0, end: s1 } = stepText(root, part)
  const evaluate = () => {
    if (!part) {
      setRoot(start)
      return
    }
    const value = controller.wb.calc.evaluatePart(part, controller.sheetIndex, at.row, at.col)
    setRoot(replacePart(root, part, valueNode(value ?? { error: "#VALUE!" })))
  }
  return (
    <Dialog
      title="Evaluate Formula"
      onCancel={onClose}
      onAccept={evaluate}
      width={520}
      actions={
        <>
          <button type="button" className="dialog-btn is-primary" onClick={evaluate} data-autofocus>
            {part ? (
              <>
                <u>E</u>valuate
              </>
            ) : (
              <>
                R<u>e</u>start
              </>
            )}
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            <u>C</u>lose
          </button>
        </>
      }
    >
      <div className="evaluate-grid">
        <span className="dialog-label">Reference:</span>
        <span className="dialog-label">
          E<u>v</u>aluation:
        </span>
        <span className="evaluate-ref">{cellRef(controller, at.row, at.col)}</span>
        <div className="evaluate-box" aria-label="Evaluation" role="textbox" aria-readonly>
          ={text.slice(0, s0)}
          {part ? <u className="evaluate-next">{text.slice(s0, s1)}</u> : null}
          {text.slice(s1)}
        </div>
      </div>
      <p className="dialog-hint">
        {part
          ? "To show the result of the underlined expression, click Evaluate. The most recent result appears italicized."
          : "The formula is worked out. Restart starts again from the formula."}
      </p>
    </Dialog>
  )
}

/**
 * Formulas › Error Checking: steps through the sheet's formulas that show errors and its numbers stored as text,
 * from the active cell on, explaining each; the sheet stays usable while it's open, as in Excel.
 */
export function ErrorCheckingDialog({
  controller,
  onClose,
  onEvaluate,
}: {
  controller: SheetController
  onClose: () => void
  /** Show Calculation Steps: Evaluate Formula on the cell. */
  onEvaluate: () => void
}) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const [at, setAt] = useState(() => {
    const a = controller.selection.active
    if (controller.audit.problemAt(a.row, a.col)) return { row: a.row, col: a.col }
    return controller.audit.nextProblem(a)
  })
  const problem = at ? controller.audit.problemAt(at.row, at.col) : null
  // Excel selects the cell it explains.
  useEffect(() => {
    if (at) controller.goTo({ r0: at.row, c0: at.col, r1: at.row, c1: at.col })
    // Only when it opens: Next and Previous select theirs.
  }, [])
  const go = (back: boolean) => {
    const next = controller.audit.nextProblem(at ?? controller.selection.active, back)
    if (!next || (at && next.row === at.row && next.col === at.col && !controller.audit.problemAt(at.row, at.col))) {
      controller.shell.alert("The error check is complete for the entire sheet.", onClose)
      return
    }
    setAt(next)
    controller.goTo({ r0: next.row, c0: next.col, r1: next.row, c1: next.col })
  }
  if (!at || !problem) {
    return (
      <Dialog title="Error Checking" onCancel={onClose} onAccept={onClose} cancelLabel={null} modeless>
        <p className="dialog-label">The error check is complete for the entire sheet.</p>
      </Dialog>
    )
  }
  const [heading, help] =
    problem.kind === "error"
      ? (ERROR_HELP[problem.code] ?? ["Error", "The formula results in an error."])
      : ["Number Stored as Text", "The number in this cell is formatted as text or preceded by an apostrophe."]
  const data = controller.sheet.cells.get(at.row, at.col)
  const shown = data?.input ?? formatValue(data?.value, controller.wb.styles.get(data?.style ?? 0).numberFormat) ?? ""
  const after = () => go(false)
  return (
    <Dialog
      title="Error Checking"
      onCancel={onClose}
      onAccept={after}
      width={520}
      modeless
      actions={
        <>
          <button type="button" className="dialog-btn" onClick={() => go(true)}>
            <u>P</u>revious
          </button>
          <button type="button" className="dialog-btn is-primary" onClick={after}>
            <u>N</u>ext
          </button>
        </>
      }
    >
      <div className="error-check">
        <div className="error-check-text">
          <strong>
            Error in cell {absoluteRef({ r0: at.row, c0: at.col, r1: at.row, c1: at.col }).replace(/\$/g, "")}
          </strong>
          <code className="error-check-formula">{shown}</code>
          <strong>{heading}</strong>
          <p>{help}</p>
        </div>
        <div className="error-check-buttons">
          {problem.kind === "error" ? (
            <button
              type="button"
              className="dialog-btn"
              onClick={() => {
                controller.goTo({ r0: at.row, c0: at.col, r1: at.row, c1: at.col })
                onEvaluate()
              }}
            >
              Show <u>C</u>alculation Steps...
            </button>
          ) : (
            <button
              type="button"
              className="dialog-btn"
              onClick={() => (controller.audit.convertToNumber(at), after())}
            >
              <u>C</u>onvert to Number
            </button>
          )}
          <button type="button" className="dialog-btn" onClick={() => (controller.audit.ignoreError(at), after())}>
            <u>I</u>gnore Error
          </button>
          <button
            type="button"
            className="dialog-btn"
            onClick={() => {
              onClose()
              controller.goTo({ r0: at.row, c0: at.col, r1: at.row, c1: at.col })
              controller.beginEdit("edit", undefined, "bar")
            }}
          >
            Edit in <u>F</u>ormula Bar
          </button>
        </div>
      </div>
    </Dialog>
  )
}
