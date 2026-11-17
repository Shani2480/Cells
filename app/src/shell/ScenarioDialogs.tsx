import { useState, useSyncExternalStore } from "react"
import {
  MAX_SCENARIO_CELLS,
  columnName,
  copyScenarios,
  parseRange,
  scenarioSummary,
  scenarioSummaryCells,
  showScenarioOp,
  type AxisEntry,
  type CellValue,
  type Op,
  type Range,
  type Scenario,
  type ScenarioSet,
} from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { authorName } from "@/settings/settings"
import { Dialog } from "./Dialog"

const abs = (row: number, col: number) => `$${columnName(col)}$${row + 1}`

/** "$B$1:$B$3,$D$4" for the changing cells of a scenario (runs down a column joined into ranges). */
export function cellsText(cells: { row: number; col: number }[]): string {
  const parts: string[] = []
  for (let i = 0; i < cells.length; i++) {
    const start = cells[i]
    let end = start
    while (i + 1 < cells.length && cells[i + 1].col === start.col && cells[i + 1].row === end.row + 1) end = cells[++i]
    parts.push(end === start ? abs(start.row, start.col) : `${abs(start.row, start.col)}:${abs(end.row, end.col)}`)
  }
  return parts.join(",")
}

/** "B1,B3:B4" or "Sheet1!B1" on the active sheet, cell by cell. Null when a part isn't a reference here. */
export function parseCells(controller: SheetController, text: string): { row: number; col: number }[] | null {
  const out: { row: number; col: number }[] = []
  const seen = new Set<string>()
  for (const raw of text.split(/[,;]/)) {
    const part = raw.trim().replace(/^=/, "")
    if (!part) continue
    const bang = part.lastIndexOf("!")
    if (bang >= 0) {
      const name = part.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'")
      if (name.toLowerCase() !== controller.sheet.name.toLowerCase()) return null
    }
    const range = parseRange(part.slice(bang + 1).replace(/\$/g, ""))
    if (!range) return null
    for (let row = range.r0; row <= range.r1; row++)
      for (let col = range.c0; col <= range.c1; col++) {
        if (seen.has(`${row},${col}`)) continue
        seen.add(`${row},${col}`)
        out.push({ row, col })
        if (out.length > MAX_SCENARIO_CELLS) return out
      }
  }
  return out.length ? out : null
}

function rangesText(ranges: readonly Range[]): string {
  return ranges
    .map((r) => (r.r0 === r.r1 && r.c0 === r.c1 ? abs(r.r0, r.c0) : `${abs(r.r0, r.c0)}:${abs(r.r1, r.c1)}`))
    .join(",")
}

/** A value typed in Scenario Values: a number when it reads as one, else the text. */
function valueOf(text: string): CellValue {
  const t = text.trim()
  if (t === "") return ""
  const n = Number(t.replace(",", "."))
  if (Number.isFinite(n) && /^[-+]?[\d.,]+(e[-+]?\d+)?$/i.test(t)) return n
  if (/^(true|false)$/i.test(t)) return t.toUpperCase() === "TRUE"
  return text
}

const textOf = (v: CellValue | undefined): string =>
  v === undefined ? "" : typeof v === "object" ? v.error : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v)

const today = () => new Date().toLocaleDateString()

type Sub =
  | { kind: "edit"; index: number | null }
  | { kind: "values"; draft: Scenario; index: number | null }
  | { kind: "summary" }
  | { kind: "merge" }
  | null

/**
 * Data › What-If Analysis › Scenario Manager: named sets of values for changing cells, to show on the sheet in turn,
 * compare in a Scenario Summary, or merge from another sheet. Kept per sheet, as in Excel.
 */
export function ScenarioManagerDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  useSyncExternalStore(controller.subscribe, controller.getVersion)
  const set = controller.sheet.scenarios
  const protection = controller.sheet.protection
  // On a protected sheet (without "Edit scenarios"), hidden scenarios aren't listed and locked ones can't change.
  const guarded = protection !== null && !protection.editScenarios
  const listed = set.scenarios.map((s, i) => ({ s, i })).filter(({ s }) => !(guarded && s.hidden))
  const [chosen, setChosen] = useState<number | null>(set.show ?? listed[0]?.i ?? null)
  const [sub, setSub] = useState<Sub>(null)
  const current = chosen !== null && listed.some((l) => l.i === chosen) ? set.scenarios[chosen] : null
  const frozen = current !== null && guarded && current.locked

  const save = (next: ScenarioSet, label: string, extra: Op[] = []) => {
    const sel = controller.selection
    controller.transact([...extra, { kind: "scenarios", sheet: controller.sheetIndex, scenarios: next }], {
      label,
      before: sel,
      after: sel,
    })
  }

  const show = () => {
    if (chosen === null || !current) return
    const next = { ...copyScenarios(set), current: chosen, show: chosen }
    save(next, "Show Scenario", [showScenarioOp(controller.sheet, controller.sheetIndex, current)])
  }

  const remove = () => {
    if (chosen === null || frozen) return
    const next = copyScenarios(set)
    next.scenarios.splice(chosen, 1)
    const fix = (i: number | null) => (i === null || i === chosen ? null : i > chosen ? i - 1 : i)
    next.current = fix(next.current)
    next.show = fix(next.show)
    save(next, "Delete Scenario")
    setChosen(next.scenarios.length ? Math.min(chosen, next.scenarios.length - 1) : null)
  }

  const keep = (draft: Scenario, index: number | null) => {
    const next = copyScenarios(set)
    if (index === null) next.scenarios.push(draft)
    else next.scenarios[index] = draft
    const at = index ?? next.scenarios.length - 1
    next.show = at
    save(next, index === null ? "Add Scenario" : "Edit Scenario")
    setChosen(at)
    setSub(null)
  }

  return (
    <>
      <Dialog
        title="Scenario Manager"
        onCancel={onClose}
        onAccept={onClose}
        width={460}
        actions={
          <>
            <button type="button" className="dialog-btn" disabled={!current} onClick={show}>
              <u>S</u>how
            </button>
            <button type="button" className="dialog-btn is-primary" onClick={onClose}>
              Close
            </button>
          </>
        }
      >
        <div className="scenario-manager">
          <div className="scenario-list-column">
            <span className="dialog-label" id="scenario-list-label">
              Sc<u>e</u>narios:
            </span>
            <div className="dialog-list scenario-list" role="listbox" aria-labelledby="scenario-list-label">
              {listed.map(({ s, i }) => (
                <button
                  key={i}
                  type="button"
                  role="option"
                  aria-selected={i === chosen}
                  className={`dialog-list-item${i === chosen ? " is-selected" : ""}`}
                  onClick={() => setChosen(i)}
                  onDoubleClick={() => {
                    setChosen(i)
                    const next = { ...copyScenarios(set), current: i, show: i }
                    save(next, "Show Scenario", [showScenarioOp(controller.sheet, controller.sheetIndex, s)])
                  }}
                >
                  {s.name}
                </button>
              ))}
              {!listed.length ? (
                <p className="dialog-hint scenario-empty">No Scenarios defined. Choose Add to add scenarios.</p>
              ) : null}
            </div>
          </div>
          <div className="scenario-buttons">
            <button type="button" className="dialog-btn" onClick={() => setSub({ kind: "edit", index: null })}>
              <u>A</u>dd...
            </button>
            <button type="button" className="dialog-btn" disabled={!current || frozen} onClick={remove}>
              <u>D</u>elete
            </button>
            <button
              type="button"
              className="dialog-btn"
              disabled={!current || frozen}
              onClick={() => chosen !== null && setSub({ kind: "edit", index: chosen })}
            >
              <u>E</u>dit...
            </button>
            <button type="button" className="dialog-btn" onClick={() => setSub({ kind: "merge" })}>
              <u>M</u>erge...
            </button>
            <button
              type="button"
              className="dialog-btn"
              disabled={!set.scenarios.length}
              onClick={() => setSub({ kind: "summary" })}
            >
              S<u>u</u>mmary...
            </button>
          </div>
        </div>
        <div className="scenario-details">
          <span className="dialog-label">Changing cells:</span>
          <span className="scenario-detail">{current ? cellsText(current.cells) : ""}</span>
          <span className="dialog-label">Comment:</span>
          <span className="scenario-detail scenario-comment">{current?.comment ?? ""}</span>
        </div>
      </Dialog>
      {sub?.kind === "edit" ? (
        <ScenarioEditDialog
          controller={controller}
          editing={sub.index === null ? null : set.scenarios[sub.index]}
          names={set.scenarios.filter((_, i) => i !== sub.index).map((s) => s.name)}
          onCancel={() => setSub(null)}
          onNext={(draft) => setSub({ kind: "values", draft, index: sub.index })}
        />
      ) : null}
      {sub?.kind === "values" ? (
        <ScenarioValuesDialog
          draft={sub.draft}
          onCancel={() => setSub(null)}
          onDone={(draft) => keep(draft, sub.index)}
        />
      ) : null}
      {sub?.kind === "summary" ? (
        <ScenarioSummaryDialog
          controller={controller}
          onCancel={() => setSub(null)}
          onDone={() => {
            setSub(null)
            onClose()
          }}
        />
      ) : null}
      {sub?.kind === "merge" ? (
        <ScenarioMergeDialog
          controller={controller}
          onCancel={() => setSub(null)}
          onMerge={(more) => {
            const next = copyScenarios(set)
            const taken = new Set(next.scenarios.map((s) => s.name.toLowerCase()))
            for (const s of more) {
              let name = s.name
              // Excel adds the date to a name that's already taken.
              if (taken.has(name.toLowerCase())) name = `${s.name} ${today()}`
              for (let n = 2; taken.has(name.toLowerCase()); n++) name = `${s.name} ${today()} (${n})`
              taken.add(name.toLowerCase())
              next.scenarios.push({ ...s, name, cells: s.cells.map((c) => ({ ...c })) })
            }
            save(next, "Merge Scenarios")
            setSub(null)
          }}
        />
      ) : null}
    </>
  )
}

/** Add Scenario / Edit Scenario: the name, changing cells, comment and protection. */
function ScenarioEditDialog({
  controller,
  editing,
  names,
  onCancel,
  onNext,
}: {
  controller: SheetController
  editing: Scenario | null
  names: string[]
  onCancel: () => void
  onNext: (draft: Scenario) => void
}) {
  const author = authorName()
  const [name, setName] = useState(editing?.name ?? "")
  const [cellText, setCellText] = useState(() => {
    if (editing) return cellsText(editing.cells)
    const r = controller.selection.ranges
    return r
      .map((x) => (x.r0 === x.r1 && x.c0 === x.c1 ? abs(x.r0, x.c0) : `${abs(x.r0, x.c0)}:${abs(x.r1, x.c1)}`))
      .join(",")
  })
  const [comment, setComment] = useState(
    editing
      ? `${editing.comment}${editing.comment ? "\n" : ""}Modified by ${author} on ${today()}`
      : `Created by ${author} on ${today()}`,
  )
  const [locked, setLocked] = useState(editing?.locked ?? true)
  const [hidden, setHidden] = useState(editing?.hidden ?? false)

  const next = () => {
    const trimmed = name.trim()
    if (!trimmed) return controller.shell.alert("Scenario name must be specified.")
    if (names.some((n) => n.toLowerCase() === trimmed.toLowerCase()))
      return controller.shell.alert("Scenario name already exists. Please type a different name.")
    const cells = parseCells(controller, cellText)
    if (!cells) return controller.shell.alert("Reference must be on the active sheet.")
    if (cells.length > MAX_SCENARIO_CELLS)
      return controller.shell.alert(`Too many changing cells. Up to ${MAX_SCENARIO_CELLS} changing cells are allowed.`)
    const go = () =>
      onNext({
        name: trimmed,
        comment,
        user: editing?.user || author,
        locked,
        hidden,
        cells: cells.map(({ row, col }) => ({
          row,
          col,
          value:
            editing?.cells.find((c) => c.row === row && c.col === col)?.value ??
            controller.sheet.cells.value(row, col) ??
            "",
        })),
      })
    const formulas = cells.some(({ row, col }) => controller.sheet.cells.input(row, col)?.startsWith("="))
    if (formulas)
      return controller.shell.alert(
        "One or more cells contain formulas that will be replaced by constant values when scenario is shown.",
        go,
      )
    go()
  }

  return (
    <Dialog title={editing ? "Edit Scenario" : "Add Scenario"} onCancel={onCancel} onAccept={next} width={420}>
      <div className="scenario-edit">
        <label className="dialog-label" htmlFor="scenario-name">
          Scenario <u>n</u>ame:
        </label>
        <input
          id="scenario-name"
          className="dialog-field"
          autoFocus
          maxLength={255}
          value={name}
          onChange={(e) => setName(e.target.value)}
        />
        <label className="dialog-label" htmlFor="scenario-cells">
          Changing <u>c</u>ells:
        </label>
        <input
          id="scenario-cells"
          className="dialog-field"
          value={cellText}
          onChange={(e) => setCellText(e.target.value)}
        />
        <p className="dialog-hint">Ctrl+click cells to select non-adjacent changing cells.</p>
        <label className="dialog-label" htmlFor="scenario-comment">
          C<u>o</u>mment:
        </label>
        <textarea
          id="scenario-comment"
          className="dialog-field scenario-comment-field"
          rows={3}
          value={comment}
          onChange={(e) => setComment(e.target.value)}
        />
        <fieldset className="dialog-group">
          <legend>Protection</legend>
          <label className="dialog-radio">
            <input type="checkbox" checked={locked} onChange={(e) => setLocked(e.target.checked)} />
            <span>
              <u>P</u>revent changes
            </span>
          </label>
          <label className="dialog-radio">
            <input type="checkbox" checked={hidden} onChange={(e) => setHidden(e.target.checked)} />
            <span>
              Hi<u>d</u>e
            </span>
          </label>
        </fieldset>
      </div>
    </Dialog>
  )
}

/** Scenario Values: one field per changing cell, filled with the values it has now (or had in the scenario). */
function ScenarioValuesDialog({
  draft,
  onCancel,
  onDone,
}: {
  draft: Scenario
  onCancel: () => void
  onDone: (draft: Scenario) => void
}) {
  const [texts, setTexts] = useState(() => draft.cells.map((c) => textOf(c.value)))
  const accept = () =>
    onDone({ ...draft, cells: draft.cells.map((c, i) => ({ row: c.row, col: c.col, value: valueOf(texts[i]) })) })
  return (
    <Dialog title="Scenario Values" onCancel={onCancel} onAccept={accept} width={340}>
      <p className="dialog-label">Enter values for each of the changing cells.</p>
      <div className="scenario-values">
        {draft.cells.map((c, i) => {
          const id = `scenario-value-${i}`
          return (
            <div key={`${c.row},${c.col}`} className="scenario-value-row">
              <label className="dialog-label" htmlFor={id}>
                {i + 1}: {abs(c.row, c.col)}
              </label>
              <input
                id={id}
                className="dialog-field"
                autoFocus={i === 0}
                aria-label={abs(c.row, c.col)}
                value={texts[i]}
                onChange={(e) => setTexts(texts.map((t, k) => (k === i ? e.target.value : t)))}
              />
            </div>
          )
        })}
      </div>
    </Dialog>
  )
}

/** Scenario Summary: a new sheet comparing the scenarios' changing cells and the result cells with each. */
function ScenarioSummaryDialog({
  controller,
  onCancel,
  onDone,
}: {
  controller: SheetController
  onCancel: () => void
  onDone: () => void
}) {
  const set = controller.sheet.scenarios
  const [results, setResults] = useState(() =>
    rangesText(set.resultCells.length ? set.resultCells : controller.selection.ranges),
  )
  const accept = () => {
    const ranges: Range[] = []
    for (const part of results.split(/[,;]/)) {
      if (!part.trim()) continue
      const cells = parseCells(controller, part)
      const r = parseRange(part.trim().replace(/^=/, "").replace(/^.*!/, "").replace(/\$/g, ""))
      if (!cells || !r) return controller.shell.alert("Reference must be on the active sheet.")
      ranges.push(r)
    }
    const source = controller.sheetIndex
    const summary = scenarioSummary(controller.wb, source, ranges)
    const laid = scenarioSummaryCells(summary)
    const names = controller.wb.sheets.map((s) => s.name.toLowerCase())
    let name = "Scenario Summary"
    for (let n = 2; names.includes(name.toLowerCase()); n++) name = `Scenario Summary ${n}`
    // Excel puts the report before the sheet and shows it.
    const at = source
    const after = { ...copyScenarios(set), resultCells: ranges }
    const ops: Op[] = [
      { kind: "scenarios", sheet: source, scenarios: after },
      { kind: "addSheet", sheet: at, name },
      { kind: "setCells", sheet: at, cells: laid.cells },
      { kind: "format", sheet: at, ranges: laid.headings, patch: { bold: true } },
      { kind: "format", sheet: at, ranges: [laid.headings[0]], patch: { fill: "#808080", fontColor: "#FFFFFF" } },
      {
        kind: "format",
        sheet: at,
        ranges: [{ r0: 3, c0: 4, r1: 2 + summary.changing.length, c1: 3 + summary.names.length }],
        patch: { fill: "#D9D9D9" },
      },
    ]
    const entry = controller.transact(ops, { label: "Scenario Summary", before: controller.selection, sheetAfter: at })
    if (!entry) return
    controller.activate(at)
    // The table's columns fit their contents (the notes under it run over, as in Excel).
    const fit = controller.columnFit
    const widths: AxisEntry[] = []
    for (let col = laid.table.c0; fit && col <= laid.table.c1; col++) {
      const size = fit(col, { r0: laid.table.r0, r1: laid.table.r1 })
      if (size !== null && size > controller.sheet.cols.size(col)) widths.push({ index: col, size, hidden: false })
    }
    if (widths.length) controller.wb.amend(entry, [{ kind: "axis", sheet: at, axis: "cols", entries: widths }])
    onDone()
  }
  return (
    <Dialog title="Scenario Summary" onCancel={onCancel} onAccept={accept} width={320}>
      <fieldset className="dialog-group">
        <legend>Report type</legend>
        <label className="dialog-radio">
          <input type="radio" name="scenario-report" defaultChecked />
          <span>
            Scenario <u>s</u>ummary
          </span>
        </label>
        <label className="dialog-radio" title="PivotTables come with a later version.">
          <input type="radio" name="scenario-report" disabled />
          <span>
            Scenario <u>P</u>ivotTable report
          </span>
        </label>
      </fieldset>
      <label className="dialog-label" htmlFor="scenario-results">
        <u>R</u>esult cells:
      </label>
      <input
        id="scenario-results"
        className="dialog-field"
        autoFocus
        value={results}
        onChange={(e) => setResults(e.target.value)}
      />
    </Dialog>
  )
}

/** Merge Scenarios: copies the scenarios of another sheet of this workbook. */
function ScenarioMergeDialog({
  controller,
  onCancel,
  onMerge,
}: {
  controller: SheetController
  onCancel: () => void
  onMerge: (scenarios: Scenario[]) => void
}) {
  const sheets = controller.wb.sheets
  const others = sheets.map((s, i) => ({ s, i })).filter(({ i }) => i !== controller.sheetIndex)
  const [from, setFrom] = useState<number | null>(others[0]?.i ?? null)
  const count = from === null ? 0 : sheets[from].scenarios.scenarios.length
  return (
    <Dialog
      title="Merge Scenarios"
      onCancel={onCancel}
      onAccept={() => (from === null ? onCancel() : onMerge(sheets[from].scenarios.scenarios))}
      width={320}
    >
      <span className="dialog-label">
        Merge scenarios from <u>s</u>heet:
      </span>
      <div className="dialog-list" role="listbox" aria-label="Sheets">
        {others.map(({ s, i }) => (
          <button
            key={i}
            type="button"
            role="option"
            aria-selected={i === from}
            className={`dialog-list-item${i === from ? " is-selected" : ""}`}
            onClick={() => setFrom(i)}
          >
            {s.name}
          </button>
        ))}
      </div>
      <p className="dialog-hint">
        {from === null
          ? "This workbook has no other sheets."
          : `There ${count === 1 ? "is 1 scenario" : `are ${count === 0 ? "no" : count} scenarios`} on source sheet.`}
      </p>
    </Dialog>
  )
}
