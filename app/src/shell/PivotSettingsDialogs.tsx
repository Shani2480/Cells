import {
  autoDataFieldName,
  PIVOT_FN_LABEL,
  PIVOT_FNS,
  pivotFieldName,
  pivotOptions,
  type CellStyle,
  type PivotFieldSettings,
  type PivotFn,
  type PivotOptions,
  type PivotSubtotal,
} from "@cells/core"
import { useState, type ReactNode } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"
import { NumberTab } from "./format-cells/FormatCellsDialog"

/** Format Cells with its Number tab only, as Value Field Settings › Number Format opens it. */
export function NumberFormatDialog({
  controller,
  code,
  onAccept,
  onClose,
}: {
  controller: SheetController
  code: string | undefined
  onAccept: (code: string | undefined) => void
  onClose: () => void
}) {
  const [style, setStyle] = useState<CellStyle>(() => ({ ...controller.wb.styles.get(0), numberFormat: code }))
  return (
    <Dialog
      title="Format Cells"
      onCancel={onClose}
      onAccept={() => {
        onAccept(style.numberFormat)
        onClose()
      }}
      width={560}
    >
      <div className="fc-page" role="tabpanel" aria-label="Number">
        <NumberTab
          controller={controller}
          style={style}
          set={(fields) =>
            setStyle((s) => {
              const next = { ...s, ...(fields as CellStyle) }
              if (!next.numberFormat) delete next.numberFormat
              return next
            })
          }
        />
      </div>
    </Dialog>
  )
}

function Tabs<T extends string>({
  tabs,
  tab,
  setTab,
  children,
}: {
  tabs: { id: T; label: string }[]
  tab: T
  setTab: (t: T) => void
  children: ReactNode
}) {
  return (
    <>
      <div className="fc-tabs" role="tablist">
        {tabs.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            aria-selected={tab === t.id}
            className="fc-tab"
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>
      <div className="fc-page pivot-settings-page" role="tabpanel" aria-label={tabs.find((t) => t.id === tab)!.label}>
        {children}
      </div>
    </>
  )
}

function Check({
  label,
  checked,
  set,
  disabled,
}: {
  label: string
  checked: boolean
  set: (on: boolean) => void
  disabled?: boolean
}) {
  return (
    <label className="dialog-radio">
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => set(e.target.checked)} />
      <span>{label}</span>
    </label>
  )
}

function FnList({
  label,
  picked,
  pick,
  multiple,
  disabled,
}: {
  label: string
  picked: PivotFn[]
  pick: (fn: PivotFn) => void
  multiple?: boolean
  disabled?: boolean
}) {
  return (
    <div
      className={`dialog-list pivot-fn-list${disabled ? " is-disabled" : ""}`}
      role="listbox"
      aria-label={label}
      aria-multiselectable={multiple || undefined}
      aria-disabled={disabled || undefined}
    >
      {PIVOT_FNS.map((fn) => (
        <button
          key={fn}
          type="button"
          role="option"
          aria-selected={picked.includes(fn)}
          disabled={disabled}
          className={`dialog-list-item${picked.includes(fn) ? " is-selected" : ""}`}
          onClick={() => pick(fn)}
        >
          {PIVOT_FN_LABEL[fn]}
        </button>
      ))}
    </div>
  )
}

/**
 * Value Field Settings: the value field's Custom Name and how it summarizes (picking a function renames an automatic
 * name as Excel does), and its Number Format. Show Values As comes with step 5.
 */
export function ValueFieldSettingsDialog({
  controller,
  index,
  onClose,
}: {
  controller: SheetController
  index: number
  onClose: () => void
}) {
  const at = controller.pivots.activeWithCache()
  const v = at?.table.values[index]
  const [name, setName] = useState(v?.name ?? "")
  const [fn, setFn] = useState<PivotFn>(v?.fn ?? "sum")
  const [numFmt, setNumFmt] = useState(v?.numFmt)
  const [formatting, setFormatting] = useState(false)
  if (!at || !v) return null
  const { table, cache } = at
  const pick = (next: PivotFn) => {
    // An automatic name follows the function.
    if (name === autoDataFieldName(table, cache, v.field, fn)) setName(autoDataFieldName(table, cache, v.field, next))
    setFn(next)
  }
  const accept = () => {
    const problem = controller.pivots.setValueSettings(index, { name, fn, numFmt })
    if (problem) return controller.shell.alert(problem)
    onClose()
  }
  return (
    <>
      <Dialog
        title="Value Field Settings"
        onCancel={onClose}
        onAccept={accept}
        width={380}
        extra={{ label: "Number Format", onSelect: () => setFormatting(true) }}
      >
        <p className="dialog-label">Source Name: {pivotFieldName(table, cache, v.field)}</p>
        <div className="goal-seek-fields">
          <label className="dialog-label" htmlFor="pivot-value-name">
            Custom Name:
          </label>
          <input
            id="pivot-value-name"
            className="dialog-field"
            autoFocus
            value={name}
            onChange={(e) => setName(e.target.value)}
          />
        </div>
        <Tabs
          tabs={[{ id: "summarize" as const, label: "Summarize Values By" }]}
          tab="summarize"
          setTab={() => undefined}
        >
          <p className="dialog-label">Summarize value field by</p>
          <p className="dialog-hint">
            Choose the type of calculation that you want to use to summarize data from the selected field
          </p>
          <FnList label="Summarize value field by" picked={[fn]} pick={pick} />
        </Tabs>
      </Dialog>
      {formatting ? (
        <NumberFormatDialog
          controller={controller}
          code={numFmt}
          onAccept={setNumFmt}
          onClose={() => setFormatting(false)}
        />
      ) : null}
    </>
  )
}

/**
 * Field Settings for a row, column or report filter field: Custom Name; Subtotals (Automatic, None, Custom with one
 * or more functions); Layout: outline or tabular form, labels from the next field in the same column (compact),
 * subtotals at the top, repeat item labels, a blank line after each item, and show items with no data.
 */
export function FieldSettingsDialog({
  controller,
  field,
  onClose,
}: {
  controller: SheetController
  field: number
  onClose: () => void
}) {
  const at = controller.pivots.activeWithCache()
  const f = at?.table.fields[field]
  const [tab, setTab] = useState<"subtotals" | "layout">("subtotals")
  const [s, setS] = useState<PivotFieldSettings | null>(() =>
    at && f
      ? {
          name: pivotFieldName(at.table, at.cache, field),
          subtotals: f.subtotals,
          outline: f.outline,
          compact: f.compact,
          subtotalTop: f.subtotalTop,
          repeatLabels: f.repeatLabels,
          insertBlankRow: f.insertBlankRow,
          showAll: f.showAll,
        }
      : null,
  )
  if (!at || !s) return null
  const set = (patch: Partial<PivotFieldSettings>) => setS({ ...s, ...patch })
  const mode: "auto" | "none" | "custom" = !s.subtotals.length
    ? "none"
    : s.subtotals.length === 1 && s.subtotals[0] === "default"
      ? "auto"
      : "custom"
  const custom = s.subtotals.filter((x): x is PivotFn => x !== "default")
  const toggle = (fn: PivotFn) => {
    const next = custom.includes(fn) ? custom.filter((x) => x !== fn) : [...custom, fn]
    // Excel's order, whatever order they were picked in; none picked is no subtotals.
    set({ subtotals: PIVOT_FNS.filter((x) => next.includes(x)) as PivotSubtotal[] })
  }
  const accept = () => {
    const problem = controller.pivots.setFieldSettings(field, s)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }
  return (
    <Dialog title="Field Settings" onCancel={onClose} onAccept={accept} width={400}>
      <p className="dialog-label">Source Name: {at.cache.fields[field]?.name}</p>
      <div className="goal-seek-fields">
        <label className="dialog-label" htmlFor="pivot-field-name">
          Custom Name:
        </label>
        <input
          id="pivot-field-name"
          className="dialog-field"
          autoFocus
          value={s.name}
          onChange={(e) => set({ name: e.target.value })}
        />
      </div>
      <Tabs
        tabs={[
          { id: "subtotals", label: "Subtotals & Filters" },
          { id: "layout", label: "Layout & Print" },
        ]}
        tab={tab}
        setTab={setTab}
      >
        {tab === "subtotals" ? (
          <fieldset className="dialog-group">
            <legend>Subtotals</legend>
            <label className="dialog-radio">
              <input
                type="radio"
                name="pivot-subtotals"
                checked={mode === "auto"}
                onChange={() => set({ subtotals: ["default"] })}
              />
              <span>Automatic</span>
            </label>
            <label className="dialog-radio">
              <input
                type="radio"
                name="pivot-subtotals"
                checked={mode === "none"}
                onChange={() => set({ subtotals: [] })}
              />
              <span>None</span>
            </label>
            <label className="dialog-radio">
              <input
                type="radio"
                name="pivot-subtotals"
                checked={mode === "custom"}
                onChange={() => set({ subtotals: ["sum"] })}
              />
              <span>Custom</span>
            </label>
            <p className="dialog-hint">Select one or more functions:</p>
            <FnList label="Functions" picked={custom} pick={toggle} multiple disabled={mode !== "custom"} />
          </fieldset>
        ) : (
          <>
            <fieldset className="dialog-group">
              <legend>Layout</legend>
              <label className="dialog-radio">
                <input type="radio" name="pivot-form" checked={s.outline} onChange={() => set({ outline: true })} />
                <span>Show item labels in outline form</span>
              </label>
              <div className="pivot-settings-indent">
                <Check
                  label="Display labels from the next field in the same column (compact form)"
                  checked={s.compact}
                  disabled={!s.outline}
                  set={(compact) => set({ compact })}
                />
                <Check
                  label="Display subtotals at the top of each group"
                  checked={s.subtotalTop}
                  disabled={!s.outline}
                  set={(subtotalTop) => set({ subtotalTop })}
                />
              </div>
              <label className="dialog-radio">
                <input
                  type="radio"
                  name="pivot-form"
                  checked={!s.outline}
                  onChange={() => set({ outline: false, compact: false })}
                />
                <span>Show item labels in tabular form</span>
              </label>
              <Check
                label="Repeat item labels"
                checked={s.repeatLabels}
                set={(repeatLabels) => set({ repeatLabels })}
              />
              <Check
                label="Insert blank line after each item label"
                checked={s.insertBlankRow}
                set={(insertBlankRow) => set({ insertBlankRow })}
              />
              <Check label="Show items with no data" checked={s.showAll} set={(showAll) => set({ showAll })} />
            </fieldset>
          </>
        )}
      </Tabs>
    </Dialog>
  )
}

type OptionsTab = "layout" | "totals" | "display" | "data"

/**
 * PivotTable Options: the name; Layout & Format (error and empty cell captions, autofit), Totals & Filters (grand
 * totals), Display (+/− buttons, the Values row) and Data (show details, refresh when the file opens).
 */
export function PivotOptionsDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const at = controller.pivots.activeWithCache()
  const [o, setO] = useState<PivotOptions | null>(() => (at ? pivotOptions(at.table, at.cache) : null))
  const [tab, setTab] = useState<OptionsTab>("layout")
  if (!at || !o) return null
  const set = (patch: Partial<PivotOptions>) => setO({ ...o, ...patch })
  const accept = () => {
    const problem = controller.pivots.setOptions(o)
    if (problem) return controller.shell.alert(problem)
    onClose()
  }
  return (
    <Dialog title="PivotTable Options" onCancel={onClose} onAccept={accept} width={440}>
      <div className="goal-seek-fields">
        <label className="dialog-label" htmlFor="pivot-name">
          PivotTable Name:
        </label>
        <input
          id="pivot-name"
          className="dialog-field"
          autoFocus
          value={o.name}
          onChange={(e) => set({ name: e.target.value })}
        />
      </div>
      <Tabs
        tabs={[
          { id: "layout", label: "Layout & Format" },
          { id: "totals", label: "Totals & Filters" },
          { id: "display", label: "Display" },
          { id: "data", label: "Data" },
        ]}
        tab={tab}
        setTab={setTab}
      >
        {tab === "layout" ? (
          <fieldset className="dialog-group">
            <legend>Format</legend>
            <div className="pivot-settings-row">
              <Check label="For error values show:" checked={o.showError} set={(showError) => set({ showError })} />
              <input
                className="dialog-field"
                aria-label="Error value caption"
                disabled={!o.showError}
                value={o.errorCaption}
                onChange={(e) => set({ errorCaption: e.target.value })}
              />
            </div>
            <div className="pivot-settings-row">
              <Check
                label="For empty cells show:"
                checked={o.showMissing}
                set={(showMissing) => set({ showMissing })}
              />
              <input
                className="dialog-field"
                aria-label="Empty cell caption"
                disabled={!o.showMissing}
                value={o.missingCaption}
                onChange={(e) => set({ missingCaption: e.target.value })}
              />
            </div>
            <Check label="Autofit column widths on update" checked={o.autoFit} set={(autoFit) => set({ autoFit })} />
          </fieldset>
        ) : null}
        {tab === "totals" ? (
          <fieldset className="dialog-group">
            <legend>Grand Totals</legend>
            <Check
              label="Show grand totals for rows"
              checked={o.colGrandTotals}
              set={(colGrandTotals) => set({ colGrandTotals })}
            />
            <Check
              label="Show grand totals for columns"
              checked={o.rowGrandTotals}
              set={(rowGrandTotals) => set({ rowGrandTotals })}
            />
          </fieldset>
        ) : null}
        {tab === "display" ? (
          <fieldset className="dialog-group">
            <legend>Display</legend>
            <Check label="Show expand/collapse buttons" checked={o.showDrill} set={(showDrill) => set({ showDrill })} />
            <Check label="Show the Values row" checked={o.valuesRow} set={(valuesRow) => set({ valuesRow })} />
          </fieldset>
        ) : null}
        {tab === "data" ? (
          <fieldset className="dialog-group">
            <legend>PivotTable Data</legend>
            <Check
              label="Refresh data when opening the file"
              checked={o.refreshOnLoad}
              set={(refreshOnLoad) => set({ refreshOnLoad })}
            />
            <Check label="Enable show details" checked={o.enableDrill} set={(enableDrill) => set({ enableDrill })} />
          </fieldset>
        ) : null}
      </Tabs>
    </Dialog>
  )
}
