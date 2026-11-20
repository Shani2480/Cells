import {
  autoDataFieldName,
  defaultBaseField,
  SHOW_AS_BASE_FIELD,
  SHOW_AS_BASE_ITEM,
  SHOW_AS_CHOICES,
  type PivotShowAs,
  type PivotShowAsSetting,
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
/** A list box of plain choices (Base field, Base item). */
function ChoiceList<T>({
  label,
  options,
  value,
  pick,
  disabled,
}: {
  label: string
  options: { value: T; label: string }[]
  value: T | undefined
  pick: (v: T) => void
  disabled?: boolean
}) {
  return (
    <div
      className={`dialog-list pivot-base-list${disabled ? " is-disabled" : ""}`}
      role="listbox"
      aria-label={label}
      aria-disabled={disabled || undefined}
    >
      {options.map((o) => (
        <button
          key={String(o.value)}
          type="button"
          role="option"
          aria-selected={!disabled && o.value === value}
          disabled={disabled}
          className={`dialog-list-item${!disabled && o.value === value ? " is-selected" : ""}`}
          onClick={() => pick(o.value)}
        >
          {o.label}
        </button>
      ))}
    </div>
  )
}

/** The base field and item lists of Show Values As (greyed out when the choice doesn't use them). */
function BaseLists({
  controller,
  show,
  set,
}: {
  controller: SheetController
  show: PivotShowAsSetting
  set: (s: PivotShowAsSetting) => void
}) {
  const at = controller.pivots.activeWithCache()
  if (!at) return null
  const { table, cache } = at
  const usesField = show.as !== null && SHOW_AS_BASE_FIELD.includes(show.as)
  const usesItem = show.as !== null && SHOW_AS_BASE_ITEM.includes(show.as)
  const baseField = show.baseField ?? defaultBaseField(table)
  const items = usesItem ? controller.pivots.baseItems(baseField) : []
  return (
    <div className="pivot-base-lists">
      <div>
        <p className="dialog-label">Base field:</p>
        <ChoiceList
          label="Base field"
          options={cache.fields.map((_, f) => ({ value: f, label: pivotFieldName(table, cache, f) }))}
          value={baseField}
          pick={(f) => set({ ...show, baseField: f, baseItem: show.baseItem === "next" ? "next" : "previous" })}
          disabled={!usesField}
        />
      </div>
      <div>
        <p className="dialog-label">Base item:</p>
        <ChoiceList<number | "previous" | "next">
          label="Base item"
          options={[
            { value: "previous", label: "(previous)" },
            { value: "next", label: "(next)" },
            ...items.map((i) => ({ value: i.x, label: i.label })),
          ]}
          value={show.baseItem ?? "previous"}
          pick={(item) => set({ ...show, baseField, baseItem: item })}
          disabled={!usesItem}
        />
      </div>
    </div>
  )
}

/** A value field's Show Values As as the dialogs edit it (base field and item filled in). */
function showOf(controller: SheetController, index: number): PivotShowAsSetting {
  const at = controller.pivots.activeWithCache()
  const v = at?.table.values[index]
  return {
    as: v?.showAs ?? null,
    baseField: v?.baseField ?? (at ? defaultBaseField(at.table) : 0),
    baseItem: v?.baseItem ?? "previous",
  }
}

/**
 * The cell menu's Show Values As for a choice that needs a base: "Show Values As (Sum of Amount)", with the choice,
 * the base field and, for % Of and the differences, the base item.
 */
export function ShowValuesAsDialog({
  controller,
  index,
  as,
  onClose,
}: {
  controller: SheetController
  index: number
  as: PivotShowAs
  onClose: () => void
}) {
  const at = controller.pivots.activeWithCache()
  const v = at?.table.values[index]
  const [show, setShow] = useState<PivotShowAsSetting>(() => ({ ...showOf(controller, index), as }))
  if (!at || !v) return null
  const label = SHOW_AS_CHOICES.find((c) => c.as === as)?.label ?? ""
  const usesItem = SHOW_AS_BASE_ITEM.includes(as)
  return (
    <Dialog
      title={`Show Values As (${v.name})`}
      onCancel={onClose}
      onAccept={() => {
        const problem = controller.pivots.setShowAs(index, show)
        if (problem) return controller.shell.alert(problem)
        onClose()
      }}
      width={340}
    >
      <p className="dialog-label">Calculation: {label}</p>
      <div className="pivot-base-single">
        <label className="dialog-label" htmlFor="pivot-base-field">
          Base Field:
        </label>
        <select
          id="pivot-base-field"
          className="dialog-field"
          value={show.baseField}
          onChange={(e) => setShow({ ...show, baseField: Number(e.target.value), baseItem: "previous" })}
        >
          {at.cache.fields.map((_, f) => (
            <option key={f} value={f}>
              {pivotFieldName(at.table, at.cache, f)}
            </option>
          ))}
        </select>
        {usesItem ? (
          <>
            <label className="dialog-label" htmlFor="pivot-base-item">
              Base Item:
            </label>
            <select
              id="pivot-base-item"
              className="dialog-field"
              value={String(show.baseItem)}
              onChange={(e) => {
                const t = e.target.value
                setShow({ ...show, baseItem: t === "previous" || t === "next" ? t : Number(t) })
              }}
            >
              <option value="previous">(previous)</option>
              <option value="next">(next)</option>
              {controller.pivots.baseItems(show.baseField ?? 0).map((i) => (
                <option key={i.x} value={i.x}>
                  {i.label}
                </option>
              ))}
            </select>
          </>
        ) : null}
      </div>
    </Dialog>
  )
}

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
  const [tab, setTab] = useState<"summarize" | "show">("summarize")
  const [show, setShow] = useState<PivotShowAsSetting>(() => showOf(controller, index))
  if (!at || !v) return null
  const { table, cache } = at
  const pick = (next: PivotFn) => {
    // An automatic name follows the function.
    if (name === autoDataFieldName(table, cache, v.field, fn)) setName(autoDataFieldName(table, cache, v.field, next))
    setFn(next)
  }
  const accept = () => {
    const problem = controller.pivots.setValueSettings(index, { name, fn, numFmt, show })
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
          tabs={[
            { id: "summarize" as const, label: "Summarize Values By" },
            { id: "show" as const, label: "Show Values As" },
          ]}
          tab={tab}
          setTab={setTab}
        >
          {tab === "summarize" ? (
            <>
              <p className="dialog-label">Summarize value field by</p>
              <p className="dialog-hint">
                Choose the type of calculation that you want to use to summarize data from the selected field
              </p>
              <FnList label="Summarize value field by" picked={[fn]} pick={pick} />
            </>
          ) : (
            <>
              <div className="pivot-show-as-row">
                <label className="dialog-label" htmlFor="pivot-show-as">
                  Show values as
                </label>
                <select
                  id="pivot-show-as"
                  className="dialog-field"
                  value={show.as ?? ""}
                  onChange={(e) => setShow({ ...show, as: (e.target.value || null) as PivotShowAs | null })}
                >
                  {SHOW_AS_CHOICES.map((c) => (
                    <option key={c.as ?? ""} value={c.as ?? ""}>
                      {c.label}
                    </option>
                  ))}
                </select>
              </div>
              <BaseLists controller={controller} show={show} set={setShow} />
            </>
          )}
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
