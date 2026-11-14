import {
  DEFAULT_VALIDATION,
  sameValidation,
  twoValues,
  usesOperator,
  type ErrorStyle,
  type Validation,
  type ValidationOperator,
  type ValidationType,
} from "@cells/core"
import { CircleX, Info, TriangleAlert } from "lucide-react"
import { useState } from "react"
import type { SheetController } from "@/sheet/controller"
import {
  blankError,
  fieldError,
  fieldFormula,
  fieldLabels,
  fieldText,
  LIST_SOURCE_ERROR,
  ORDER_ERROR,
  type AlertAnswer,
  type ValidationAlert,
  type ValidationDialogInit,
} from "@/sheet/validation"
import { Dialog } from "./Dialog"

const TYPES: { id: ValidationType; label: string }[] = [
  { id: "any", label: "Any value" },
  { id: "whole", label: "Whole number" },
  { id: "decimal", label: "Decimal" },
  { id: "list", label: "List" },
  { id: "date", label: "Date" },
  { id: "time", label: "Time" },
  { id: "textLength", label: "Text length" },
  { id: "custom", label: "Custom" },
]

const OPERATORS: { id: ValidationOperator; label: string }[] = [
  { id: "between", label: "between" },
  { id: "notBetween", label: "not between" },
  { id: "equal", label: "equal to" },
  { id: "notEqual", label: "not equal to" },
  { id: "greaterThan", label: "greater than" },
  { id: "lessThan", label: "less than" },
  { id: "greaterThanOrEqual", label: "greater than or equal to" },
  { id: "lessThanOrEqual", label: "less than or equal to" },
]

const STYLES: { id: ErrorStyle; label: string }[] = [
  { id: "stop", label: "Stop" },
  { id: "warning", label: "Warning" },
  { id: "information", label: "Information" },
]

type Tab = "settings" | "input" | "error"
const TABS: { id: Tab; label: string }[] = [
  { id: "settings", label: "Settings" },
  { id: "input", label: "Input Message" },
  { id: "error", label: "Error Alert" },
]

/** Excel's alert pictures: a red cross (Stop), a yellow triangle (Warning), a blue i (Information). */
export function AlertIcon({ style, size = 32 }: { style: ErrorStyle; size?: number }) {
  const Icon = style === "stop" ? CircleX : style === "warning" ? TriangleAlert : Info
  return (
    <span className={`dv-alert-icon is-${style}`} aria-hidden>
      <Icon size={size} strokeWidth={1.5} />
    </span>
  )
}

/**
 * Data Validation: Settings (what may go in), Input Message (shown while the cell is selected) and Error Alert
 * (shown when an entry breaks the rule), with Clear All, as Excel has them.
 */
export function DataValidationDialog({
  controller,
  init,
  onClose,
}: {
  controller: SheetController
  init: ValidationDialogInit
  onClose: () => void
}) {
  const opened = init.rule
  const hadRule = !sameValidation(opened, DEFAULT_VALIDATION)
  const [tab, setTab] = useState<Tab>("settings")
  const [rule, setRule] = useState<Validation>(opened)
  const [value1, setValue1] = useState(() => fieldText(opened.formula1, opened.type))
  const [value2, setValue2] = useState(() => fieldText(opened.formula2, opened.type))
  const [sameToo, setSameToo] = useState(false)
  const set = (patch: Partial<Validation>) => setRule((r) => ({ ...r, ...patch }))

  const operatorOn = usesOperator(rule.type)
  const op = operatorOn ? rule.operator : "between"
  const [label1, label2] = fieldLabels(rule.type, op)
  const two = operatorOn && twoValues(op)

  const setType = (type: ValidationType) => {
    // Another type starts with empty values, as in Excel.
    if (type !== rule.type) {
      setValue1("")
      setValue2("")
    }
    set({ type })
  }

  const clearAll = () => {
    setRule({ ...DEFAULT_VALIDATION })
    setValue1("")
    setValue2("")
  }

  const accept = () => {
    const out: Validation = { ...rule, operator: op, formula1: "", formula2: "" }
    if (rule.type !== "any") {
      const f1 = fieldFormula(value1, rule.type)
      if (rule.type === "list" && f1 === null) return controller.shell.alert(LIST_SOURCE_ERROR)
      if (f1 === null) return controller.shell.alert(value1.trim() ? fieldError(label1) : blankError(label1))
      out.formula1 = f1
      if (two && label2) {
        const f2 = fieldFormula(value2, rule.type)
        if (f2 === null) return controller.shell.alert(value2.trim() ? fieldError(label2) : blankError(label2))
        out.formula2 = f2
        const [a, b] = [Number(f1), Number(f2)]
        if (/^[-\d.]/.test(f1) && /^[-\d.]/.test(f2) && !Number.isNaN(a) && !Number.isNaN(b) && b < a)
          return controller.shell.alert(ORDER_ERROR)
      }
    }
    if (out.type !== "list") out.dropDown = true
    onClose()
    controller.validation.apply(init.ranges, out, sameToo && hadRule ? opened : null)
  }

  const check = (label: string, checked: boolean, onChange: (v: boolean) => void, disabled = false) => (
    <label className={disabled ? "fc-check is-disabled" : "fc-check"}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} />
      {label}
    </label>
  )

  return (
    <Dialog
      title="Data Validation"
      width={440}
      onCancel={onClose}
      onAccept={accept}
      actions={
        <>
          <button type="button" className="dialog-btn dv-clear" onClick={clearAll}>
            Clear All
          </button>
          <button type="button" className="dialog-btn is-primary" onClick={accept}>
            OK
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            Cancel
          </button>
        </>
      }
    >
      <div className="fc-tabs" role="tablist">
        {TABS.map((t) => (
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
      <div className="fc-page dv-page" role="tabpanel" aria-label={TABS.find((t) => t.id === tab)!.label}>
        {tab === "settings" ? (
          <div className="fc-column">
            <div className="dv-heading">Validation criteria</div>
            <div className="dv-grid">
              <label className="dialog-label" htmlFor="dv-allow">
                Allow:
              </label>
              <span />
              <select
                id="dv-allow"
                className="dialog-field"
                value={rule.type}
                onChange={(e) => setType(e.target.value as ValidationType)}
              >
                {TYPES.map((t) => (
                  <option key={t.id} value={t.id}>
                    {t.label}
                  </option>
                ))}
              </select>
              {check("Ignore blank", rule.allowBlank, (allowBlank) => set({ allowBlank }), rule.type === "any")}
              <label className="dialog-label" htmlFor="dv-data">
                Data:
              </label>
              <span />
              <select
                id="dv-data"
                className="dialog-field"
                value={op}
                disabled={!operatorOn}
                onChange={(e) => set({ operator: e.target.value as ValidationOperator })}
              >
                {OPERATORS.map((o) => (
                  <option key={o.id} value={o.id}>
                    {o.label}
                  </option>
                ))}
              </select>
              {rule.type === "list" ? (
                check("In-cell dropdown", rule.dropDown, (dropDown) => set({ dropDown }))
              ) : (
                <span />
              )}
            </div>
            {rule.type !== "any" ? (
              <div className="fc-column dv-values">
                <label className="dialog-label" htmlFor="dv-value1">
                  {label1}:
                </label>
                <input
                  id="dv-value1"
                  className="dialog-field"
                  value={value1}
                  spellCheck={false}
                  onChange={(e) => setValue1(e.target.value)}
                />
                {two && label2 ? (
                  <>
                    <label className="dialog-label" htmlFor="dv-value2">
                      {label2}:
                    </label>
                    <input
                      id="dv-value2"
                      className="dialog-field"
                      value={value2}
                      spellCheck={false}
                      onChange={(e) => setValue2(e.target.value)}
                    />
                  </>
                ) : null}
              </div>
            ) : null}
            <div className="dv-bottom">
              {check("Apply these changes to all other cells with the same settings", sameToo, setSameToo, !hadRule)}
            </div>
          </div>
        ) : null}
        {tab === "input" ? (
          <div className="fc-column">
            {check("Show input message when cell is selected", rule.showInput, (showInput) => set({ showInput }))}
            <div className="dv-heading">When cell is selected, show this input message:</div>
            <label className="dialog-label" htmlFor="dv-prompt-title">
              Title:
            </label>
            <input
              id="dv-prompt-title"
              className="dialog-field"
              maxLength={32}
              value={rule.promptTitle}
              onChange={(e) => set({ promptTitle: e.target.value })}
            />
            <label className="dialog-label" htmlFor="dv-prompt">
              Input message:
            </label>
            <textarea
              id="dv-prompt"
              className="dialog-field dv-text"
              maxLength={255}
              value={rule.prompt}
              onChange={(e) => set({ prompt: e.target.value })}
              onKeyDown={(e) => e.key === "Enter" && e.stopPropagation()}
            />
          </div>
        ) : null}
        {tab === "error" ? (
          <div className="fc-column">
            {check("Show error alert after invalid data is entered", rule.showError, (showError) => set({ showError }))}
            <div className="dv-heading">When user enters invalid data, show this error alert:</div>
            <div className="dv-error">
              <div className="fc-column">
                <label className="dialog-label" htmlFor="dv-style">
                  Style:
                </label>
                <select
                  id="dv-style"
                  className="dialog-field"
                  value={rule.errorStyle}
                  onChange={(e) => set({ errorStyle: e.target.value as ErrorStyle })}
                >
                  {STYLES.map((s) => (
                    <option key={s.id} value={s.id}>
                      {s.label}
                    </option>
                  ))}
                </select>
                <AlertIcon style={rule.errorStyle} size={44} />
              </div>
              <div className="fc-column fc-grow">
                <label className="dialog-label" htmlFor="dv-error-title">
                  Title:
                </label>
                <input
                  id="dv-error-title"
                  className="dialog-field"
                  maxLength={32}
                  value={rule.errorTitle}
                  onChange={(e) => set({ errorTitle: e.target.value })}
                />
                <label className="dialog-label" htmlFor="dv-error">
                  Error message:
                </label>
                <textarea
                  id="dv-error"
                  className="dialog-field dv-text"
                  maxLength={225}
                  value={rule.error}
                  onChange={(e) => set({ error: e.target.value })}
                  onKeyDown={(e) => e.key === "Enter" && e.stopPropagation()}
                />
              </div>
            </div>
          </div>
        ) : null}
      </div>
    </Dialog>
  )
}

/** The buttons of each alert, as Excel's: Stop asks Retry or Cancel, Warning "Continue?", Information OK. */
const BUTTONS: Record<ErrorStyle, { answer: AlertAnswer; label: string }[]> = {
  stop: [
    { answer: "retry", label: "Retry" },
    { answer: "cancel", label: "Cancel" },
  ],
  warning: [
    { answer: "yes", label: "Yes" },
    { answer: "no", label: "No" },
    { answer: "cancel", label: "Cancel" },
  ],
  information: [
    { answer: "ok", label: "OK" },
    { answer: "cancel", label: "Cancel" },
  ],
}

/** The error alert after an entry breaks its cell's rule. */
export function ValidationAlertDialog({
  alert,
  onAnswer,
}: {
  alert: ValidationAlert
  onAnswer: (answer: AlertAnswer) => void
}) {
  const buttons = BUTTONS[alert.style]
  return (
    <Dialog
      title={alert.title}
      width={400}
      onCancel={() => onAnswer("cancel")}
      onAccept={() => onAnswer(buttons[0].answer)}
      actions={buttons.map((b, i) => (
        <button
          key={b.answer}
          type="button"
          className={i === 0 ? "dialog-btn is-primary" : "dialog-btn"}
          data-autofocus={i === 0 ? true : undefined}
          onClick={() => onAnswer(b.answer)}
        >
          {b.label}
        </button>
      ))}
    >
      <div className="dv-alert" role="alert">
        <AlertIcon style={alert.style} />
        <p className="dialog-message">
          {alert.message}
          {alert.style === "warning" ? (
            <>
              <br />
              <br />
              Continue?
            </>
          ) : null}
        </p>
      </div>
    </Dialog>
  )
}
