import {
  parseInput,
  parseRange,
  rangeName,
  type CFEntry,
  type CFRule,
  type CFStyle,
  type CellIsOperator,
  type TimePeriod,
} from "@cells/core"
import { useEffect, useMemo, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { appliesToText, CF_PRESETS, describeRule, nextRuleId, PERIODS, periodText } from "@/sheet/conditional"
import { Dialog } from "../Dialog"

/** What the user typed as a rule value: a formula ("=$A$1"), a number or date, or text (quoted). */
export function valueFormula(text: string): string {
  const t = text.trim()
  if (t.startsWith("=")) return t.slice(1)
  const parsed = parseInput(t)
  if (parsed && typeof parsed.value === "number") return String(parsed.value)
  return `"${t.replace(/"/g, '""')}"`
}

/** The kinds of quick dialog, with Excel's titles and prompts. */
const QUICK = {
  greaterThan: { title: "Greater Than", prompt: "Format cells that are GREATER THAN:" },
  lessThan: { title: "Less Than", prompt: "Format cells that are LESS THAN:" },
  between: { title: "Between", prompt: "Format cells that are BETWEEN:" },
  equal: { title: "Equal To", prompt: "Format cells that are EQUAL TO:" },
  text: { title: "Text That Contains", prompt: "Format cells that contain the text:" },
  date: { title: "A Date Occurring", prompt: "Format cells that contain a date occurring:" },
  duplicate: { title: "Duplicate Values", prompt: "Format cells that contain:" },
  top: { title: "Top 10 Items", prompt: "Format cells that rank in the TOP:" },
  topPercent: { title: "Top 10%", prompt: "Format cells that rank in the TOP:" },
  bottom: { title: "Bottom 10 Items", prompt: "Format cells that rank in the BOTTOM:" },
  bottomPercent: { title: "Bottom 10%", prompt: "Format cells that rank in the BOTTOM:" },
  above: { title: "Above Average", prompt: "Format cells that are ABOVE AVERAGE:" },
  below: { title: "Below Average", prompt: "Format cells that are BELOW AVERAGE:" },
} as const

export type QuickKind = keyof typeof QUICK

export const isQuickKind = (k: string): k is QuickKind => k in QUICK

function PresetSelect({
  value,
  onChange,
  label = "with",
}: {
  value: number
  onChange: (i: number) => void
  label?: string
}) {
  return (
    <label className="cf-with">
      <span>{label}</span>
      <select
        className="dialog-field"
        value={value}
        onChange={(e) => onChange(Number(e.target.value))}
        aria-label="Format with"
      >
        {CF_PRESETS.map((p, i) => (
          <option key={p.label} value={i}>
            {p.label}
          </option>
        ))}
        <option disabled title="Not built yet.">
          Custom Format...
        </option>
      </select>
    </label>
  )
}

/** Highlight Cells Rules and Top/Bottom Rules: one value (or two), a preset format, a live preview. */
export function QuickRuleDialog({
  controller,
  kind,
  onClose,
}: {
  controller: SheetController
  kind: QuickKind
  onClose: () => void
}) {
  const cf = controller.conditional
  const spec = QUICK[kind]
  const suggested = useMemo(
    () =>
      kind === "between"
        ? cf.suggestion("between")
        : kind === "greaterThan" || kind === "lessThan" || kind === "equal"
          ? cf.suggestion(kind)
          : [""],
    [cf, kind],
  )
  const [a, setA] = useState(suggested[0] ?? "")
  const [b, setB] = useState(suggested[1] ?? "")
  const [text, setText] = useState("")
  const [period, setPeriod] = useState<TimePeriod>("yesterday")
  const [unique, setUnique] = useState(false)
  const [rank, setRank] = useState("10")
  const [preset, setPreset] = useState(0)
  const style = CF_PRESETS[preset].style

  const rule = ((): CFRule | null => {
    switch (kind) {
      case "greaterThan":
      case "lessThan":
      case "equal":
        return a.trim() ? { kind: "cellIs", operator: kind as CellIsOperator, values: [valueFormula(a)], style } : null
      case "between":
        return a.trim() && b.trim()
          ? { kind: "cellIs", operator: "between", values: [valueFormula(a), valueFormula(b)], style }
          : null
      case "text":
        return text ? { kind: "text", operator: "contains", text, style } : null
      case "date":
        return { kind: "date", period, style }
      case "duplicate":
        return { kind: "duplicate", unique, style }
      case "top":
      case "topPercent":
      case "bottom":
      case "bottomPercent": {
        const n = Math.round(Number(rank))
        const percent = kind.endsWith("Percent")
        if (!Number.isFinite(n) || n < 1 || n > (percent ? 100 : 1000)) return null
        return { kind: "top", bottom: kind.startsWith("bottom"), percent, rank: n, style }
      }
      case "above":
      case "below":
        return { kind: "average", below: kind === "below", equal: false, style }
    }
  })()

  // Excel shows the rule on the sheet while the dialog is open.
  const key = JSON.stringify(rule)
  useEffect(() => {
    cf.setPreview(rule)
    // `key` stands for the rule (a new object on every render).
  }, [key, cf])
  useEffect(() => () => cf.setPreview(null), [cf])

  const accept = () => {
    if (!rule) return
    cf.add(rule)
    onClose()
  }
  const percent = kind === "topPercent" || kind === "bottomPercent"
  return (
    <Dialog title={spec.title} onCancel={onClose} onAccept={accept} width={kind === "between" ? 560 : 460}>
      <span className="dialog-label">{spec.prompt}</span>
      <div className="cf-quick-row">
        {kind === "greaterThan" || kind === "lessThan" || kind === "equal" ? (
          <input className="dialog-field" aria-label="Value" value={a} onChange={(e) => setA(e.target.value)} />
        ) : null}
        {kind === "between" ? (
          <>
            <input className="dialog-field" aria-label="From" value={a} onChange={(e) => setA(e.target.value)} />
            <span>and</span>
            <input className="dialog-field" aria-label="To" value={b} onChange={(e) => setB(e.target.value)} />
          </>
        ) : null}
        {kind === "text" ? (
          <input className="dialog-field" aria-label="Text" value={text} onChange={(e) => setText(e.target.value)} />
        ) : null}
        {kind === "date" ? (
          <select
            className="dialog-field"
            aria-label="Date"
            value={period}
            onChange={(e) => setPeriod(e.target.value as TimePeriod)}
          >
            {PERIODS.map((p) => (
              <option key={p} value={p}>
                {periodText(p)}
              </option>
            ))}
          </select>
        ) : null}
        {kind === "duplicate" ? (
          <select
            className="dialog-field"
            aria-label="Duplicate or unique"
            value={unique ? "unique" : "duplicate"}
            onChange={(e) => setUnique(e.target.value === "unique")}
          >
            <option value="duplicate">Duplicate</option>
            <option value="unique">Unique</option>
          </select>
        ) : null}
        {kind === "top" || kind === "bottom" || percent ? (
          <>
            <input
              className="dialog-field cf-rank"
              type="number"
              min={1}
              max={percent ? 100 : 1000}
              aria-label="Rank"
              value={rank}
              onChange={(e) => setRank(e.target.value)}
            />
            {percent ? <span>%</span> : null}
          </>
        ) : null}
        <PresetSelect
          value={preset}
          onChange={setPreset}
          label={
            kind === "duplicate"
              ? "values with"
              : kind === "above" || kind === "below"
                ? "for the selected range with"
                : "with"
          }
        />
      </div>
    </Dialog>
  )
}

// ── New Formatting Rule ───────────────────────────────────────────────────────────────────────────────────────

const RULE_TYPES = [
  { id: "values", label: "Format all cells based on their values", built: false },
  { id: "contains", label: "Format only cells that contain", built: true },
  { id: "top", label: "Format only top or bottom ranked values", built: true },
  { id: "average", label: "Format only values that are above or below average", built: true },
  { id: "unique", label: "Format only unique or duplicate values", built: true },
  { id: "formula", label: "Use a formula to determine which cells to format", built: true },
] as const

type RuleType = (typeof RULE_TYPES)[number]["id"]

const OPERATORS: { op: CellIsOperator; label: string }[] = [
  { op: "between", label: "between" },
  { op: "notBetween", label: "not between" },
  { op: "equal", label: "equal to" },
  { op: "notEqual", label: "not equal to" },
  { op: "greaterThan", label: "greater than" },
  { op: "lessThan", label: "less than" },
  { op: "greaterThanOrEqual", label: "greater than or equal to" },
  { op: "lessThanOrEqual", label: "less than or equal to" },
]

/**
 * New Formatting Rule. The format comes from the presets of the quick dialogs: Excel's Format... button (a Format
 * Cells dialog for rules) is not built yet.
 */
export function NewRuleDialog({
  controller,
  start,
  onRule,
  onClose,
  title = "New Formatting Rule",
}: {
  controller: SheetController
  start: string
  /** Where the rule goes (Manage Rules keeps it in its own list); default: added to the sheet. */
  onRule?: (rule: CFRule) => void
  onClose: () => void
  title?: string
}) {
  const [type, setType] = useState<RuleType>(RULE_TYPES.some((t) => t.id === start) ? (start as RuleType) : "contains")
  const [subject, setSubject] = useState<"value" | "text" | "date" | "blanks" | "noBlanks" | "errors" | "noErrors">(
    "value",
  )
  const [operator, setOperator] = useState<CellIsOperator>("between")
  const [a, setA] = useState("")
  const [b, setB] = useState("")
  const [textOp, setTextOp] = useState<"contains" | "notContains" | "beginsWith" | "endsWith">("contains")
  const [period, setPeriod] = useState<TimePeriod>("yesterday")
  const [bottom, setBottom] = useState(false)
  const [rank, setRank] = useState("10")
  const [percent, setPercent] = useState(false)
  const [below, setBelow] = useState(false)
  const [equal, setEqual] = useState(false)
  const [unique, setUnique] = useState(false)
  const [formula, setFormula] = useState("")
  const [preset, setPreset] = useState(0)
  const [error, setError] = useState<string | null>(null)
  const style: CFStyle = CF_PRESETS[preset].style

  const build = (): CFRule | string => {
    switch (type) {
      case "values":
        return "Not built yet."
      case "contains":
        switch (subject) {
          case "value": {
            const two = operator === "between" || operator === "notBetween"
            if (!a.trim() || (two && !b.trim())) return "Enter a value."
            return {
              kind: "cellIs",
              operator,
              values: two ? [valueFormula(a), valueFormula(b)] : [valueFormula(a)],
              style,
            }
          }
          case "text":
            return a ? { kind: "text", operator: textOp, text: a, style } : "Enter a value."
          case "date":
            return { kind: "date", period, style }
          case "blanks":
          case "noBlanks":
            return { kind: "blanks", blank: subject === "blanks", style }
          case "errors":
          case "noErrors":
            return { kind: "errors", errors: subject === "errors", style }
        }
        break
      case "top": {
        const n = Math.round(Number(rank))
        if (!Number.isFinite(n) || n < 1 || n > (percent ? 100 : 1000))
          return percent ? "Enter a whole number between 1 and 100." : "Enter a whole number between 1 and 1000."
        return { kind: "top", bottom, percent, rank: n, style }
      }
      case "average":
        return { kind: "average", below, equal, style }
      case "unique":
        return { kind: "duplicate", unique, style }
      case "formula": {
        const f = formula.trim().replace(/^=/, "")
        return f ? { kind: "expression", formula: f, style } : "Enter a formula."
      }
    }
    return "Not built yet."
  }

  const accept = () => {
    const rule = build()
    if (typeof rule === "string") {
      setError(rule)
      return
    }
    if (onRule) onRule(rule)
    else controller.conditional.add(rule)
    onClose()
  }

  return (
    <Dialog title={title} onCancel={onClose} onAccept={accept} width={520}>
      <span className="dialog-label">Select a Rule Type:</span>
      <div className="dialog-list cf-rule-types" role="listbox" aria-label="Rule type">
        {RULE_TYPES.map((t) => (
          <button
            key={t.id}
            type="button"
            role="option"
            aria-selected={t.id === type}
            disabled={!t.built}
            title={t.built ? undefined : "Not built yet."}
            className={`dialog-list-item${t.id === type ? " is-selected" : ""}`}
            onClick={() => {
              setType(t.id)
              setError(null)
            }}
          >
            ► {t.label}
          </button>
        ))}
      </div>
      <span className="dialog-label">Edit the Rule Description:</span>
      <div className="cf-rule-edit">
        {type === "contains" ? (
          <>
            <span className="dialog-label">Format only cells with:</span>
            <div className="cf-quick-row">
              <select
                className="dialog-field"
                aria-label="Subject"
                value={subject}
                onChange={(e) => setSubject(e.target.value as typeof subject)}
              >
                <option value="value">Cell Value</option>
                <option value="text">Specific Text</option>
                <option value="date">Dates Occurring</option>
                <option value="blanks">Blanks</option>
                <option value="noBlanks">No Blanks</option>
                <option value="errors">Errors</option>
                <option value="noErrors">No Errors</option>
              </select>
              {subject === "value" ? (
                <>
                  <select
                    className="dialog-field"
                    aria-label="Operator"
                    value={operator}
                    onChange={(e) => setOperator(e.target.value as CellIsOperator)}
                  >
                    {OPERATORS.map((o) => (
                      <option key={o.op} value={o.op}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                  <input className="dialog-field" aria-label="Value" value={a} onChange={(e) => setA(e.target.value)} />
                  {operator === "between" || operator === "notBetween" ? (
                    <>
                      <span>and</span>
                      <input
                        className="dialog-field"
                        aria-label="Second value"
                        value={b}
                        onChange={(e) => setB(e.target.value)}
                      />
                    </>
                  ) : null}
                </>
              ) : null}
              {subject === "text" ? (
                <>
                  <select
                    className="dialog-field"
                    aria-label="Text operator"
                    value={textOp}
                    onChange={(e) => setTextOp(e.target.value as typeof textOp)}
                  >
                    <option value="contains">containing</option>
                    <option value="notContains">not containing</option>
                    <option value="beginsWith">beginning with</option>
                    <option value="endsWith">ending with</option>
                  </select>
                  <input className="dialog-field" aria-label="Value" value={a} onChange={(e) => setA(e.target.value)} />
                </>
              ) : null}
              {subject === "date" ? (
                <select
                  className="dialog-field"
                  aria-label="Date"
                  value={period}
                  onChange={(e) => setPeriod(e.target.value as TimePeriod)}
                >
                  {PERIODS.map((p) => (
                    <option key={p} value={p}>
                      {periodText(p)}
                    </option>
                  ))}
                </select>
              ) : null}
            </div>
          </>
        ) : null}
        {type === "top" ? (
          <>
            <span className="dialog-label">Format values that rank in the:</span>
            <div className="cf-quick-row">
              <select
                className="dialog-field"
                aria-label="Top or bottom"
                value={bottom ? "bottom" : "top"}
                onChange={(e) => setBottom(e.target.value === "bottom")}
              >
                <option value="top">Top</option>
                <option value="bottom">Bottom</option>
              </select>
              <input
                className="dialog-field cf-rank"
                aria-label="Rank"
                value={rank}
                onChange={(e) => setRank(e.target.value)}
              />
              <label className="fc-check">
                <input type="checkbox" checked={percent} onChange={(e) => setPercent(e.target.checked)} />
                <span>% of the selected range</span>
              </label>
            </div>
          </>
        ) : null}
        {type === "average" ? (
          <>
            <span className="dialog-label">Format values that are:</span>
            <select
              className="dialog-field"
              aria-label="Above or below"
              value={`${below ? "below" : "above"}${equal ? "Equal" : ""}`}
              onChange={(e) => {
                setBelow(e.target.value.startsWith("below"))
                setEqual(e.target.value.endsWith("Equal"))
              }}
            >
              <option value="above">above</option>
              <option value="below">below</option>
              <option value="aboveEqual">equal or above</option>
              <option value="belowEqual">equal or below</option>
            </select>
            <span className="dialog-label">the average for the selected range</span>
          </>
        ) : null}
        {type === "unique" ? (
          <>
            <span className="dialog-label">Format all:</span>
            <select
              className="dialog-field"
              aria-label="Duplicate or unique"
              value={unique ? "unique" : "duplicate"}
              onChange={(e) => setUnique(e.target.value === "unique")}
            >
              <option value="duplicate">duplicate</option>
              <option value="unique">unique</option>
            </select>
            <span className="dialog-label">values in the selected range</span>
          </>
        ) : null}
        {type === "formula" ? (
          <>
            <span className="dialog-label">Format values where this formula is true:</span>
            <input
              className="dialog-field"
              aria-label="Formula"
              value={formula}
              onChange={(e) => setFormula(e.target.value)}
              placeholder="="
            />
          </>
        ) : null}
        <div className="cf-quick-row">
          <span className="dialog-label">Preview:</span>
          <span className="cf-format-preview" style={presetPreview(style)}>
            AaBbCcYyZz
          </span>
          <PresetSelect value={preset} onChange={setPreset} label="Format:" />
        </div>
      </div>
      {error ? (
        <span className="dialog-label" role="alert">
          {error}
        </span>
      ) : null}
    </Dialog>
  )
}

export function presetPreview(style: CFStyle): React.CSSProperties {
  return {
    background: style.fill,
    color: style.fontColor,
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? "italic" : undefined,
    border: style.borderTop ? `1px solid ${style.borderTop.color ?? "#000"}` : undefined,
  }
}

// ── Conditional Formatting Rules Manager ───────────────────────────────────────────────────────────────────────

/** Parses "Applies to" ("=$A$1:$B$5,$D$1"); null when it isn't a reference. */
export function parseAppliesTo(text: string) {
  const parts = text.replace(/^=/, "").replace(/\$/g, "").split(",")
  const ranges = parts.map((p) => parseRange(p.trim()))
  return ranges.every((r) => r !== null) && ranges.length ? ranges : null
}

export function ManageRulesDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [entries, setEntries] = useState<CFEntry[]>(() => JSON.parse(JSON.stringify(controller.sheet.conditional)))
  const [scope, setScope] = useState<"selection" | "sheet">("selection")
  const [chosen, setChosen] = useState<number | null>(null)
  const [adding, setAdding] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const selection = controller.selection.ranges
  const overlaps = (e: CFEntry) =>
    e.ranges.some((r) => selection.some((s) => r.r0 <= s.r1 && r.r1 >= s.r0 && r.c0 <= s.c1 && r.c1 >= s.c0))
  const shown = entries.filter((e) => scope === "sheet" || overlaps(e))
  const index = entries.findIndex((e) => e.id === chosen)

  const move = (dir: -1 | 1) => {
    const shownIndex = shown.findIndex((e) => e.id === chosen)
    const other = shown[shownIndex + dir]
    if (!other) return
    const i = entries.findIndex((e) => e.id === other.id)
    const next = [...entries]
    ;[next[index], next[i]] = [next[i], next[index]]
    setEntries(next)
  }

  const apply = () => {
    controller.conditional.replace(entries)
  }

  return (
    <>
      <Dialog
        title="Conditional Formatting Rules Manager"
        onCancel={onClose}
        onAccept={() => {
          if (error) return
          apply()
          onClose()
        }}
        width={720}
      >
        <label className="cf-quick-row">
          <span className="dialog-label">Show formatting rules for:</span>
          <select
            className="dialog-field"
            aria-label="Show formatting rules for"
            value={scope}
            onChange={(e) => setScope(e.target.value as typeof scope)}
          >
            <option value="selection">Current Selection</option>
            <option value="sheet">This Worksheet</option>
          </select>
        </label>
        <div className="cf-manage-toolbar">
          <button type="button" className="dialog-btn" onClick={() => setAdding(true)}>
            New Rule...
          </button>
          <button type="button" className="dialog-btn" disabled title="Not built yet.">
            Edit Rule...
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={index < 0}
            onClick={() => {
              setEntries(entries.filter((e) => e.id !== chosen))
              setChosen(null)
            }}
          >
            Delete Rule
          </button>
          <button
            type="button"
            className="dialog-btn"
            disabled={index < 0}
            onClick={() => {
              const copy = {
                ...JSON.parse(JSON.stringify(entries[index])),
                id: Math.max(nextRuleId(controller), ...entries.map((e) => e.id + 1)),
              }
              const next = [...entries]
              next.splice(index + 1, 0, copy)
              setEntries(next)
            }}
          >
            Duplicate Rule
          </button>
          <button
            type="button"
            className="dialog-btn"
            aria-label="Move Up"
            title="Move Up"
            disabled={index < 0}
            onClick={() => move(-1)}
          >
            ▲
          </button>
          <button
            type="button"
            className="dialog-btn"
            aria-label="Move Down"
            title="Move Down"
            disabled={index < 0}
            onClick={() => move(1)}
          >
            ▼
          </button>
        </div>
        <div className="cf-manage-table" role="grid" aria-label="Rules">
          <div className="cf-manage-head" role="row">
            <span role="columnheader">Rule (applied in order shown)</span>
            <span role="columnheader">Format</span>
            <span role="columnheader">Applies to</span>
            <span role="columnheader">Stop If True</span>
          </div>
          {shown.map((e) => (
            <div
              key={e.id}
              role="row"
              aria-selected={e.id === chosen}
              className={`cf-manage-row${e.id === chosen ? " is-selected" : ""}`}
              onClick={() => setChosen(e.id)}
            >
              <span role="gridcell">{describeRule(e.rule)}</span>
              <span role="gridcell">
                {"style" in e.rule ? (
                  <span className="cf-format-preview" style={presetPreview(e.rule.style)}>
                    AaBbCcYyZz
                  </span>
                ) : null}
              </span>
              <span role="gridcell">
                <input
                  className="dialog-field"
                  aria-label="Applies to"
                  defaultValue={appliesToText(e.ranges)}
                  onBlur={(ev) => {
                    const ranges = parseAppliesTo(ev.target.value)
                    if (!ranges) {
                      setError("The reference is not valid.")
                      return
                    }
                    setError(null)
                    setEntries(entries.map((x) => (x.id === e.id ? { ...x, ranges } : x)))
                  }}
                />
              </span>
              <span role="gridcell">
                <input
                  type="checkbox"
                  aria-label="Stop If True"
                  checked={e.stopIfTrue}
                  disabled={!("style" in e.rule)}
                  onChange={(ev) =>
                    setEntries(entries.map((x) => (x.id === e.id ? { ...x, stopIfTrue: ev.target.checked } : x)))
                  }
                />
              </span>
            </div>
          ))}
        </div>
        {error ? (
          <span className="dialog-label" role="alert">
            {error}
          </span>
        ) : null}
        <div className="cf-manage-apply">
          <button type="button" className="dialog-btn" onClick={apply}>
            Apply
          </button>
        </div>
      </Dialog>
      {adding ? (
        <NewRuleDialog
          controller={controller}
          start="contains"
          onClose={() => setAdding(false)}
          onRule={(rule) =>
            setEntries([
              {
                id: Math.max(nextRuleId(controller), ...entries.map((e) => e.id + 1)),
                ranges: selection.map((r) => ({ ...r })),
                rule,
                stopIfTrue: false,
              },
              ...entries,
            ])
          }
        />
      ) : null}
    </>
  )
}

export const selectionText = (c: SheetController) => c.selection.ranges.map((r) => rangeName(r)).join(",")
