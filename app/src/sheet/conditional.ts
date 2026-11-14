import {
  clearRulesIn,
  defaultIconThresholds,
  ICON_SETS,
  rangeName,
  type CFEntry,
  type CFRule,
  type CFStyle,
  type CellIsOperator,
  type IconSetName,
  type Range,
  type TimePeriod,
} from "@cells/core"
import type { SheetController } from "./controller"

/** The "with" list of Excel's quick rule dialogs (colours read from Excel's own presets). */
export const CF_PRESETS: { label: string; style: CFStyle }[] = [
  { label: "Light Red Fill with Dark Red Text", style: { fill: "#FFC7CE", fontColor: "#9C0006" } },
  { label: "Yellow Fill with Dark Yellow Text", style: { fill: "#FFEB9C", fontColor: "#9C5700" } },
  { label: "Green Fill with Dark Green Text", style: { fill: "#C6EFCE", fontColor: "#006100" } },
  { label: "Light Red Fill", style: { fill: "#FFC7CE" } },
  { label: "Red Text", style: { fontColor: "#9C0006" } },
  {
    label: "Red Border",
    style: {
      borderTop: { style: "thin", color: "#9C0006" },
      borderBottom: { style: "thin", color: "#9C0006" },
      borderLeft: { style: "thin", color: "#9C0006" },
      borderRight: { style: "thin", color: "#9C0006" },
    },
  },
]

/** Data Bars gallery: Excel's six colours, gradient and solid. */
export const DATA_BAR_COLORS = [
  { name: "Blue", color: "#638EC6" },
  { name: "Green", color: "#63C384" },
  { name: "Red", color: "#FF555A" },
  { name: "Orange", color: "#FFB628" },
  { name: "Light Blue", color: "#008AEF" },
  { name: "Purple", color: "#D6007B" },
]

const G = "#63BE7B"
const Y = "#FFEB84"
const R = "#F8696B"
const W = "#FCFCFF"
const B = "#5A8AC6"

/** Color Scales gallery, in Excel's order (first colour for the lowest value). */
export const COLOR_SCALES: { name: string; colors: string[] }[] = [
  { name: "Green - Yellow - Red Color Scale", colors: [R, Y, G].reverse() },
  { name: "Red - Yellow - Green Color Scale", colors: [R, Y, G] },
  { name: "Green - White - Red Color Scale", colors: [R, W, G].reverse() },
  { name: "Red - White - Green Color Scale", colors: [R, W, G] },
  { name: "Blue - White - Red Color Scale", colors: [R, W, B].reverse() },
  { name: "Red - White - Blue Color Scale", colors: [R, W, B] },
  { name: "White - Red Color Scale", colors: [R, W] },
  { name: "Red - White Color Scale", colors: [W, R] },
  { name: "Green - White Color Scale", colors: [W, G] },
  { name: "White - Green Color Scale", colors: [G, W] },
  { name: "Green - Yellow Color Scale", colors: ["#FFEF9C", G] },
  { name: "Yellow - Green Color Scale", colors: [G, "#FFEF9C"] },
]

export const ICON_SET_GROUPS: { heading: string; sets: { set: IconSetName; name: string }[] }[] = [
  {
    heading: "Directional",
    sets: [
      { set: "3Arrows", name: "3 Arrows (Colored)" },
      { set: "3ArrowsGray", name: "3 Arrows (Gray)" },
      { set: "4Arrows", name: "4 Arrows (Colored)" },
      { set: "5Arrows", name: "5 Arrows (Colored)" },
    ],
  },
  {
    heading: "Shapes",
    sets: [
      { set: "3TrafficLights1", name: "3 Traffic Lights (Unrimmed)" },
      { set: "3Signs", name: "3 Signs" },
    ],
  },
  {
    heading: "Indicators",
    sets: [
      { set: "3Symbols", name: "3 Symbols (Circled)" },
      { set: "3Flags", name: "3 Flags" },
    ],
  },
  {
    heading: "Ratings",
    sets: [
      { set: "3Stars", name: "3 Stars" },
      { set: "5Quarters", name: "5 Quarters" },
    ],
  },
]

const OPERATOR_TEXT: Record<CellIsOperator, string> = {
  greaterThan: ">",
  lessThan: "<",
  greaterThanOrEqual: ">=",
  lessThanOrEqual: "<=",
  equal: "=",
  notEqual: "<>",
  between: "between",
  notBetween: "not between",
}

const PERIOD_TEXT: Record<TimePeriod, string> = {
  yesterday: "Yesterday",
  today: "Today",
  tomorrow: "Tomorrow",
  last7Days: "In the last 7 days",
  lastWeek: "Last week",
  thisWeek: "This week",
  nextWeek: "Next week",
  lastMonth: "Last month",
  thisMonth: "This month",
  nextMonth: "Next month",
}

export const PERIODS = Object.keys(PERIOD_TEXT) as TimePeriod[]
export const periodText = (p: TimePeriod) => PERIOD_TEXT[p]

/** A value as Excel shows it in rule descriptions: formulas with "=", text in quotes. */
const shownValue = (v: string) => (/^-?\d+(\.\d+)?$/.test(v) || v.startsWith('"') ? v : "=" + v)

/** Manage Rules' "Rule (applied in order shown)" column. */
export function describeRule(rule: CFRule): string {
  switch (rule.kind) {
    case "cellIs": {
      const [a, b] = rule.values.map(shownValue)
      if (rule.operator === "between" || rule.operator === "notBetween")
        return `Cell Value ${OPERATOR_TEXT[rule.operator]} ${a} and ${b}`
      return `Cell Value ${OPERATOR_TEXT[rule.operator]} ${a}`
    }
    case "text": {
      const verb = {
        contains: "contains",
        notContains: "does not contain",
        beginsWith: "begins with",
        endsWith: "ends with",
      }
      return `Cell Value ${verb[rule.operator]} '${rule.text}'`
    }
    case "date":
      return PERIOD_TEXT[rule.period]
    case "duplicate":
      return rule.unique ? "Unique Values" : "Duplicate Values"
    case "top":
      return `${rule.bottom ? "Bottom" : "Top"} ${rule.rank}${rule.percent ? "%" : ""}`
    case "average":
      return `${rule.below ? "Below" : "Above"} Average`
    case "blanks":
      return rule.blank ? "Blanks" : "No Blanks"
    case "errors":
      return rule.errors ? "Errors" : "No Errors"
    case "expression":
      return `Formula: =${rule.formula}`
    case "colorScale":
      return "Graded Color Scale"
    case "dataBar":
      return "Data Bar"
    case "iconSet":
      return "Icon Set"
  }
}

/** "Applies to": the ranges as absolute references (=$A$1:$B$5). */
export function appliesToText(ranges: readonly Range[]): string {
  return "=" + ranges.map((r) => absolute(rangeName(r))).join(",")
}

const absolute = (name: string) => name.replace(/([A-Z]+)(\d+)/g, "$$$1$$$2").replace(/(^|:)([A-Z]+)(?=:|$)/g, "$1$$$2")

/** The next free rule id of the workbook (ids only need to be unique within a sheet's list). */
export function nextRuleId(c: SheetController): number {
  let max = 0
  for (const sheet of c.wb.sheets) for (const e of sheet.conditional) max = Math.max(max, e.id)
  return max + 1
}

/** Conditional formatting commands on the selection. */
export class ConditionalCommands {
  /** A rule shown while a dialog is open (Excel previews as you choose); not part of the workbook. */
  preview: CFEntry | null = null

  constructor(private readonly c: SheetController) {}

  private commit(entries: CFEntry[], label: string) {
    if (this.c.edit) this.c.commitEdit()
    const sel = this.c.selection
    this.c.transact([{ kind: "conditional", sheet: this.c.sheetIndex, entries }], { label, before: sel, after: sel })
  }

  /** A new rule over the selection, at the top of the list (Excel gives new rules the highest priority). */
  add(rule: CFRule, label = "Conditional Formatting") {
    this.preview = null
    const entry: CFEntry = {
      id: nextRuleId(this.c),
      ranges: this.c.selection.ranges.map((r) => ({ ...r })),
      rule,
      stopIfTrue: false,
    }
    this.commit([entry, ...this.c.sheet.conditional], label)
  }

  dataBar(color: string, gradient: boolean) {
    this.add({ kind: "dataBar", min: { type: "automatic" }, max: { type: "automatic" }, color, gradient })
  }

  colorScale(colors: string[]) {
    const points =
      colors.length === 3
        ? [
            { at: { type: "min" as const }, color: colors[0] },
            { at: { type: "percentile" as const, value: "50" }, color: colors[1] },
            { at: { type: "max" as const }, color: colors[2] },
          ]
        : [
            { at: { type: "min" as const }, color: colors[0] },
            { at: { type: "max" as const }, color: colors[1] },
          ]
    this.add({ kind: "colorScale", points })
  }

  iconSet(set: IconSetName) {
    this.add({
      kind: "iconSet",
      set,
      reverse: false,
      thresholds: defaultIconThresholds(ICON_SETS[set]),
      showValue: true,
    })
  }

  clearSelected() {
    const next = clearRulesIn(this.c.sheet.conditional, this.c.selection.ranges)
    if (
      next.length !== this.c.sheet.conditional.length ||
      JSON.stringify(next) !== JSON.stringify(this.c.sheet.conditional)
    )
      this.commit(next, "Clear Rules")
  }

  clearSheet() {
    if (this.c.sheet.conditional.length) this.commit([], "Clear Rules")
  }

  /** Manage Rules › OK: the whole edited list. */
  replace(entries: CFEntry[]) {
    this.commit(entries, "Conditional Formatting")
  }

  setPreview(rule: CFRule | null) {
    this.preview = rule
      ? { id: -1, ranges: this.c.selection.ranges.map((r) => ({ ...r })), rule, stopIfTrue: false }
      : null
    this.c.changed()
  }

  /** Numbers in the selection, for the value a quick dialog suggests. */
  suggestion(kind: "greaterThan" | "lessThan" | "between" | "equal"): string[] {
    const values: number[] = []
    for (const r of this.c.selection.ranges)
      this.c.sheet.cells.forEach(r, (_row, _col, d) => typeof d.value === "number" && values.push(d.value), true)
    if (!values.length) return kind === "between" ? ["", ""] : [""]
    values.sort((a, b) => a - b)
    const round = (n: number) => String(Math.round(n * 100) / 100)
    const at = (p: number) => values[Math.min(values.length - 1, Math.floor(p * values.length))]
    if (kind === "between") return [round(at(0.25)), round(at(0.75))]
    if (kind === "equal") return [round(at(0))]
    return [round(at(0.5))]
  }
}
