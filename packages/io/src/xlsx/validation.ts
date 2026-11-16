import {
  DEFAULT_VALIDATION,
  parseRange,
  type Range,
  type Validation,
  type ValidationEntry,
  type ValidationOperator,
  type ValidationType,
} from "@cells/core"
import { escapeXmlText } from "@cells/ooxml"
import { at, attrText, kid, kids, parseFragment, type XNode } from "./dom"
import { formulaFromFile, formulaToFile } from "./formula"

/**
 * A worksheet's `<dataValidations>`, as Excel writes them. A rule whose formula points at another sheet goes in the
 * x14 extension instead (`<x14:dataValidation>` with `<xm:f>` and `<xm:sqref>`), as Excel 2010 and later write it.
 */

export const X14_VALIDATION_URI = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"
const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"

const TYPES = new Set<ValidationType>(["whole", "decimal", "list", "date", "time", "textLength", "custom"])
const OPERATORS = new Set<ValidationOperator>([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "lessThan",
  "greaterThanOrEqual",
  "lessThanOrEqual",
])

const flag = (v: string | undefined) => v === "1" || v === "true"
const fromFile = (text: string) => (text ? formulaFromFile(text).slice(1) : "")

function ranges(sqref: string): Range[] {
  return sqref
    .split(/\s+/)
    .map((r) => parseRange(r))
    .filter((r): r is Range => r !== null)
}

/** One `<dataValidation>` (or `<x14:dataValidation>`): its rule; the caller gives the ranges. */
function ruleOf(node: XNode, formula: (n: XNode | undefined) => string): Validation {
  const type = at(node, "type") as ValidationType | undefined
  const operator = at(node, "operator") as ValidationOperator | undefined
  const style = at(node, "errorStyle")
  return {
    ...DEFAULT_VALIDATION,
    type: type && TYPES.has(type) ? type : "any",
    operator: operator && OPERATORS.has(operator) ? operator : "between",
    formula1: formula(kid(node, "formula1")),
    formula2: formula(kid(node, "formula2")),
    allowBlank: flag(at(node, "allowBlank")),
    // showDropDown="1" hides the in-cell dropdown (the attribute's sense is reversed).
    dropDown: !flag(at(node, "showDropDown")),
    showInput: flag(at(node, "showInputMessage")),
    promptTitle: at(node, "promptTitle") ?? "",
    prompt: at(node, "prompt") ?? "",
    showError: flag(at(node, "showErrorMessage")),
    errorStyle: style === "warning" || style === "information" ? style : "stop",
    errorTitle: at(node, "errorTitle") ?? "",
    error: at(node, "error") ?? "",
  }
}

/** The rules of `<dataValidations>`. */
export function readValidations(root: XNode): ValidationEntry[] {
  const out: ValidationEntry[] = []
  for (const node of kids(root, "dataValidation")) {
    const r = ranges(at(node, "sqref") ?? "")
    if (r.length) out.push({ ranges: r, rule: ruleOf(node, (n) => fromFile(n?.text ?? "")) })
  }
  return out
}

/**
 * The x14 rules in an `<extLst>`, and the extLst without them (empty when nothing else is left), so Cells writes
 * the rules itself.
 */
export function readValidationExt(source: string, rootTag: string): { entries: ValidationEntry[]; rest: string } {
  if (!source.includes(X14_VALIDATION_URI)) return { entries: [], rest: source }
  const entries: ValidationEntry[] = []
  const doc = parseFragment(source, rootTag)
  for (const ext of kids(doc.root, "ext")) {
    if (at(ext, "uri") !== X14_VALIDATION_URI) continue
    for (const node of kids(kid(ext, "dataValidations"), "dataValidation")) {
      const r = ranges(kid(node, "sqref")?.text ?? "")
      const formula = (n: XNode | undefined) => fromFile(kid(n, "f")?.text ?? "")
      if (r.length) entries.push({ ranges: r, rule: ruleOf(node, formula) })
    }
  }
  const rest = source.replace(/<(\w+:)?ext\b[^>]*\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}[\s\S]*?<\/\1?ext>/gi, "")
  return { entries, rest: /<(\w+:)?ext\b/.test(rest) ? rest : "" }
}

/** Whether a rule must go in the x14 extension: a formula that names another sheet. */
const needsExt = (rule: Validation) => [rule.formula1, rule.formula2].some((f) => /!/.test(f.replace(/"[^"]*"/g, "")))

function attrsOf(rule: Validation) {
  return attrText({
    type: rule.type === "any" ? undefined : rule.type,
    errorStyle: rule.errorStyle === "stop" ? undefined : rule.errorStyle,
    operator: rule.operator === "between" ? undefined : rule.operator,
    allowBlank: rule.allowBlank,
    showDropDown: rule.type === "list" && !rule.dropDown,
    showInputMessage: rule.showInput,
    showErrorMessage: rule.showError,
    errorTitle: rule.errorTitle || undefined,
    error: rule.error || undefined,
    promptTitle: rule.promptTitle || undefined,
    prompt: rule.prompt || undefined,
  })
}

/** `<dataValidations>` for the rules kept in the sheet itself ("" when none). */
export function validationsXml(entries: readonly ValidationEntry[], rangeRef: (r: Range) => string): string {
  const own = entries.filter((e) => !needsExt(e.rule))
  if (!own.length) return ""
  const parts = own.map((e) => {
    const f = (tag: string, text: string) => (text ? `<${tag}>${escapeXmlText(formulaToFile(text))}</${tag}>` : "")
    const body = f("formula1", e.rule.formula1) + (needsTwo(e.rule) ? f("formula2", e.rule.formula2) : "")
    const sqref = e.ranges.map(rangeRef).join(" ")
    return body
      ? `<dataValidation${attrsOf(e.rule)} sqref="${sqref}">${body}</dataValidation>`
      : `<dataValidation${attrsOf(e.rule)} sqref="${sqref}"/>`
  })
  return `<dataValidations count="${own.length}">${parts.join("")}</dataValidations>`
}

/** The x14 `<ext>` for rules that point at other sheets ("" when none). */
export function validationExtXml(entries: readonly ValidationEntry[], rangeRef: (r: Range) => string): string {
  const ext = entries.filter((e) => needsExt(e.rule))
  if (!ext.length) return ""
  const parts = ext.map((e) => {
    const f = (tag: string, text: string) =>
      text ? `<x14:${tag}><xm:f>${escapeXmlText(formulaToFile(text))}</xm:f></x14:${tag}>` : ""
    const body = f("formula1", e.rule.formula1) + (needsTwo(e.rule) ? f("formula2", e.rule.formula2) : "")
    return `<x14:dataValidation${attrsOf(e.rule)}>${body}<xm:sqref>${e.ranges.map(rangeRef).join(" ")}</xm:sqref></x14:dataValidation>`
  })
  return (
    `<ext uri="${X14_VALIDATION_URI}" xmlns:x14="${NS_X14}">` +
    `<x14:dataValidations count="${ext.length}" xmlns:xm="${NS_XM}">${parts.join("")}</x14:dataValidations></ext>`
  )
}

const needsTwo = (rule: Validation) =>
  (rule.operator === "between" || rule.operator === "notBetween") &&
  rule.type !== "list" &&
  rule.type !== "custom" &&
  rule.type !== "any"

/** Puts an `<ext>` into the extLst kept from the file, or makes one. */
export function withExt(extLst: string | undefined, ext: string): string {
  if (!ext) return extLst ?? ""
  if (!extLst) return `<extLst>${ext}</extLst>`
  return extLst.replace(/<\/((\w+:)?extLst)>\s*$/, `${ext}</$1>`)
}
