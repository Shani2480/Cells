import type { TableStyleElement } from "./styles"

/**
 * The workbook's own table and PivotTable styles (Design › New Table Style / New PivotTable Style), Excel's
 * `<tableStyles>`: a name, which galleries show it, and the look of each element (keyed by Excel's element type, so
 * elements Cells doesn't draw are kept too).
 */
export type CustomTableStyle = {
  name: string
  /** Shown with table styles (Excel's table, default on). */
  table: boolean
  /** Shown with PivotTable styles (Excel's pivot, default on). */
  pivot: boolean
  elements: Record<string, TableStyleElement>
}

/** The styles with the defaults for new tables and PivotTables (Set As Default). */
export type TableStyleSet = {
  styles: CustomTableStyle[]
  /** None: TableStyleMedium2. */
  defaultTable: string | null
  /** None: PivotStyleLight16. */
  defaultPivot: string | null
}

export const DEFAULT_PIVOT_STYLE = "PivotStyleLight16"

export const copyCustomStyle = (s: CustomTableStyle): CustomTableStyle => ({
  ...s,
  elements: Object.fromEntries(
    Object.entries(s.elements).map(([k, e]) => [k, { ...e, ...(e.border ? { border: { ...e.border } } : {}) }]),
  ),
})

export const copyTableStyleSet = (set: TableStyleSet): TableStyleSet => ({
  ...set,
  styles: set.styles.map(copyCustomStyle),
})

/** The workbook's style of a name (names ignore case, as Excel's), shown with `kind`'s styles. */
export function customStyle(
  styles: readonly CustomTableStyle[],
  name: string,
  kind: "table" | "pivot",
): CustomTableStyle | undefined {
  const lower = name.toLowerCase()
  return styles.find((s) => s[kind] && s.name.toLowerCase() === lower)
}

/** The elements of New PivotTable Style's Table Element list, in Excel's order, with its names. */
export const PIVOT_STYLE_ELEMENTS: { key: string; label: string; stripe?: true }[] = [
  { key: "wholeTable", label: "Whole Table" },
  { key: "pageFieldLabels", label: "Report Filter Labels" },
  { key: "pageFieldValues", label: "Report Filter Values" },
  { key: "firstColumnStripe", label: "First Column Stripe", stripe: true },
  { key: "secondColumnStripe", label: "Second Column Stripe", stripe: true },
  { key: "firstRowStripe", label: "First Row Stripe", stripe: true },
  { key: "secondRowStripe", label: "Second Row Stripe", stripe: true },
  { key: "firstColumn", label: "First Column" },
  { key: "headerRow", label: "Header Row" },
  { key: "firstHeaderCell", label: "First Header Cell" },
  { key: "subtotalColumn1", label: "Subtotal Column 1" },
  { key: "subtotalColumn2", label: "Subtotal Column 2" },
  { key: "subtotalColumn3", label: "Subtotal Column 3" },
  { key: "subtotalRow1", label: "Subtotal Row 1" },
  { key: "subtotalRow2", label: "Subtotal Row 2" },
  { key: "subtotalRow3", label: "Subtotal Row 3" },
  { key: "blankRow", label: "Blank Row" },
  { key: "columnSubheading1", label: "Column Subheading 1" },
  { key: "columnSubheading2", label: "Column Subheading 2" },
  { key: "columnSubheading3", label: "Column Subheading 3" },
  { key: "rowSubheading1", label: "Row Subheading 1" },
  { key: "rowSubheading2", label: "Row Subheading 2" },
  { key: "rowSubheading3", label: "Row Subheading 3" },
  { key: "totalRow", label: "Grand Total Row" },
  { key: "lastColumn", label: "Grand Total Column" },
]

/** Excel's messages for a style name. */
export const STYLE_NAME_TAKEN = "This style name already exists."
export const STYLE_NAME_EMPTY = "Type a name for the style."

/**
 * Why a style can't take `name` (null: it can): empty, or the name of another style, the workbook's or a built-in
 * one (`builtIn` tells those; `except` is the style being modified).
 */
export function customStyleNameProblem(
  name: string,
  styles: readonly CustomTableStyle[],
  builtIn: (name: string) => boolean,
  except?: string,
): string | null {
  const n = name.trim()
  if (!n) return STYLE_NAME_EMPTY
  const lower = n.toLowerCase()
  if (lower === except?.toLowerCase()) return null
  if (builtIn(n) || styles.some((s) => s.name.toLowerCase() === lower)) return STYLE_NAME_TAKEN
  return null
}

/** The name New PivotTable Style offers: "PivotTable Style 1", 2… the first one free. */
export function newCustomStyleName(styles: readonly CustomTableStyle[], base = "PivotTable Style"): string {
  const taken = new Set(styles.map((s) => s.name.toLowerCase()))
  let n = 1
  while (taken.has(`${base} ${n}`.toLowerCase())) n++
  return `${base} ${n}`
}

/** The name Duplicate offers: the style's name with " 2" (" 3"… while taken). */
export function duplicateStyleName(name: string, styles: readonly CustomTableStyle[]): string {
  const taken = new Set(styles.map((s) => s.name.toLowerCase()))
  let n = 2
  while (taken.has(`${name} ${n}`.toLowerCase())) n++
  return `${name} ${n}`
}

/** What Element Formatting says about an element ("Bold; Fill; Borders"), empty when it has no format. */
export function elementSummary(e: TableStyleElement | undefined): string {
  if (!e) return ""
  const out: string[] = []
  const font = [
    e.bold ? "Bold" : e.bold === false ? "Not Bold" : "",
    e.italic ? "Italic" : e.italic === false ? "Not Italic" : "",
    e.underline && e.underline !== "none" ? "Underline" : "",
    e.strike ? "Strikethrough" : "",
    e.color ? "Font Color" : "",
  ].filter(Boolean)
  out.push(...font)
  if (e.fill) out.push("Shaded")
  if (e.border && Object.values(e.border).some(Boolean)) out.push("Bordered")
  if (e.stripe && e.stripe > 1) out.push(`Stripe Size ${e.stripe}`)
  return out.join("; ")
}

/** An element with nothing set (Clear): none. */
export function isEmptyElement(e: TableStyleElement): boolean {
  return !Object.entries(e).some(([k, v]) =>
    k === "border" ? Object.values(v as object).some(Boolean) : k === "stripe" ? (v as number) > 1 : v !== undefined,
  )
}
