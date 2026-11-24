import type { CellAddress, Range } from "../address"
import type { CellValue } from "../values"

/**
 * PivotTables (Phase 7, Part 10; plan in docs/pivot-plan.md). The model follows Excel's own parts
 * (pivotCacheDefinition, pivotCacheRecords, pivotTableDefinition) so a file reads in and writes back without losing
 * what it says, with Excel's names where they are clear and plain ones where they aren't.
 */

/** A value in a PivotTable's cache: what a source cell held (null: empty). */
export type PivotValue = CellValue | null

/** One column of the source data, as the cache keeps it. */
export type PivotCacheField = {
  name: string
  /**
   * The distinct values the PivotTable's items point at (Excel's shared items), in the order first met. Null: not
   * kept (a number column nobody has put on an axis); the engine finds the values in the records when it needs them.
   */
  items: PivotValue[] | null
  /** The column's number format (dates show as dates in the labels). */
  numFmt?: string
  /** The column holds dates (written as dates in the cache). */
  dates?: boolean
  /** A calculated field's formula over the other fields (=Amount/Units); it has no values in the records. */
  formula?: string
  /**
   * Grouping: number ranges on the field itself (`base` its own index), or a field of its own made from `base` — dates
   * by Months, Quarters and so on, or items gathered by Group Selection. A field of its own has no values in the
   * records.
   */
  group?: PivotFieldGroup
}

/** What a range group gathers by: numbers, or a part of a date. */
export type PivotGroupBy = "range" | "seconds" | "minutes" | "hours" | "days" | "months" | "quarters" | "years"

export type PivotFieldGroup =
  | {
      kind: "range"
      base: number
      by: PivotGroupBy
      /** Numbers, or date serials. */
      start: number
      end: number
      /** The size of each range (numbers), or days per item (`by` "days"). */
      interval: number
      autoStart: boolean
      autoEnd: boolean
      /** The group items: "<start", the ranges, ">end". */
      items: PivotValue[]
    }
  | {
      kind: "discrete"
      base: number
      /** The group item of each of the base field's shared items. */
      map: number[]
      items: PivotValue[]
    }

/** Where the cache reads its data: a range on a sheet, or a table or defined name. */
export type PivotCacheSource = { kind: "range"; sheet: string; ref: string } | { kind: "name"; name: string }

export type PivotCache = {
  id: number
  source: PivotCacheSource
  fields: PivotCacheField[]
  /** One row per source row, its values in the order of `fields`. */
  records: PivotValue[][]
  /** Refresh when the file opens (Excel's "Refresh data when opening the file"). */
  refreshOnLoad: boolean
  /** Excel 2010's id for the cache, which PivotTable slicers name it by (x14 pivotCacheId). */
  slicerId?: number
  /**
   * Calculated items (Fields, Items & Sets › Calculated Item), in solve order: the last one listed wins where two
   * meet. Each is a shared item of its field (`x`) that no record points at.
   */
  calcItems?: PivotCalcItem[]
}

/** A calculated item: its field, its shared item, and its formula as a file keeps it (Product[Pens]+Product[Ink]*2). */
export type PivotCalcItem = { field: number; x: number; formula: string }

/** How a value field summarizes (Value Field Settings › Summarize Values By). */
export type PivotFn =
  "sum" | "count" | "average" | "max" | "min" | "product" | "countNums" | "stdDev" | "stdDevp" | "var" | "varp"

/** A field's subtotal: "default" uses each value field's own function (Field Settings › Automatic). */
export type PivotSubtotal = "default" | PivotFn

export type PivotSort = "manual" | "ascending" | "descending"

/** One item of a field, in the order the PivotTable shows it. `x`: the cache's shared item it stands for. */
export type PivotItem = {
  x: number
  /** Unticked in the field's filter. */
  hidden?: true
  /** Collapsed (−): its inner items are hidden and its line shows the totals. */
  collapsed?: true
  /** A name typed over the item's label. */
  name?: string
}

/** The PivotTable's settings for one cache field. */
export type PivotField = {
  axis: "row" | "col" | "page" | null
  /** Used by at least one value field. */
  dataField: boolean
  /** A name typed over the field's own (Field Settings › Custom Name). */
  name?: string
  items: PivotItem[]
  /** [] for none (Field Settings › Subtotals › None). */
  subtotals: PivotSubtotal[]
  /** Subtotals at the top of each group (outline and compact forms only). */
  subtotalTop: boolean
  /** Outline form (false: tabular). */
  outline: boolean
  /** Compact form: the next field's labels go in this field's column. */
  compact: boolean
  /** A blank line after each item. */
  insertBlankRow: boolean
  /** Show items with no data. */
  showAll: boolean
  sort: PivotSort
  /**
   * Sort by a value instead of the labels (More Sort Options › Ascending/Descending by "Sum of Amount"): the index in
   * `values` of the value field whose totals order the items, within the item above. `sort` gives the direction.
   */
  sortBy?: number
  /** Repeat item labels on every line. */
  repeatLabels: boolean
  /** The report filter lets more than one item be picked. */
  multipleItems: boolean
  numFmt?: string
}

/** A value field (the Values area). */
export type PivotDataField = {
  name: string
  field: number
  fn: PivotFn
  numFmt?: string
  /** Show Values As (left out: No Calculation). */
  showAs?: PivotShowAs
  /** The field Show Values As compares within (an index into the cache's fields). */
  baseField?: number
  /** The item it compares with: a cache item of `baseField`, or the one before or after each item. */
  baseItem?: number | "previous" | "next"
}

/**
 * Value Field Settings › Show Values As: Excel's showDataAs choices and the four Excel 2010 added (percentOfParent*,
 * percentOfRunningTotal, rank*), which a file keeps in an x14 extension.
 */
export type PivotShowAs =
  | "percentOfTotal"
  | "percentOfCol"
  | "percentOfRow"
  | "percent"
  | "percentOfParentRow"
  | "percentOfParentCol"
  | "percentOfParent"
  | "difference"
  | "percentDiff"
  | "runTotal"
  | "percentOfRunningTotal"
  | "rankAscending"
  | "rankDescending"
  | "index"

/** The choices that compare with a base field, and those that also need a base item. */
export const SHOW_AS_BASE_FIELD: readonly PivotShowAs[] = [
  "percent",
  "percentOfParent",
  "difference",
  "percentDiff",
  "runTotal",
  "percentOfRunningTotal",
  "rankAscending",
  "rankDescending",
]
export const SHOW_AS_BASE_ITEM: readonly PivotShowAs[] = ["percent", "difference", "percentDiff"]

/** A report filter: the item picked (an index into its field's items), or null for (All) / several items. */
export type PivotPage = { field: number; item: number | null }

/** How a Label Filter compares an item's label (Excel's caption filters; text is matched ignoring case). */
export type PivotLabelOp =
  | "equal"
  | "notEqual"
  | "beginsWith"
  | "notBeginsWith"
  | "endsWith"
  | "notEndsWith"
  | "contains"
  | "notContains"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual"
  | "between"
  | "notBetween"

/** How a Value Filter compares an item's total. */
export type PivotValueOp =
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual"
  | "between"
  | "notBetween"

/**
 * A Label Filter, Value Filter or Top 10 on a row or column field (Excel's `<filter>`). Value filters and Top 10
 * judge each item by its total of value field `measure` (an index into `values`) within the item above it.
 */
export type PivotFilter = { field: number; id: number } & (
  | { kind: "label"; op: PivotLabelOp; value: string; value2?: string }
  | { kind: "value"; op: PivotValueOp; measure: number; value: number; value2?: number }
  | { kind: "top"; by: "count" | "percent" | "sum"; top: boolean; measure: number; value: number }
  /** A timeline's pick: dates from `value` to `value2` (serials, whole days, both ends in), on any field. */
  | { kind: "date"; op: "between"; value: number; value2: number }
)

/** The Values pseudo field's place among the row or column fields ("Σ Values"). */
export const VALUES_FIELD = -2

export type PivotTableStyle = {
  name: string | null
  rowHeaders: boolean
  colHeaders: boolean
  rowStripes: boolean
  colStripes: boolean
  lastColumn: boolean
}

export type PivotTable = {
  name: string
  cacheId: number
  /** The top-left cell of the table (below the report filters). */
  at: CellAddress
  /**
   * The table's cells as last laid out (Excel's location), without the report filters: what the sheet's cells show
   * and what edits must leave alone until the next refresh.
   */
  location: Range
  /** One per cache field. */
  fields: PivotField[]
  /** Row and column fields in order; VALUES_FIELD marks where the value fields go. */
  rows: number[]
  cols: number[]
  pages: PivotPage[]
  values: PivotDataField[]
  /** The Grand Total row at the bottom (Excel's rowGrandTotals). */
  rowGrandTotals: boolean
  /** The Grand Total column at the right (Excel's colGrandTotals). */
  colGrandTotals: boolean
  /** Defaults the table's fields start with, and the Values field's form. */
  compact: boolean
  outline: boolean
  compactData: boolean
  captions: {
    /** "Values" by default. */
    data: string
    grandTotal?: string
    rowHeader?: string
    colHeader?: string
    /** Shown for errors when `showError`. */
    error?: string
    /** Shown in empty value cells. */
    missing?: string
  }
  showError: boolean
  showMissing: boolean
  /**
   * PivotTable Options › Display › Show the Values row off (Excel 365's default for new PivotTables): with the value
   * fields as the only column field, their names head the columns with no row above them.
   */
  hideValuesRow?: true
  /** Label Filters, Value Filters and Top 10, in the order they were set. */
  filters?: PivotFilter[]
  /** PivotTable Options › Data › Enable show details (Excel's enableDrill); false only when off. */
  enableDrill?: false
  /**
   * PivotTable Options › Layout & Format › Autofit column widths on update (Excel's useAutoFormatting, off when a file
   * leaves it out); false only when off. New PivotTables have it on.
   */
  autoFit?: false
  /** PivotTable Options › Display › Show expand/collapse buttons (Excel's showDrill); false only when off. */
  showDrill?: false
  /** Analyze › Show › Field Headers (Excel's showHeaders): false hides the Row/Column Labels headers and their buttons. */
  showHeaders?: false
  style: PivotTableStyle
  /**
   * The file part the PivotTable was read from: saving finds what Cells keeps unread there by it, so a renamed
   * PivotTable keeps it too.
   */
  part?: string
}

export const PIVOT_FN_LABEL: Record<PivotFn, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
  product: "Product",
  countNums: "Count Numbers",
  stdDev: "StdDev",
  stdDevp: "StdDevp",
  var: "Var",
  varp: "Varp",
}

/** A field as Excel makes it for a new PivotTable: on no axis, items in the cache's order. */
export function newPivotField(itemCount: number, table?: Pick<PivotTable, "compact" | "outline">): PivotField {
  return {
    axis: null,
    dataField: false,
    items: Array.from({ length: itemCount }, (_, x) => ({ x })),
    subtotals: ["default"],
    subtotalTop: true,
    outline: table?.outline ?? true,
    compact: table?.compact ?? true,
    insertBlankRow: false,
    showAll: false,
    sort: "manual",
    repeatLabels: false,
    multipleItems: false,
  }
}
