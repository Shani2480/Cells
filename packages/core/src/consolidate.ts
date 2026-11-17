import { columnName, type Range } from "./address"
import { contentForInput } from "./edit"
import type { AxisEntry, CellWrite, Op } from "./ops"
import type { Sheet } from "./sheet"
import type { StyleTable } from "./styles"
import type { CellValue } from "./values"

/**
 * Data › Data Tools › Consolidate: several ranges summarised into one, by position or by the labels in their top
 * row and left column, with one function (Sum, Average, …). With "Create links to source data" each result is a
 * formula over detail rows that link to the source cells, grouped in an outline so level 1 shows only the results.
 */

/** The dialog's Function list, its name in Excel files and the worksheet function a linked summary uses. */
export const CONSOLIDATE_FUNCTIONS: { label: string; xml: string; formula: string }[] = [
  { label: "Sum", xml: "sum", formula: "SUM" },
  { label: "Count", xml: "count", formula: "COUNTA" },
  { label: "Average", xml: "average", formula: "AVERAGE" },
  { label: "Max", xml: "max", formula: "MAX" },
  { label: "Min", xml: "min", formula: "MIN" },
  { label: "Product", xml: "product", formula: "PRODUCT" },
  { label: "Count Numbers", xml: "countNums", formula: "COUNT" },
  { label: "StdDev", xml: "stdDev", formula: "STDEV" },
  { label: "StdDevp", xml: "stdDevp", formula: "STDEVP" },
  { label: "Var", xml: "var", formula: "VAR" },
  { label: "Varp", xml: "varp", formula: "VARP" },
]

/**
 * What the dialog last used on a sheet, saved with it (Excel's `dataConsolidate`): the references as typed (so a
 * renamed sheet is not followed, as in Excel), the function and the options.
 */
export type ConsolidateSettings = {
  /** Index into CONSOLIDATE_FUNCTIONS. */
  fn: number
  refs: string[]
  top: boolean
  left: boolean
  links: boolean
}

export const copyConsolidate = (s: ConsolidateSettings | null): ConsolidateSettings | null =>
  s ? { ...s, refs: [...s.refs] } : null

export type ConsolidateSource = { sheet: number; range: Range }

export type ConsolidateOptions = { fn: number; top: boolean; left: boolean; links: boolean }

/** A source cell that adds to a result. */
type Part = { source: number; row: number; col: number }

/** The value Consolidate works out of the values that meet in one place. */
export function combine(fn: number, values: CellValue[]): number | undefined {
  const nums = values.filter((v): v is number => typeof v === "number")
  const n = nums.length
  const sum = nums.reduce((a, b) => a + b, 0)
  const variance = (sample: boolean) => {
    const d = sample ? n - 1 : n
    if (d <= 0) return undefined
    const mean = sum / n
    return nums.reduce((a, b) => a + (b - mean) ** 2, 0) / d
  }
  switch (CONSOLIDATE_FUNCTIONS[fn]?.xml ?? "sum") {
    case "count":
      return values.length || undefined
    case "countNums":
      return n || undefined
    case "average":
      return n ? sum / n : undefined
    case "max":
      return n ? Math.max(...nums) : undefined
    case "min":
      return n ? Math.min(...nums) : undefined
    case "product":
      return n ? nums.reduce((a, b) => a * b, 1) : undefined
    case "stdDev": {
      const v = variance(true)
      return v === undefined ? undefined : Math.sqrt(v)
    }
    case "stdDevp": {
      const v = variance(false)
      return v === undefined ? undefined : Math.sqrt(v)
    }
    case "var":
      return variance(true)
    case "varp":
      return variance(false)
    default:
      return n ? sum : undefined
  }
}

const labelText = (v: CellValue | undefined) =>
  v === undefined ? "" : typeof v === "object" ? v.error : String(v).trim()

/** One axis of the result: by label (in the order first seen) or by position. */
class Keys {
  readonly labels: { key: string; text: CellValue }[] = []
  private index = new Map<string, number>()
  size = 0

  /** The result line for a label; -1 for an empty label, which Excel leaves out. */
  label(value: CellValue | undefined): number {
    const text = labelText(value)
    if (!text) return -1
    const key = text.toLowerCase()
    let at = this.index.get(key)
    if (at === undefined) {
      at = this.labels.length
      this.index.set(key, at)
      this.labels.push({ key, text: value as CellValue })
    }
    this.size = Math.max(this.size, this.labels.length)
    return at
  }

  position(i: number): number {
    this.size = Math.max(this.size, i + 1)
    return i
  }
}

export type Consolidation = {
  /** Result rows and columns (labels not counted). */
  rows: number
  cols: number
  rowLabels: CellValue[] | null
  colLabels: CellValue[] | null
  /** The source cells that meet at each result, by `row * cols + col`. */
  parts: Map<number, Part[]>
}

/** Where every source cell lands in the result. */
export function consolidation(sheets: readonly Sheet[], sources: ConsolidateSource[], o: ConsolidateOptions) {
  const rows = new Keys()
  const cols = new Keys()
  const placed: Part[][] = []
  const at: [number, number][][] = []
  sources.forEach(({ sheet: s, range }, source) => {
    const sheet = sheets[s]
    const dataR0 = range.r0 + (o.top ? 1 : 0)
    const dataC0 = range.c0 + (o.left ? 1 : 0)
    const colAt: number[] = []
    for (let c = dataC0; c <= range.c1; c++)
      colAt.push(o.top ? cols.label(sheet.cells.value(range.r0, c)) : cols.position(c - dataC0))
    const cells: Part[] = []
    const where: [number, number][] = []
    for (let r = dataR0; r <= range.r1; r++) {
      const row = o.left ? rows.label(sheet.cells.value(r, range.c0)) : rows.position(r - dataR0)
      if (row < 0) continue
      for (let c = dataC0; c <= range.c1; c++) {
        const col = colAt[c - dataC0]
        if (col < 0) continue
        cells.push({ source, row: r, col: c })
        where.push([row, col])
      }
    }
    placed.push(cells)
    at.push(where)
  })
  const parts = new Map<number, Part[]>()
  placed.forEach((cells, s) =>
    cells.forEach((p, i) => {
      const [row, col] = at[s][i]
      const value = sheets[sources[p.source].sheet].cells.value(p.row, p.col)
      if (value === undefined || value === "") return
      const key = row * cols.size + col
      const list = parts.get(key)
      if (list) list.push(p)
      else parts.set(key, [p])
    }),
  )
  const result: Consolidation = {
    rows: rows.size,
    cols: cols.size,
    rowLabels: o.left ? rows.labels.map((l) => l.text) : null,
    colLabels: o.top ? cols.labels.map((l) => l.text) : null,
    parts,
  }
  return result
}

const quoteSheet = (name: string) => (/^[A-Za-z_][\w.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
const absRef = (row: number, col: number) => `$${columnName(col)}$${row + 1}`

/**
 * Consolidate into `dest` on sheet `index`: its top-left cell, or (more than one cell) only that area. Values are
 * written (formulas with links), keeping each destination cell's format. Returns the operations and the area filled,
 * or Excel's message when it can't.
 */
export function consolidateOps(
  sheets: readonly Sheet[],
  styles: StyleTable,
  index: number,
  dest: Range,
  sources: ConsolidateSource[],
  o: ConsolidateOptions,
): { ops: Op[]; area: Range } | { problem: string } {
  if (!sources.length) return { problem: "Add at least one reference to consolidate." }
  const sheet = sheets[index]
  if (o.links && sources.some((s) => s.sheet === index))
    return { problem: "Links can't be created to source data on the same sheet as the consolidation." }
  const c = consolidation(sheets, sources, o)
  if (!c.rows || !c.cols) return { problem: "No data was consolidated." }
  const fn = CONSOLIDATE_FUNCTIONS[o.fn] ?? CONSOLIDATE_FUNCTIONS[0]
  const limit = dest.r1 > dest.r0 || dest.c1 > dest.c0 ? dest : null
  const r0 = dest.r0
  const c0 = dest.c0
  // With links, a column after the labels names each detail row's source.
  const labelCols = (o.left ? 1 : 0) + (o.links ? 1 : 0)
  const dataC0 = c0 + labelCols
  const cells: CellWrite[] = []
  const inside = (row: number, col: number) =>
    !limit || (row >= limit.r0 && row <= limit.r1 && col >= limit.c0 && col <= limit.c1)
  const write = (row: number, col: number, value: CellValue | undefined, formula?: string) => {
    if (!inside(row, col)) return
    const current = sheet.cells.get(row, col)
    if (formula) {
      const data = contentForInput(styles, current, formula)
      if (data) cells.push({ row, col, data })
      return
    }
    const style = current?.style
    const data = value === undefined ? (style ? { style } : undefined) : { value, ...(style ? { style } : {}) }
    cells.push({ row, col, data })
  }
  let row = r0
  if (c.colLabels) {
    for (let k = 0; k < labelCols; k++) write(row, c0 + k, undefined)
    c.colLabels.forEach((label, j) => write(row, dataC0 + j, label))
    row++
  }
  const hidden: number[] = []
  const summaries: number[] = []
  for (let i = 0; i < c.rows; i++) {
    const label = c.rowLabels?.[i]
    if (!o.links) {
      if (o.left) write(row, c0, label)
      for (let j = 0; j < c.cols; j++) {
        const parts = c.parts.get(i * c.cols + j) ?? []
        write(
          row,
          dataC0 + j,
          combine(
            o.fn,
            parts.map((p) => sheets[sources[p.source].sheet].cells.value(p.row, p.col)!),
          ),
        )
      }
      row++
      continue
    }
    // Links: one detail row per source row that adds to this result row, then the summary over them.
    const detail = new Map<string, { source: number; row: number }>()
    for (let j = 0; j < c.cols; j++)
      for (const p of c.parts.get(i * c.cols + j) ?? [])
        detail.set(`${p.source}:${p.row}`, { source: p.source, row: p.row })
    const lines = [...detail.values()].sort((a, b) => a.source - b.source || a.row - b.row)
    const first = row
    for (const line of lines) {
      const src = sources[line.source]
      const name = sheets[src.sheet].name
      write(row, c0 + labelCols - 1, name)
      if (o.left) write(row, c0, undefined)
      for (let j = 0; j < c.cols; j++) {
        const part = (c.parts.get(i * c.cols + j) ?? []).find((p) => p.source === line.source && p.row === line.row)
        write(row, dataC0 + j, undefined, part ? `=${quoteSheet(name)}!${absRef(part.row, part.col)}` : undefined)
      }
      hidden.push(row)
      row++
    }
    if (o.left) write(row, c0, label)
    write(row, c0 + labelCols - 1, o.left && labelCols === 1 ? label : undefined)
    for (let j = 0; j < c.cols; j++) {
      const col = dataC0 + j
      const any = (c.parts.get(i * c.cols + j) ?? []).length > 0
      const span = `${columnName(col)}${first + 1}:${columnName(col)}${row}`
      write(row, col, undefined, any && row > first ? `=${fn.formula}(${span})` : undefined)
    }
    summaries.push(row)
    row++
  }
  const area: Range = {
    r0,
    c0,
    r1: limit ? Math.min(limit.r1, row - 1) : row - 1,
    c1: limit ? Math.min(limit.c1, dataC0 + c.cols - 1) : dataC0 + c.cols - 1,
  }
  if (sources.some((s) => s.sheet === index && overlaps(s.range, area)))
    return { problem: "Source references overlap the destination area." }
  const ops: Op[] = [{ kind: "setCells", sheet: index, cells }]
  if (o.links && hidden.length) {
    const entries: AxisEntry[] = []
    const rowEntry = (i: number, change: Partial<AxisEntry>): AxisEntry => ({
      index: i,
      size: sheet.rows.hasCustomSize(i) ? sheet.rows.storedSize(i) : null,
      hidden: sheet.rows.isHidden(i),
      ...change,
    })
    for (const r of hidden) entries.push(rowEntry(r, { hidden: true, level: 1 }))
    for (const r of summaries) entries.push(rowEntry(r, { collapsed: true, level: 0 }))
    ops.push({ kind: "axis", sheet: index, axis: "rows", entries })
  }
  return { ops, area }
}

function overlaps(a: Range, b: Range) {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
}
