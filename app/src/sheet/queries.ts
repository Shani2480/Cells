import {
  copyQueryTables,
  queryTableName,
  type CellWrite,
  type Op,
  type QueryTable,
  type Sheet,
  type StyleTable,
  type TextImport,
} from "@cells/core"
import { decodeText, placeCells, splitText } from "@cells/io"

/**
 * Data › Queries & Connections for text files: an import that stays linked to its file, so Refresh reads the file
 * again into the same place. Cells refreshes the way Excel's "Overwrite existing cells with new data, clear unused
 * cells" does: the new data starts at the range's top-left cell and what the old data covered beyond it is cleared.
 */

/** The cells a text file gives with the import's choices, from (0, 0), and the size of the block. */
export async function textCells(
  styles: StyleTable,
  bytes: Uint8Array,
  options: TextImport,
): Promise<{ cells: CellWrite[]; rows: number; cols: number }> {
  const rows = splitText(decodeText(bytes, options.encoding).text, options)
  const cells: CellWrite[] = []
  let height = 0
  let width = 0
  await placeCells(styles, rows, options, (row, col, data) => {
    cells.push({ row, col, data })
    height = Math.max(height, row + 1)
    width = Math.max(width, col + 1)
  })
  return { cells, rows: Math.max(1, height), cols: Math.max(1, width) }
}

/** A new link for an import: names from the file, unique on the sheet and in the workbook. */
export function newQueryTable(
  sheets: readonly Sheet[],
  fileName: string,
  source: string,
  options: TextImport,
  size: { rows: number; cols: number },
): QueryTable {
  const all = sheets.flatMap((s) => s.queryTables)
  const name = queryTableName(
    fileName,
    all.map((q) => q.name),
  )
  return {
    name,
    connection: queryTableName(
      fileName,
      all.map((q) => q.connection),
    ),
    description: "",
    range: { r0: 0, c0: 0, r1: size.rows - 1, c1: size.cols - 1 },
    source,
    options,
    refreshOnLoad: false,
  }
}

/** The query table whose range holds the cell. */
export function queryTableAt(sheet: Sheet, row: number, col: number): QueryTable | null {
  return (
    sheet.queryTables.find((q) => row >= q.range.r0 && row <= q.range.r1 && col >= q.range.c0 && col <= q.range.c1) ??
    null
  )
}

/** Refresh: the old range cleared where the new data doesn't reach, the new data written, the range resized. */
export function refreshOps(
  index: number,
  sheet: Sheet,
  name: string,
  data: { cells: CellWrite[]; rows: number; cols: number },
  source?: string,
): Op[] {
  const table = sheet.queryTables.find((q) => q.name === name)
  if (!table) return []
  const { r0, c0 } = table.range
  const cells: CellWrite[] = data.cells.map((c) => ({ row: r0 + c.row, col: c0 + c.col, data: c.data }))
  const written = new Set(cells.map((c) => `${c.row},${c.col}`))
  const old = table.range
  for (let row = old.r0; row <= old.r1; row++)
    for (let col = old.c0; col <= old.c1; col++)
      if (!written.has(`${row},${col}`) && sheet.cells.get(row, col)) cells.push({ row, col, data: undefined })
  // Cells inside the new block that the file left empty are cleared too.
  for (let row = r0; row < r0 + data.rows; row++)
    for (let col = c0; col < c0 + data.cols; col++)
      if (!written.has(`${row},${col}`) && sheet.cells.get(row, col) && !(row <= old.r1 && col <= old.c1))
        cells.push({ row, col, data: undefined })
  const tables = copyQueryTables(sheet.queryTables).map((q) =>
    q.name === name
      ? { ...q, source: source ?? q.source, range: { r0, c0, r1: r0 + data.rows - 1, c1: c0 + data.cols - 1 } }
      : q,
  )
  const ops: Op[] = []
  if (cells.length) ops.push({ kind: "setCells", sheet: index, cells })
  ops.push({ kind: "queryTables", sheet: index, tables })
  return ops
}

/** A change to one query table's settings (name, description, refresh on open), or its removal (`null`). */
export function queryTableOps(index: number, sheet: Sheet, name: string, next: QueryTable | null): Op[] {
  const tables = copyQueryTables(sheet.queryTables).flatMap((q) => (q.name !== name ? [q] : next ? [next] : []))
  return [{ kind: "queryTables", sheet: index, tables }]
}

/** The file's name, from a path, content URI or plain name. */
export function sourceName(source: string): string {
  const decoded = (() => {
    try {
      return decodeURIComponent(source)
    } catch {
      return source
    }
  })()
  return decoded.split(/[\\/:]/).pop() || source
}
