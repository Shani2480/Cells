/**
 * Excel's Data Model (Power Pivot), as far as Cells can see it: the tables in the model and the relationships between
 * them. Excel keeps the model itself in a binary part Cells can't write, and trusts only that part, so Cells shows the
 * model (Data › Relationships, Manage Data Model) and keeps it as it was; it's changed in Excel.
 */

export type ModelTable = {
  /** The model's name for the table. */
  name: string
  /** The workbook connection that loads it. */
  connection: string
  /** The workbook table it's loaded from ("Orders"), or null when it comes from elsewhere (a query, a database). */
  source: string | null
}

/** Many rows of `fromTable` look up one row of `toTable` (Excel's foreign key → primary key). */
export type ModelRelationship = { fromTable: string; fromColumn: string; toTable: string; toColumn: string }

export type DataModel = { tables: ModelTable[]; relationships: ModelRelationship[] }

/** The workbook table a worksheet connection's range name points at ("Orders[]" or "Orders[#All]" → "Orders"). */
export function sourceTable(formula: string | undefined): string | null {
  const m = /^=?'?([^[\]'!]+)'?\[(?:#All|#Data)?\]$/i.exec(formula?.trim() ?? "")
  return m ? m[1] : null
}
