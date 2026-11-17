import { columnName, lastCell, type Sheet, type Workbook } from "@cells/core"

/** Review › Workbook Statistics: what a sheet, and the whole workbook, holds. */
export type SheetCounts = {
  cells: number
  tables: number
  formulas: number
  charts: number
  comments: number
  notes: number
}

export type WorkbookStatistics = {
  sheet: SheetCounts & { name: string; end: string }
  workbook: SheetCounts & { sheets: number }
}

export function sheetCounts(sheet: Sheet): SheetCounts {
  let formulas = 0
  sheet.cells.forEachFormula(() => formulas++)
  let comments = 0
  let notes = 0
  for (const a of sheet.annotations.all()) {
    if (a.value.kind === "comment") comments++
    else notes++
  }
  return {
    cells: sheet.cells.valueCount,
    tables: sheet.tables.length,
    formulas,
    charts: sheet.charts.length,
    comments,
    notes,
  }
}

export function workbookStatistics(wb: Workbook, active: number): WorkbookStatistics {
  const all = wb.sheets.map(sheetCounts)
  const sum = (key: keyof SheetCounts) => all.reduce((n, c) => n + c[key], 0)
  const sheet = wb.sheets[active]
  const end = lastCell(sheet)
  return {
    sheet: { ...all[active], name: sheet.name, end: columnName(end.col) + (end.row + 1) },
    workbook: {
      sheets: wb.sheets.length,
      cells: sum("cells"),
      tables: sum("tables"),
      formulas: sum("formulas"),
      charts: sum("charts"),
      comments: sum("comments"),
      notes: sum("notes"),
    },
  }
}
