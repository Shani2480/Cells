import { formatMark, formatValue, formatWithCode, hasFormatMarks, type CellValue, type Workbook } from "@cells/core"
import { encodeText, fits1252, type TextEncoding } from "./encoding"
import { joinDelimited } from "./split"

export type TextSave = {
  delimiter: string
  encoding: TextEncoding
  bom: boolean
}

/** Excel's text formats by extension: CSV UTF-8 for .csv, Text (Tab delimited) for .txt and .tsv (UTF-8 here too). */
export function textSaveFor(fileName: string): TextSave {
  return /\.csv$/i.test(fileName)
    ? { delimiter: ",", encoding: "utf-8", bom: true }
    : { delimiter: "\t", encoding: "utf-8", bom: true }
}

/**
 * One sheet as CSV or delimited text, as Excel saves it: each cell's shown text (number formats applied, General in
 * a default-width column, so 123456789012 is 1.23457E+11), from A1 to the end of the used range with every row as
 * wide as the widest, lines ended with CR LF. Windows-1252 that can't hold the text switches to UTF-8 with a BOM.
 */
export function saveText(wb: Workbook, sheetIndex: number, o: TextSave): Uint8Array {
  const sheet = wb.sheets[sheetIndex]
  const used = sheet.cells.usedBounds()
  const rows: string[][] = []
  if (used) {
    const width = used.c1 + 1
    for (let r = 0; r <= used.r1; r++) rows.push(new Array<string>(width).fill(""))
    for (const { row, col, data } of sheet.cells.entries({ r0: 0, c0: 0, r1: used.r1, c1: used.c1 })) {
      if (data.value === undefined || data.value === "") continue
      const numberFormat = data.style ? wb.styles.get(data.style).numberFormat : undefined
      let text = shownText(data.value, numberFormat)
      if (hasFormatMarks(text)) text = [...text].map((c) => markText(c)).join("")
      rows[row][col] = text
    }
  }
  const body = joinDelimited(rows, o.delimiter)
  if (o.encoding === "windows-1252" && !fits1252(body)) return encodeText(body, "utf-8", true)
  return encodeText(body, o.encoding, o.bom)
}

/** The text a cell shows, as the grid draws it (General as in a default-width column). */
function shownText(value: CellValue, code: string | undefined): string {
  const formatted = code && code !== "General"
  if (typeof value === "string") return formatted && code.includes("@") ? formatWithCode(value, code).text : value
  if (typeof value === "number" && formatted) return formatWithCode(value, code).text
  return formatValue(value) ?? String(value)
}

/** A format's padding (`_)`) is a space in text; a repeat fill (`*-`) is left out. */
function markText(c: string): string {
  const mark = formatMark(c)
  if (!mark) return c
  return mark.kind === "pad" ? " " : ""
}
