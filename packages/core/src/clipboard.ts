/**
 * Clipboard text formats as Excel writes and reads them: tab-separated rows (fields with tabs, line breaks or quotes
 * are quoted, quotes doubled; every row ends with CRLF) and a plain HTML table.
 */

export function encodeTSV(rows: readonly (readonly string[])[]): string {
  return rows.map((row) => row.map(quoteField).join("\t") + "\r\n").join("")
}

function quoteField(field: string) {
  return /[\t\r\n"]/.test(field) ? `"${field.replace(/"/g, '""')}"` : field
}

/**
 * Parses tab-separated text into rows. A field that starts with a quote runs to its closing quote (and may hold
 * tabs and line breaks); a final line break does not make an empty row.
 */
export function decodeTSV(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ""
  let i = 0
  let fieldStart = true
  while (i < text.length) {
    const ch = text[i]
    if (fieldStart && ch === '"') {
      // A quoted field only if a closing quote is followed by a separator or the end; else it is literal text.
      const end = findClosingQuote(text, i + 1)
      if (end !== -1) {
        field = text.slice(i + 1, end).replace(/""/g, '"')
        i = end + 1
        fieldStart = false
        continue
      }
    }
    fieldStart = false
    if (ch === "\t") {
      row.push(field)
      field = ""
      fieldStart = true
      i++
    } else if (ch === "\r" || ch === "\n") {
      row.push(field)
      rows.push(row)
      row = []
      field = ""
      fieldStart = true
      i += ch === "\r" && text[i + 1] === "\n" ? 2 : 1
    } else {
      field += ch
      i++
    }
  }
  if (!fieldStart || field !== "" || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  return rows
}

function findClosingQuote(text: string, from: number): number {
  for (let i = from; i < text.length; i++) {
    if (text[i] === '"') {
      if (text[i + 1] === '"') {
        i++
        continue
      }
      const next = text[i + 1]
      return next === undefined || next === "\t" || next === "\r" || next === "\n" ? i : -1
    }
  }
  return -1
}

function escapeHtml(text: string) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
}

/** A minimal HTML table other apps (and Inkspire) paste as a table. Line breaks in a cell become <br>. */
export function encodeHTML(rows: readonly (readonly string[])[]): string {
  const body = rows
    .map(
      (row) =>
        "<tr>" + row.map((field) => `<td>${escapeHtml(field).replace(/\r?\n/g, "<br>")}</td>`).join("") + "</tr>",
    )
    .join("")
  return `<meta charset="utf-8"><table><tbody>${body}</tbody></table>`
}
