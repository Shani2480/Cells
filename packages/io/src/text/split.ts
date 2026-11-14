/**
 * Splitting text into fields, as Excel's Text Import Wizard does, and joining them again as Excel saves CSV and
 * tab-delimited text.
 */
export type DelimitedOptions = {
  /** One-character delimiters; any of them ends a field. */
  delimiters: readonly string[]
  /** The text qualifier ('"', "'"), or null for none. */
  qualifier: string | null
  /** "Treat consecutive delimiters as one". */
  consecutive: boolean
}

export const CSV: DelimitedOptions = { delimiters: [","], qualifier: '"', consecutive: false }

/**
 * Rows of fields. Lines end at CR LF, LF or CR outside a qualifier; inside one, a line break stays in the field (as
 * LF) and a doubled qualifier is one. Text after a closing qualifier joins the field, as in Excel ("ab"cd → abcd).
 * A last line break doesn't start a row. Stops after `limit` rows.
 */
export function parseDelimited(text: string, o: DelimitedOptions, limit = Infinity): string[][] {
  const rows: string[][] = []
  const delims = new Set(o.delimiters)
  const q = o.qualifier
  const n = text.length
  let i = 0
  let row: string[] = []
  // Whether the field about to start follows a delimiter (for consecutive delimiters).
  let afterDelimiter = false
  while (i < n && rows.length < limit) {
    let field = ""
    let quoted = false
    if (q && text[i] === q) {
      quoted = true
      i++
      let start = i
      for (;;) {
        const at = text.indexOf(q, i)
        if (at < 0) {
          field += text.slice(start, n)
          i = n
          break
        }
        field += text.slice(start, at)
        if (text[at + 1] === q) {
          field += q
          i = at + 2
          start = i
          continue
        }
        i = at + 1
        break
      }
      if (field.includes("\r")) field = field.replace(/\r\n?/g, "\n")
    }
    // The rest of the field, up to a delimiter or a line end.
    const start = i
    while (i < n) {
      const c = text[i]
      if (delims.has(c) || c === "\n" || c === "\r") break
      i++
    }
    field += text.slice(start, i)
    const c = i < n ? text[i] : ""
    if (delims.has(c)) {
      if (!(o.consecutive && afterDelimiter && field === "" && !quoted)) row.push(field)
      afterDelimiter = true
      i++
      continue
    }
    if (!(o.consecutive && afterDelimiter && field === "" && !quoted && row.length)) row.push(field)
    rows.push(row)
    row = []
    afterDelimiter = false
    if (c === "\r" && text[i + 1] === "\n") i += 2
    else if (c) i++
  }
  // The text ended right after a delimiter: the row's last field is empty.
  if (afterDelimiter && rows.length < limit) {
    if (!o.consecutive) row.push("")
    rows.push(row)
  }
  return rows
}

/** The lines of a text, without their line breaks (a last line break doesn't start a line). */
export function textLines(text: string, limit = Infinity): string[] {
  const lines: string[] = []
  let i = 0
  while (i < text.length && lines.length < limit) {
    let end = i
    while (end < text.length && text[end] !== "\n" && text[end] !== "\r") end++
    lines.push(text.slice(i, end))
    i = text[end] === "\r" && text[end + 1] === "\n" ? end + 2 : end + 1
  }
  return lines
}

/** Fixed-width fields: each line cut at the character positions in `breaks` (ascending), each field trimmed. */
export function splitFixed(lines: readonly string[], breaks: readonly number[]): string[][] {
  return lines.map((line) => {
    const fields: string[] = []
    let from = 0
    for (const at of [...breaks, Infinity]) {
      fields.push(line.slice(from, at === Infinity ? undefined : at).trim())
      from = at
    }
    return fields
  })
}

const CANDIDATES = [",", ";", "\t", "|"]

/**
 * The delimiter a text most likely uses: the one that splits the first lines into the same number of fields (more
 * than one) most often. Ties go to the earlier candidate (comma first).
 */
export function detectDelimiter(text: string, fallback = ","): string {
  let best = fallback
  let bestScore = 0
  for (const d of CANDIDATES) {
    const rows = parseDelimited(text, { delimiters: [d], qualifier: '"', consecutive: false }, 30).filter(
      (r) => r.length > 1 || r[0] !== "",
    )
    if (!rows.length) continue
    const counts = new Map<number, number>()
    for (const r of rows) counts.set(r.length, (counts.get(r.length) ?? 0) + 1)
    let mode = 1
    let modeRows = 0
    for (const [fields, times] of counts)
      if (times > modeRows || (times === modeRows && fields > mode)) [mode, modeRows] = [fields, times]
    if (mode < 2) continue
    // Consistent rows count most; more fields break ties between equally consistent delimiters.
    const score = (modeRows / rows.length) * 1000 + Math.min(mode, 100)
    if (score > bestScore) [best, bestScore] = [d, score]
  }
  return best
}

/**
 * Where a fixed-width text's columns likely start: after a run of spaces that every sampled line has (or is too
 * short to reach), where some line has text again.
 */
export function detectBreaks(lines: readonly string[]): number[] {
  const sample = lines.filter((l) => l.trim() !== "").slice(0, 50)
  if (!sample.length) return []
  const width = Math.max(...sample.map((l) => l.length))
  const blank = (col: number) => sample.every((l) => col >= l.length || l[col] === " ")
  const breaks: number[] = []
  for (let col = 1; col < width; col++) if (blank(col - 1) && !blank(col)) breaks.push(col)
  return breaks
}

/** A field written as Excel writes it: quoted when it holds a comma, quote, line break, tab or the delimiter. */
export function quoteField(text: string, delimiter: string): string {
  if (/[",\r\n\t]/.test(text) || (delimiter && text.includes(delimiter))) return `"${text.replace(/"/g, '""')}"`
  return text
}

/** Rows joined with `delimiter`, every line ended with CR LF (Excel's own line ends). */
export function joinDelimited(rows: readonly (readonly string[])[], delimiter: string): string {
  let out = ""
  for (const row of rows) out += row.map((f) => quoteField(f, delimiter)).join(delimiter) + "\r\n"
  return out
}
