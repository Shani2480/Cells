import { currentRegion, parseFormula, rangeName, type Range } from "@cells/core"
import type { SheetController } from "@/sheet/controller"

/**
 * The AI tools' requests and how their replies are read (docs/ai.md). Pure: the controller is only read. Each tool
 * sends the cells it is about (never more than SAMPLE_ROWS rows), and every reply is checked before anything is put
 * into the workbook: a formula must parse, a chart must be one Cells draws.
 */

export type AiTool = "formula" | "explain" | "chart" | "summary"

/** The most rows of data a request carries. */
export const SAMPLE_ROWS = 60
const SAMPLE_COLS = 26

/** Charts the AI may suggest: the common ones, by their Insert Chart id. */
export const SUGGESTED_CHARTS: Record<string, string> = {
  column: "Clustered Column",
  "column-stacked": "Stacked Column",
  hbar: "Clustered Bar",
  "line-markers": "Line with Markers",
  pie: "Pie",
  doughnut: "Doughnut",
  area: "Area",
  scatter: "Scatter",
  "combo-column-line-secondary": "Clustered Column - Line on Secondary Axis",
}

/** Cells as "A1: Item | B1: Cost" lines, row by row: what the model reads (formats shown, as the person sees them). */
export function describeCells(c: SheetController, range: Range): string {
  const sheet = c.sheet
  const r1 = Math.min(range.r1, range.r0 + SAMPLE_ROWS - 1)
  const c1 = Math.min(range.c1, range.c0 + SAMPLE_COLS - 1)
  const block = c.rangeContent(c.sheetIndex, { r0: range.r0, c0: range.c0, r1, c1 })
  const rows = block.text.replace(/\r\n$/, "").split("\r\n")
  const lines: string[] = []
  for (let r = range.r0; r <= r1; r++) {
    const texts = rows[r - range.r0]?.split("\t") ?? []
    const cells: string[] = []
    for (let col = range.c0; col <= c1; col++) {
      const shown = texts[col - range.c0] ?? ""
      const input = sheet.cells.get(r, col)?.input
      if (!shown && !input) continue
      const name = rangeName({ r0: r, c0: col, r1: r, c1: col })
      cells.push(input?.startsWith("=") ? `${name}: ${input} (shows ${shown})` : `${name}: ${shown}`)
    }
    if (cells.length) lines.push(cells.join(" | "))
  }
  const more = range.r1 > r1 ? `\n(${range.r1 - r1} more rows not shown)` : ""
  return lines.join("\n") + more
}

/** The data a tool is about: the selection, or the table of data around the active cell when one cell is selected. */
export function dataRange(c: SheetController): Range {
  const sel = c.selection.ranges[c.selection.ranges.length - 1]
  const single = sel.r0 === sel.r1 && sel.c0 === sel.c1
  return single ? currentRegion(c.sheet, c.selection.active) : sel
}

const RULES = `You work inside Cells, a spreadsheet that follows Microsoft Excel. Formulas use Excel's English function
names and syntax: commas between arguments, "." for decimals, A1 references. Answer in the language the person writes
in.`

export type Request = { system: string; prompt: string; context: string }

/** Formula from Words: a formula for the active cell. The reply is JSON: {"formula": "=…", "explanation": "…"}. */
export function formulaRequest(c: SheetController, description: string): Request {
  const at = c.selection.active
  const cell = rangeName({ r0: at.row, c0: at.col, r1: at.row, c1: at.col })
  const region = currentRegion(c.sheet, at)
  const near = {
    r0: Math.max(0, region.r0),
    c0: region.c0,
    r1: Math.max(region.r1, at.row),
    c1: Math.max(region.c1, at.col),
  }
  return {
    system: `${RULES}
Write one Excel formula for the cell the person names. Reply with JSON only, no code fence:
{"formula": "=...", "explanation": "one or two plain sentences on what it does"}
Use only functions Excel has. Refer to the cells by address. If the request can't be done with a formula, reply
{"formula": "", "explanation": "why not"}.`,
    prompt: `In cell ${cell} of sheet "${c.sheet.name}": ${description.trim()}`,
    context: `Cells near ${cell}:\n${describeCells(c, near)}`,
  }
}

/** Explain Formula: the active cell's formula in plain words. */
export function explainRequest(c: SheetController): Request | null {
  const at = c.selection.active
  const input = c.sheet.cells.get(at.row, at.col)?.input
  if (!input?.startsWith("=")) return null
  const cell = rangeName({ r0: at.row, c0: at.col, r1: at.row, c1: at.col })
  return {
    system: `${RULES}
Explain the formula to someone who uses spreadsheets but doesn't know every function. Plain text, no Markdown
headings: first one sentence on what it works out, then a short list ("- ") of its parts in order, then any problem
you see (a range that misses rows, an error it may give). Keep it under 150 words.`,
    prompt: `Explain the formula in ${cell} of sheet "${c.sheet.name}": ${input}`,
    context: `Cells around it:\n${describeCells(c, currentRegion(c.sheet, at))}`,
  }
}

/** Suggest a Chart: one chart type for the data. The reply is JSON: {"chart": id, "title": "…", "why": "…"}. */
export function chartRequest(c: SheetController): Request {
  const range = dataRange(c)
  return {
    system: `${RULES}
Choose the one chart that shows this data best. Reply with JSON only, no code fence:
{"chart": "<one of: ${Object.keys(SUGGESTED_CHARTS).join(", ")}>", "title": "a short chart title", "why": "one sentence"}`,
    prompt: `Suggest a chart for the data in ${rangeName(range)} of sheet "${c.sheet.name}".`,
    context: describeCells(c, range),
  }
}

/** Summarize Data: what the data says, in a few lines. */
export function summaryRequest(c: SheetController): Request {
  const range = dataRange(c)
  return {
    system: `${RULES}
Summarize what the data shows for someone about to present it: three to five short lines starting with "- ", each one
fact with its numbers (totals, highest and lowest, changes, anything unusual). Plain text, no headings. Don't invent
values that aren't in the data; say so if the data is too small to summarize.`,
    prompt: `Summarize the data in ${rangeName(range)} of sheet "${c.sheet.name}".`,
    context: describeCells(c, range),
  }
}

/** The first JSON object in a reply (models sometimes add words around it); null when there is none. */
function jsonIn(reply: string): Record<string, unknown> | null {
  const start = reply.indexOf("{")
  const end = reply.lastIndexOf("}")
  if (start < 0 || end <= start) return null
  try {
    const value = JSON.parse(reply.slice(start, end + 1)) as unknown
    return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null
  } catch {
    return null
  }
}

const str = (v: unknown) => (typeof v === "string" ? v.trim() : "")

export type FormulaReply = { formula: string; explanation: string } | { problem: string }

/** The formula the model wrote, if it is one Cells can read. */
export function readFormulaReply(reply: string): FormulaReply {
  const json = jsonIn(reply)
  if (!json) return { problem: "The AI didn't answer with a formula. Try describing it another way." }
  let formula = str(json.formula)
  const explanation = str(json.explanation)
  if (!formula) return { problem: explanation || "The AI couldn't write a formula for that." }
  if (!formula.startsWith("=")) formula = `=${formula}`
  const parsed = parseFormula(formula.slice(1))
  if (!parsed.ok)
    return { problem: `The AI wrote ${formula}, which isn't a formula Cells can read (${parsed.message}).` }
  return { formula, explanation }
}

export type ChartReply = { chart: string; label: string; title: string; why: string } | { problem: string }

export function readChartReply(reply: string): ChartReply {
  const json = jsonIn(reply)
  const chart = str(json?.chart)
  const label = SUGGESTED_CHARTS[chart]
  if (!json || !label) return { problem: "The AI didn't suggest a chart Cells can draw. Try again." }
  return { chart, label, title: str(json.title), why: str(json.why) }
}
