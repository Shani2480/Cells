import { readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { parseCell, parseInput, Workbook, type CellWrite } from "../src"

/**
 * Replays the formulas Excel evaluated (tests/excel-golden, made by the local tools/excel-golden.py) and compares
 * our results with Excel's. Numbers must agree to Excel's 15 significant digits.
 */

type Expected = number | string | boolean | null | { error: string }
type Case = { formula: string; value?: Expected; array?: Expected[][]; rejected?: string }

const DIR = join(__dirname, "../../../tests/excel-golden")
const cases = JSON.parse(readFileSync(join(DIR, "cases.json"), "utf8")) as {
  cell: string
  fixture: Record<string, string>
}
/** Cases where we knowingly differ from Excel, with the reason. */
const known = JSON.parse(readFileSync(join(DIR, "known-differences.json"), "utf8")) as Record<string, string>
/** Functions whose last digit may differ from Excel's, with how far (tests/excel-golden/tolerances.json says why). */
const tolerances = JSON.parse(readFileSync(join(DIR, "tolerances.json"), "utf8")) as {
  families: { relative: number; functions: string[] }[]
}
function toleranceOf(formula: string): number {
  const name = /^=([A-Z][A-Z0-9.]*)\(/.exec(formula)?.[1]
  return tolerances.families.find((f) => name && f.functions.includes(name))?.relative ?? 0
}

function fixtureBook() {
  const wb = new Workbook(["Sheet1"])
  const cells: CellWrite[] = Object.entries(cases.fixture).map(([ref, text]) => {
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)!
    return { row, col, data: { value: parsed.value, input: parsed.input } }
  })
  wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Fixture" })
  return wb
}

function actual(wb: Workbook, row: number, col: number): Expected {
  const v = wb.sheets[0].cells.value(row, col)
  if (v === undefined) return null
  if (typeof v === "object" && v !== null && "error" in v) return { error: v.error }
  return v as Expected
}

function same(a: Expected, b: Expected, tolerance = 0): boolean {
  if (typeof a === "number" && typeof b === "number") {
    if (a === b) return true
    if (Number(a.toPrecision(15)) === Number(b.toPrecision(15))) return true
    return tolerance > 0 && Math.abs(a - b) <= tolerance * Math.max(Math.abs(a), Math.abs(b))
  }
  if (a && b && typeof a === "object" && typeof b === "object") return a.error === b.error
  return a === b
}

const show = (v: Expected) => (v && typeof v === "object" ? v.error : JSON.stringify(v))

const target = parseCell(cases.cell)!
const wb = fixtureBook()

/** Enters `formula` at the case cell, returns what it shows (all of its spill for arrays), then clears it. */
function evaluate(c: Case): Expected | Expected[][] {
  const parsed = parseInput(c.formula)!
  wb.transact(
    [{ kind: "setCells", sheet: 0, cells: [{ ...target, data: { value: parsed.value, input: parsed.input } }] }],
    { label: "Case" },
  )
  let result: Expected | Expected[][]
  if (c.array) {
    result = c.array.map((row, r) => row.map((_, k) => actual(wb, target.row + r, target.col + k)))
  } else {
    result = actual(wb, target.row, target.col)
  }
  wb.transact([{ kind: "setCells", sheet: 0, cells: [{ ...target, data: undefined }] }], { label: "Clear" })
  return result
}

const groups = readdirSync(DIR).filter(
  (f) => f.endsWith(".json") && !["cases.json", "known-differences.json", "tolerances.json"].includes(f),
)

describe.each(groups)("Excel conformance: %s", (file) => {
  const golden = JSON.parse(readFileSync(join(DIR, file), "utf8")) as { cases: Case[] }
  const runnable = golden.cases.filter((c) => !c.rejected)
  it.each(runnable.map((c) => [c.formula, c] as const))("%s", (_, c) => {
    const got = evaluate(c)
    let ok: boolean
    let message: string
    if (c.array) {
      const rows = got as Expected[][]
      const tolerance = toleranceOf(c.formula)
      ok = c.array.every((row, r) => row.every((v, k) => same(v === null ? 0 : v, rows[r][k] ?? 0, tolerance)))
      message = `Excel ${JSON.stringify(c.array)}, ours ${JSON.stringify(rows)}`
    } else {
      ok = same(c.value!, got as Expected, toleranceOf(c.formula))
      message = `Excel ${show(c.value!)}, ours ${show(got as Expected)}`
    }
    if (known[c.formula]) {
      // Documented difference: fail loudly once it is fixed so the list stays honest.
      expect(ok, `${c.formula} now matches Excel; remove it from known-differences.json`).toBe(false)
      return
    }
    expect(ok, message).toBe(true)
  })
})
