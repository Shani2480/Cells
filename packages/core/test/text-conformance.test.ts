import { describe, expect, it } from "vitest"
import { parseInput, Workbook } from "../src"
import fixture from "./fixtures/excel-text.json"

type Case = { code: string; value: number | string; text: string }

/** TEXT(value, code) as Cells calculates it. */
function text(wb: Workbook, value: number | string, code: string): unknown {
  const arg = typeof value === "number" ? String(value) : `"${value}"`
  const parsed = parseInput(`=TEXT(${arg},"${code.replace(/"/g, '""')}")`)!
  wb.transact([{ kind: "setCells", sheet: 0, cells: [{ row: 0, col: 0, data: parsed }] }], { label: "t" })
  return wb.sheets[0].cells.value(0, 0)
}

describe("number formats match Excel's TEXT", () => {
  const wb = new Workbook(["Sheet1"])
  const cases = (fixture as { cases: Case[] }).cases
  it.each(cases.map((c) => [c.code, c.value, c.text] as const))("TEXT(%s) of %s", (code, value, expected) => {
    expect(text(wb, value, code)).toEqual(expected === "#ERR" ? { error: "#VALUE!" } : expected)
  })
})
