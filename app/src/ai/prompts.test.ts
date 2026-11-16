import { describe, expect, it } from "vitest"
import { SheetController } from "@/sheet/controller"
import { aiTarget } from "./client"
import {
  chartRequest,
  describeCells,
  explainRequest,
  formulaRequest,
  readChartReply,
  readFormulaReply,
  summaryRequest,
} from "./prompts"

function budget() {
  const c = new SheetController()
  const s = c.wb.sheets[0]
  s.cells.set(0, 0, { value: "Item" })
  s.cells.set(0, 1, { value: "Cost" })
  s.cells.set(1, 0, { value: "Rent" })
  s.cells.set(1, 1, { value: 1200 })
  s.cells.set(2, 0, { value: "Water" })
  s.cells.set(2, 1, { value: 45 })
  s.cells.set(3, 1, { value: 1245, input: "=SUM(B2:B3)" })
  return c
}

describe("AI requests", () => {
  it("describes cells by address, with formulas and what they show", () => {
    const c = budget()
    expect(describeCells(c, { r0: 0, c0: 0, r1: 3, c1: 1 })).toBe(
      "A1: Item | B1: Cost\nA2: Rent | B2: 1200\nA3: Water | B3: 45\nB4: =SUM(B2:B3) (shows 1245)",
    )
  })

  it("asks for a formula for the active cell, with the cells near it", () => {
    const c = budget()
    c.goTo({ r0: 4, c0: 1, r1: 4, c1: 1 })
    const r = formulaRequest(c, "the average cost")
    expect(r.prompt).toBe('In cell B5 of sheet "Sheet1": the average cost')
    expect(r.context).toContain("B3: 45")
    expect(r.system).toContain('{"formula"')
  })

  it("explains only a formula; charts and summaries take the data around the active cell", () => {
    const c = budget()
    c.goTo({ r0: 1, c0: 1, r1: 1, c1: 1 })
    expect(explainRequest(c)).toBeNull()
    c.goTo({ r0: 3, c0: 1, r1: 3, c1: 1 })
    expect(explainRequest(c)!.prompt).toContain("=SUM(B2:B3)")
    expect(chartRequest(c).prompt).toContain("A1:B4")
    expect(summaryRequest(c).context).toContain("A2: Rent")
  })
})

describe("AI replies", () => {
  it("takes a formula only when Cells can read it", () => {
    expect(readFormulaReply('Sure! {"formula": "=AVERAGE(B2:B3)", "explanation": "The mean cost."}')).toEqual({
      formula: "=AVERAGE(B2:B3)",
      explanation: "The mean cost.",
    })
    expect(readFormulaReply('{"formula": "SUM(B2:B3)", "explanation": ""}')).toMatchObject({ formula: "=SUM(B2:B3)" })
    expect(readFormulaReply('{"formula": "=SUM(B2:", "explanation": ""}')).toHaveProperty("problem")
    expect(readFormulaReply('{"formula": "", "explanation": "Needs a date column."}')).toEqual({
      problem: "Needs a date column.",
    })
    expect(readFormulaReply("I can't")).toHaveProperty("problem")
  })

  it("takes a chart only from the ones it may suggest", () => {
    expect(readChartReply('{"chart": "pie", "title": "Costs", "why": "Parts of a whole."}')).toEqual({
      chart: "pie",
      label: "Pie",
      title: "Costs",
      why: "Parts of a whole.",
    })
    expect(readChartReply('{"chart": "surface-3d"}')).toHaveProperty("problem")
  })
})

describe("where requests go", () => {
  const base = { aiModels: {}, aiAddresses: {} }
  it("follows the shared provider choice", () => {
    expect(aiTarget({ ...base, aiProvider: "ollama" })).toEqual({
      provider: "ollama",
      model: undefined,
      baseUrl: "http://127.0.0.1:11434",
    })
    expect(aiTarget({ ...base, aiProvider: "lmstudio", aiModels: { lmstudio: "qwen3-8b" } })).toEqual({
      provider: "compatible",
      model: "qwen3-8b",
      baseUrl: "http://127.0.0.1:1234/v1",
    })
    expect(aiTarget({ ...base, aiProvider: "anthropic" })).toEqual({ provider: "anthropic", model: undefined })
    expect(aiTarget({ ...base, aiProvider: "" })).toHaveProperty("problem")
    expect(aiTarget({ ...base, aiProvider: "builtin" })).toHaveProperty("problem")
  })
})
