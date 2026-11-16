import { describe, expect, it } from "vitest"
import { nextPart, parseFormula, replacePart, stepText, valueNode, Workbook, type Node } from "../src"

function steps(wb: Workbook, row: number, col: number): string[] {
  const parsed = parseFormula(wb.sheets[0].cells.input(row, col)!.slice(1))
  if (!parsed.ok) throw new Error("parse")
  let root: Node = parsed.ast
  const out: string[] = []
  for (let guard = 0; guard < 20; guard++) {
    const part = nextPart(root)
    const { text, start, end } = stepText(root, part)
    out.push(part ? `${text.slice(0, start)}[${text.slice(start, end)}]${text.slice(end)}` : text)
    if (!part) break
    root = replacePart(root, part, valueNode(wb.calc.evaluatePart(part, 0, row, col)!))
  }
  return out
}

describe("Evaluate Formula's steps", () => {
  it("works from the innermost, leftmost part out, keeping ranges for the function that uses them", () => {
    const wb = new Workbook(["Sheet1"])
    const s = wb.sheets[0]
    s.cells.set(0, 0, { value: 2 })
    s.cells.set(1, 0, { value: 3 })
    s.cells.set(0, 1, { value: 0, input: "=SUM(A1:A2)*A1+1" })
    s.cells.set(1, 1, { value: 0, input: '=IF(A1>1,"big",A2/0)' })
    wb.calc.rebuild()
    wb.calculate(true)
    expect(steps(wb, 0, 1)).toEqual(["[SUM(A1:A2)]*A1+1", "5*[A1]+1", "[5*2]+1", "[10+1]", "11"])
    expect(steps(wb, 1, 1)).toEqual(['IF([A1]>1,"big",A2/0)', 'IF([2>1],"big",A2/0)', '[IF(TRUE,"big",A2/0)]', '"big"'])
  })
})
