import { parseCell, parseInput, Workbook, type CellWrite } from "@cells/core"
import { describe, expect, it } from "vitest"
import { createEngine } from "./engine"
import { inProcessPort, WorkerCalculator } from "./host"
import type { FromWorker, Port, ToWorker } from "./protocol"

function type(wb: Workbook, cells: Record<string, string>) {
  const writes: CellWrite[] = Object.entries(cells).map(([ref, text]) => {
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)
    return { row, col, data: parsed ? { value: parsed.value, input: parsed.input } : undefined }
  })
  wb.transact([{ kind: "setCells", sheet: 0, cells: writes }], { label: "Typing" })
}

const value = (wb: Workbook, ref: string) => {
  const { row, col } = parseCell(ref)!
  return wb.sheets[0].cells.value(row, col)
}

/** A port whose answers wait until `flush`, like a real worker's. */
function queuedPort() {
  const queue: FromWorker[] = []
  let handler: ((m: FromWorker) => void) | null = null
  const engine = createEngine((m) => queue.push(m))
  const port: Port = {
    post: (m: ToWorker) => engine(structuredClone(m)),
    listen: (h) => {
      handler = h
    },
  }
  return { port, flush: () => queue.splice(0).forEach((m) => handler?.(m)) }
}

describe("calculation in a worker", () => {
  it("values arrive from the engine; the UI workbook does not calculate itself", () => {
    const wb = new Workbook(["Sheet1"])
    const { port, flush } = queuedPort()
    const calculator = new WorkerCalculator(wb, port)
    type(wb, { A1: "2", A2: "=A1*10" })
    expect(value(wb, "A2")).toBe("") // typed, not calculated yet
    expect(calculator.outstanding).toBe(2)
    flush()
    expect(value(wb, "A2")).toBe(20)
    expect(calculator.outstanding).toBe(0)
    expect(wb.calc.formulaCount).toBe(0)
  })

  it("starts from what the workbook already holds", () => {
    const wb = new Workbook(["Sheet1"])
    type(wb, { A1: "5", B1: "=A1+1" }) // calculated locally before the worker takes over
    type(wb, { A1: "6" })
    const { port, flush } = queuedPort()
    new WorkerCalculator(wb, port)
    type(wb, { A1: "7" })
    flush()
    expect(value(wb, "B1")).toBe(8)
  })

  it("follows undo, structural edits and spills", () => {
    const wb = new Workbook(["Sheet1"])
    new WorkerCalculator(wb, inProcessPort())
    type(wb, { A1: "1", A2: "2", A3: "3", B1: "=SUM(A1:A3)", C1: "=SEQUENCE(3)" })
    expect(value(wb, "B1")).toBe(6)
    expect(value(wb, "C3")).toBe(3)
    expect(wb.calcStatus.spills).toEqual([{ sheet: 0, row: 0, col: 2, range: { r0: 0, c0: 2, r1: 2, c1: 2 } }])
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 1, count: 1 }], { label: "Delete" })
    expect(wb.sheets[0].cells.input(0, 1)).toBe("=SUM(A1:A2)")
    expect(value(wb, "B1")).toBe(4)
    expect(value(wb, "C3")).toBe(3) // the spill fills its rows again
    wb.undo()
    expect(value(wb, "B1")).toBe(6)
    type(wb, { C2: "x" })
    expect(value(wb, "C1")).toEqual({ error: "#SPILL!" })
    expect(value(wb, "C3")).toBeUndefined()
    wb.undo()
    expect(value(wb, "C2")).toBe(2)
  })

  it("reports circular references and manual mode", () => {
    const wb = new Workbook(["Sheet1"])
    new WorkerCalculator(wb, inProcessPort())
    type(wb, { A1: "=B1", B1: "=A1" })
    expect(wb.calcStatus.circular).not.toBeNull()
    type(wb, { B1: "1" })
    expect(wb.calcStatus.circular).toBeNull()
    wb.setCalcMode("manual")
    type(wb, { B1: "5" })
    expect(value(wb, "A1")).toBe(1)
    expect(wb.calcStatus.pending).toBe(true)
    wb.calculate()
    expect(value(wb, "A1")).toBe(5)
    expect(wb.calcStatus.pending).toBe(false)
  })
})
