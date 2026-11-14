import { describe, expect, it } from "vitest"
import { parseCell, parseInput, Workbook, type CellWrite } from "../src"

/** A workbook with cells typed as a person would type them. */
function sheet(cells: Record<string, string>, sheets = ["Sheet1"]) {
  const wb = new Workbook(sheets)
  set(wb, cells)
  return wb
}

function set(wb: Workbook, cells: Record<string, string>, sheetIndex = 0) {
  const writes: CellWrite[] = Object.entries(cells).map(([ref, text]) => {
    const { row, col } = parseCell(ref)!
    const parsed = parseInput(text)
    return { row, col, data: parsed ? { value: parsed.value, input: parsed.input } : undefined }
  })
  wb.transact([{ kind: "setCells", sheet: sheetIndex, cells: writes }], { label: "Typing" })
}

function get(wb: Workbook, ref: string, sheetIndex = 0) {
  const { row, col } = parseCell(ref)!
  const v = wb.sheets[sheetIndex].cells.value(row, col)
  return typeof v === "object" && v !== null && "error" in v ? v.error : v
}

describe("calculation", () => {
  it("calculates on entry and follows changes", () => {
    const wb = sheet({ A1: "2", A2: "3", A3: "=A1*A2", A4: "=A3+1" })
    expect(get(wb, "A3")).toBe(6)
    expect(get(wb, "A4")).toBe(7)
    set(wb, { A1: "10" })
    expect(get(wb, "A3")).toBe(30)
    expect(get(wb, "A4")).toBe(31)
  })

  it("follows ranges, including whole columns", () => {
    const wb = sheet({ A1: "1", A2: "2", B1: "=SUM(A:A)", B2: "=SUM(A1:A3)" })
    expect(get(wb, "B1")).toBe(3)
    set(wb, { A3: "4" })
    expect(get(wb, "B1")).toBe(7)
    expect(get(wb, "B2")).toBe(7)
    set(wb, { A500000: "100" })
    expect(get(wb, "B1")).toBe(107)
    expect(get(wb, "B2")).toBe(7)
  })

  it("orders formulas entered out of order", () => {
    const wb = sheet({ A1: "=A2+1", A2: "=A3+1", A3: "=5" })
    expect(get(wb, "A1")).toBe(7)
  })

  it("a formula pointing at a blank cell is 0; at text it is the text", () => {
    const wb = sheet({ A1: "=B1", A2: "=B2", B2: "hi" })
    expect(get(wb, "A1")).toBe(0)
    expect(get(wb, "A2")).toBe("hi")
  })

  it("undo restores the formula and recalculates", () => {
    const wb = sheet({ A1: "1", B1: "=A1*10" })
    set(wb, { B1: "=A1*20" })
    expect(get(wb, "B1")).toBe(20)
    wb.undo()
    expect(get(wb, "B1")).toBe(10)
    set(wb, { A1: "3" })
    expect(get(wb, "B1")).toBe(30)
  })

  it("replacing a formula with a constant stops it from recalculating", () => {
    const wb = sheet({ A1: "1", B1: "=A1" })
    set(wb, { B1: "5" })
    set(wb, { A1: "9" })
    expect(get(wb, "B1")).toBe(5)
  })

  it("detects circular references", () => {
    const wb = sheet({ A1: "=B1+1", B1: "=A1+1" })
    expect(wb.calc.circular).not.toBeNull()
    const wb2 = sheet({ A1: "=A1+1" })
    expect(get(wb2, "A1")).toBe(0)
    expect(wb2.calc.circular).toEqual({ sheet: 0, row: 0, col: 0 })
  })

  it("recalculates volatile functions on every change", () => {
    const wb = sheet({ A1: "=RAND()" })
    let n = 0
    wb.calc.random = () => (n++ % 10) / 10
    set(wb, { B1: "x" })
    const first = get(wb, "A1")
    set(wb, { B1: "y" })
    expect(get(wb, "A1")).not.toBe(first)
  })

  it("reads other sheets and 3-D references", () => {
    const wb = new Workbook(["Jan", "Feb", "Total"])
    set(wb, { A1: "5" }, 0)
    set(wb, { A1: "7" }, 1)
    set(wb, { A1: "=Jan!A1+Feb!A1", A2: "=SUM(Jan:Feb!A1)", A3: "=Nope!A1" }, 2)
    expect(get(wb, "A1", 2)).toBe(12)
    expect(get(wb, "A2", 2)).toBe(12)
    expect(get(wb, "A3", 2)).toBe("#REF!")
    set(wb, { A1: "1" }, 1)
    expect(get(wb, "A2", 2)).toBe(6)
  })

  it("unknown functions are #NAME? and keep their text", () => {
    const wb = sheet({ A1: "=FOO(1)" })
    expect(get(wb, "A1")).toBe("#NAME?")
    expect(wb.sheets[0].cells.input(0, 0)).toBe("=FOO(1)")
  })

  it("Excel's last-operation clean-up", () => {
    const wb = sheet({ A1: "=0.1+0.2-0.3", A2: "=(0.1+0.2-0.3)", A3: "=SUM(0.1,0.2,-0.3)", A4: "=SUM(0.1,0.2,-0.3,0)" })
    expect(get(wb, "A1")).toBe(0)
    expect(get(wb, "A2")).toBe(5.551115123125783e-17)
    expect(get(wb, "A3")).toBe(0)
    expect(get(wb, "A4")).not.toBe(0)
  })
})

describe("dynamic arrays", () => {
  it("spills, and spilled cells are read by other formulas", () => {
    const wb = sheet({ A1: "=SEQUENCE(3)", B1: "=A3*10", C1: "=SUM(A1#)" })
    expect([get(wb, "A1"), get(wb, "A2"), get(wb, "A3")]).toEqual([1, 2, 3])
    expect(get(wb, "B1")).toBe(30)
    expect(get(wb, "C1")).toBe(6)
    expect(wb.sheets[0].cells.isSpill(1, 0)).toBe(true)
    expect(wb.calc.spillAt(0, 2, 0)).toEqual({ r0: 0, c0: 0, r1: 2, c1: 0 })
  })

  it("a value in the way gives #SPILL!, and removing it lets the array spill again", () => {
    const wb = sheet({ A1: "=SEQUENCE(3)" })
    set(wb, { A2: "x" })
    expect(get(wb, "A1")).toBe("#SPILL!")
    expect(get(wb, "A3")).toBeUndefined()
    set(wb, { A2: "" })
    expect(get(wb, "A2")).toBe(2)
    expect(get(wb, "A1")).toBe(1)
  })

  it("undoing a blocking entry restores the spill", () => {
    const wb = sheet({ A1: "=SEQUENCE(3)" })
    set(wb, { A2: "x" })
    wb.undo()
    expect(get(wb, "A1")).toBe(1)
    expect(get(wb, "A2")).toBe(2)
  })

  it("a spill that shrinks clears the cells it no longer covers", () => {
    const wb = sheet({ B1: "3", A1: "=SEQUENCE(B1)" })
    expect(get(wb, "A3")).toBe(3)
    set(wb, { B1: "2" })
    expect(get(wb, "A3")).toBeUndefined()
    expect(get(wb, "A2")).toBe(2)
  })

  it("arithmetic on ranges spills element by element", () => {
    const wb = sheet({ A1: "1", A2: "2", B1: "=A1:A2*10", C1: '=LEN({"ab","cde"})' })
    expect([get(wb, "B1"), get(wb, "B2")]).toEqual([10, 20])
    expect([get(wb, "C1"), get(wb, "D1")]).toEqual([2, 3])
  })

  it("deleting the formula removes its spill", () => {
    const wb = sheet({ A1: "=SEQUENCE(2,2)" })
    set(wb, { A1: "" })
    expect(get(wb, "B2")).toBeUndefined()
  })
})

describe("scale", () => {
  it("recalculates a 100,000-formula chain without running out of stack", () => {
    const wb = new Workbook()
    const cells: CellWrite[] = [{ row: 0, col: 0, data: { value: 1 } }]
    for (let r = 1; r < 100_000; r++) cells.push({ row: r, col: 0, data: { value: 0, input: `=A${r}+1` } })
    wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load" })
    expect(get(wb, "A100000")).toBe(100_000)
    const t0 = performance.now()
    set(wb, { A1: "5" })
    const ms = performance.now() - t0
    expect(get(wb, "A100000")).toBe(100_004)
    expect(ms).toBeLessThan(2000)
  }, 30_000) // loading 100k formulas is slow when the suite runs in parallel; the budget is checked in calc.bench
})
