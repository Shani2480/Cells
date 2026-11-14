import type { CalcMode, CalcStatus, Calculator, Change, Op, Workbook } from "@cells/core"
import { createEngine } from "./engine"
import type { FromWorker, Port } from "./protocol"

const IDLE: CalcStatus = { mode: "automatic", pending: false, circular: null, spills: [] }

/**
 * The UI's calculator when the engine runs in a worker. The UI's workbook applies every operation at once (so typing,
 * moving and undo stay instant) and does not calculate; the same operations go to the worker, whose answers (the
 * changed values) are written back as they arrive.
 */
export class WorkerCalculator implements Calculator {
  private current: CalcStatus = IDLE
  /** Messages sent and not yet answered: the status bar can show that a calculation is running. */
  outstanding = 0

  constructor(
    private readonly wb: Workbook<unknown>,
    private readonly port: Port,
  ) {
    port.listen((message) => this.receive(message))
    port.post({ type: "init", sheets: wb.sheets.map((s) => s.name), ops: snapshotOps(wb) })
    this.outstanding++
    wb.setCalculator(this)
  }

  afterOps(ops: readonly Op[], _changes: readonly Change[]) {
    if (!ops.length) return
    this.port.post({ type: "ops", ops: [...ops] })
    this.outstanding++
  }

  calculate(full: boolean) {
    this.port.post({ type: "calculate", full })
    this.outstanding++
  }

  setMode(mode: CalcMode) {
    this.current = { ...this.current, mode }
    this.port.post({ type: "mode", mode })
    this.outstanding++
  }

  status() {
    return this.current
  }

  private receive(message: FromWorker) {
    this.outstanding = Math.max(0, this.outstanding - 1)
    this.current = message.patch.status
    this.wb.applyPatch(message.patch)
  }
}

/** The workbook as it is now, as operations: every stored cell and every row/column override. */
function snapshotOps(wb: Workbook<unknown>): Op[] {
  const ops: Op[] = []
  wb.sheets.forEach((sheet, index) => {
    const used = sheet.cells.usedBounds()
    if (used) {
      const cells = sheet.cells.entries(used).map((e) => ({ row: e.row, col: e.col, data: { ...e.data } }))
      ops.push({ kind: "setCells", sheet: index, cells })
    }
    for (const axis of ["rows", "cols"] as const) {
      const lines = axis === "rows" ? sheet.rows : sheet.cols
      const entries = lines.overrides(0, lines.count - 1)
      if (entries.length) ops.push({ kind: "axis", sheet: index, axis, entries })
    }
  })
  return ops
}

/** A port to a real Web Worker. */
export function workerPort(worker: Worker): Port {
  return {
    post: (message) => worker.postMessage(message),
    listen: (handler) => {
      worker.onmessage = (event: MessageEvent<FromWorker>) => handler(event.data)
    },
    close: () => worker.terminate(),
  }
}

/** A port to an engine in this thread (tests; and a fallback where workers are unavailable). */
export function inProcessPort(): Port {
  let handler: ((message: FromWorker) => void) | null = null
  const engine = createEngine((message) => handler?.(message))
  return {
    post: (message) => engine(message),
    listen: (h) => {
      handler = h
    },
  }
}

/** Moves calculation of `wb` into a worker when the platform has workers; otherwise it stays in-process. */
export function startCalculation(wb: Workbook<unknown>): WorkerCalculator | null {
  if (typeof Worker === "undefined") return null
  const worker = new Worker(new URL("./worker.ts", import.meta.url), { type: "module" })
  return new WorkerCalculator(wb, workerPort(worker))
}
