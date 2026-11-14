import { takePatch, Workbook } from "@cells/core"
import type { FromWorker, ToWorker } from "./protocol"

/**
 * The worker's side: a workbook that applies the operations the UI applied, calculates, and answers each message
 * with the values that changed. Kept free of worker globals so tests can run it in-process.
 */
export function createEngine(send: (message: FromWorker) => void) {
  let wb: Workbook | null = null
  let seq = 0
  return (message: ToWorker) => {
    switch (message.type) {
      case "init":
        wb = new Workbook(message.sheets)
        wb.calc.changeLog = new Set()
        if (message.ops.length) wb.transact(message.ops, { label: "Load", history: false })
        break
      case "ops":
        wb?.transact(message.ops, { label: "Edit", history: false })
        break
      case "calculate":
        wb?.calculate(message.full)
        break
      case "mode":
        wb?.setCalcMode(message.mode)
        break
    }
    if (wb) send({ type: "patch", seq: ++seq, patch: takePatch(wb) })
  }
}
