import { takePatch, Workbook } from "@cells/core"
import { openText, openXls, openXlsx } from "@cells/io"
import type { FromWorker, ToWorker } from "./protocol"

/**
 * The worker's side: a workbook that applies the operations the UI applied, calculates, and answers each message
 * with the values that changed. Kept free of worker globals so tests can run it in-process.
 */
export function createEngine(send: (message: FromWorker) => void) {
  let wb: Workbook | null = null
  let seq = 0
  // Opening a file is asynchronous; messages that arrive meanwhile wait their turn.
  let queue: Promise<void> = settled
  const handle = async (message: ToWorker) => {
    switch (message.type) {
      case "open":
        try {
          // Opened without calculating: what needs calculating (dynamic arrays, a CSV's formulas) is calculated
          // below, once changes are logged, so the values go back to the UI.
          wb = message.text
            ? (await openText(message.bytes, message.name, { ...message.text, calculate: false })).workbook
            : /\.xls$/i.test(message.name)
              ? (await openXls(message.bytes, message.name, { calculate: false })).workbook
              : (await openXlsx(message.bytes, message.name, { calculate: false })).workbook
        } catch {
          wb = new Workbook()
        }
        wb.calc.changeLog = new Set()
        wb.calculate(false)
        break
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
  return (message: ToWorker) => {
    // Without a pending open, answer at once (tests rely on the in-process engine answering synchronously).
    if (message.type !== "open" && queue === settled) {
      void handle(message)
      return
    }
    const next = queue.then(() => handle(message))
    queue = next
    void next.then(() => {
      if (queue === next) queue = settled
    })
  }
}

const settled: Promise<void> = Promise.resolve()
