import type { CalcMode, Op, ValuePatch } from "@cells/core"

/** Messages from the UI to the calculation worker. */
export type ToWorker =
  | { type: "init"; sheets: string[]; ops: Op[] }
  | { type: "ops"; ops: Op[] }
  | { type: "calculate"; full: boolean }
  | { type: "mode"; mode: CalcMode }

/** Messages from the worker: the values it calculated, after each message it handled. */
export type FromWorker = { type: "patch"; seq: number; patch: ValuePatch }

/** The two ends of a worker connection, so tests can connect the host to the engine without a real Worker. */
export type Port = {
  post(message: ToWorker): void
  listen(handler: (message: FromWorker) => void): void
  close?(): void
}
