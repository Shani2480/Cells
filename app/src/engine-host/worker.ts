/// <reference lib="webworker" />
import { createEngine } from "./engine"
import type { ToWorker } from "./protocol"

// The calculation engine, off the UI thread (PLAN.md §1.2).
const scope = self as unknown as DedicatedWorkerGlobalScope
const handle = createEngine((message) => scope.postMessage(message))
scope.onmessage = (event: MessageEvent<ToWorker>) => handle(event.data)
