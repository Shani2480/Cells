import type { AiKeyStatus, AiPort, AiRequest, AiTarget, StopSignal } from "@cells/platform"

/** An AiPort for tests: answers with `reply` (streamed in two parts) and keeps what was asked. */
export class MemoryAi implements AiPort {
  readonly available = true
  reply = ""
  failWith: string | null = null
  readonly asked: AiRequest[] = []
  readonly tested: AiTarget[] = []
  keys: AiKeyStatus = { openai: "missing", anthropic: "missing", compatible: "missing", storage: "keyring" }

  async generate(request: AiRequest, onText: (text: string) => void, signal?: StopSignal) {
    this.asked.push(request)
    if (this.failWith) throw new Error(this.failWith)
    const half = Math.floor(this.reply.length / 2)
    onText(this.reply.slice(0, half))
    await Promise.resolve()
    if (signal?.aborted) throw new Error("Stopped.")
    onText(this.reply.slice(half))
    return this.reply
  }

  async test(target: AiTarget) {
    this.tested.push(target)
    return "Connected."
  }

  async keyStatus() {
    return this.keys
  }

  async saveKey(provider: "openai" | "anthropic" | "compatible", key: string) {
    this.keys = { ...this.keys, [provider]: key ? "saved" : "missing" }
    return key ? "keyring" : "none"
  }
}
