import type { AiTarget } from "@cells/platform"
import { DEFAULT_ADDRESSES, type AiProviderName, type SharedSettings } from "@/settings/shared"

/** The provider names as File › Options › AI shows them (Inkspire's names). */
export const PROVIDER_LABELS: Record<AiProviderName, string> = {
  builtin: "Inkspire's built-in model",
  ollama: "Ollama (on this computer)",
  lmstudio: "LM Studio (on this computer)",
  compatible: "OpenAI-compatible endpoint",
  openai: "OpenAI",
  anthropic: "Anthropic",
}

/** Where on-device models run; the cloud ones get the cells sent to another company. */
export const LOCAL_PROVIDERS: AiProviderName[] = ["builtin", "ollama", "lmstudio"]

/** Where requests go with these settings, or what's missing. */
export function aiTarget(
  s: Pick<SharedSettings, "aiProvider" | "aiModels" | "aiAddresses">,
): AiTarget | { problem: string } {
  const model = s.aiProvider ? s.aiModels[s.aiProvider]?.trim() || undefined : undefined
  switch (s.aiProvider) {
    case "":
      return { problem: "Choose an AI provider in File › Options › AI." }
    case "builtin":
      return {
        problem:
          "Inkspire's built-in model doesn't run in Cells yet. Choose Ollama, LM Studio or another provider in File › Options › AI.",
      }
    case "ollama":
      return { provider: "ollama", model, baseUrl: s.aiAddresses.ollama?.trim() || DEFAULT_ADDRESSES.ollama }
    case "lmstudio":
      return { provider: "compatible", model, baseUrl: s.aiAddresses.lmstudio?.trim() || DEFAULT_ADDRESSES.lmstudio }
    case "compatible":
      return { provider: "compatible", model, baseUrl: s.aiAddresses.compatible?.trim() ?? "" }
    default:
      return { provider: s.aiProvider, model }
  }
}
