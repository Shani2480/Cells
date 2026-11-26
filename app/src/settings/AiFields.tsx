import { useEffect, useId, useState } from "react"
import type { AiKeyStatus } from "@cells/platform"
import { aiTarget, LOCAL_PROVIDERS, PROVIDER_LABELS } from "@/ai/client"
import { usePlatform } from "@/platform/context"
import type { FieldsProps } from "./SettingsFields"
import { AI_PROVIDERS, DEFAULT_ADDRESSES, type AddressedProvider, type AiProviderName } from "./shared"

const DEFAULT_MODELS: Record<AiProviderName, string> = {
  builtin: "",
  ollama: "llama3.2",
  lmstudio: "the loaded model",
  compatible: "default",
  openai: "gpt-5-mini",
  anthropic: "claude-sonnet-5-5",
}

const KEYED = ["openai", "anthropic", "compatible"] as const
type Keyed = (typeof KEYED)[number]

const isAddressed = (p: string): p is AddressedProvider => p in DEFAULT_ADDRESSES
const isKeyed = (p: string): p is Keyed => (KEYED as readonly string[]).includes(p)

/**
 * The AI tools' settings (File › Options › AI, and the setup's last step): on or off for Cells, and which AI to use,
 * shared with Inkspire. A key is saved at once into the system keyring (never into the settings file), so it isn't
 * undone by Cancel; everything else waits for OK.
 */
export function AiFields({ value, onChange, compact = false }: FieldsProps & { compact?: boolean }) {
  const id = useId()
  const platform = usePlatform()
  const [keys, setKeys] = useState<AiKeyStatus | null>(null)
  const [key, setKey] = useState("")
  const [message, setMessage] = useState<{ text: string; error: boolean } | null>(null)
  const [testing, setTesting] = useState(false)
  useEffect(() => {
    void platform.ai.keyStatus().then(setKeys)
  }, [platform])

  const provider = value.aiProvider
  const say = (text: string, error = false) => setMessage({ text, error })

  const recognition = (
    <label className="settings-check">
      <input type="checkbox" checked={value.ocrEnabled} onChange={(e) => onChange({ ocrEnabled: e.target.checked })} />
      Read scanned PDFs with text recognition on this {platform.os === "android" ? "device" : "computer"} (nothing is
      sent anywhere)
    </label>
  )

  if (!platform.ai.available)
    return (
      <div className="ai-fields">
        {recognition}
        <p className="dialog-note">The AI tools need the Cells desktop app; they aren't in the browser version.</p>
      </div>
    )

  const saveKey = async () => {
    if (!isKeyed(provider)) return
    try {
      const where = await platform.ai.saveKey(provider, key)
      setKey("")
      setKeys(await platform.ai.keyStatus())
      say(
        where === "none"
          ? "The key was removed."
          : where === "keyring"
            ? "The key is saved in the system keyring, for Cells and Inkspire."
            : "The key is saved in a file only you can read (there is no system keyring here).",
      )
    } catch (error) {
      say(error instanceof Error ? error.message : String(error), true)
    }
  }
  const test = async () => {
    const target = aiTarget(value)
    if ("problem" in target) return say(target.problem, true)
    setTesting(true)
    try {
      say(await platform.ai.test(target))
    } catch (error) {
      say(error instanceof Error ? error.message : String(error), true)
    } finally {
      setTesting(false)
    }
  }

  const keyState = isKeyed(provider) ? keys?.[provider] : undefined
  return (
    <div className="ai-fields">
      {recognition}
      <label className="settings-check">
        <input type="checkbox" checked={value.aiEnabled} onChange={(e) => onChange({ aiEnabled: e.target.checked })} />
        Turn on the AI tools (Home › AI Tools)
      </label>
      {!compact || value.aiEnabled ? (
        <div className="settings-grid">
          <label className="dialog-label" htmlFor={`${id}-provider`}>
            Use:
          </label>
          <select
            id={`${id}-provider`}
            className="dialog-field"
            value={provider}
            disabled={!value.aiEnabled}
            onChange={(e) => {
              setMessage(null)
              onChange({ aiProvider: e.target.value as AiProviderName })
            }}
          >
            {provider ? null : <option value="">Choose…</option>}
            {AI_PROVIDERS.map((p) => (
              <option key={p} value={p}>
                {PROVIDER_LABELS[p]}
              </option>
            ))}
          </select>
          {provider && provider !== "builtin" ? (
            <>
              <label className="dialog-label" htmlFor={`${id}-model`}>
                Model:
              </label>
              <input
                id={`${id}-model`}
                className="dialog-field"
                value={value.aiModels[provider] ?? ""}
                placeholder={DEFAULT_MODELS[provider]}
                spellCheck={false}
                disabled={!value.aiEnabled}
                onChange={(e) => onChange({ aiModels: { ...value.aiModels, [provider]: e.target.value } })}
              />
            </>
          ) : null}
          {isAddressed(provider) ? (
            <>
              <label className="dialog-label" htmlFor={`${id}-address`}>
                Address:
              </label>
              <input
                id={`${id}-address`}
                className="dialog-field"
                value={value.aiAddresses[provider] ?? ""}
                placeholder={DEFAULT_ADDRESSES[provider] || "https://…/v1"}
                spellCheck={false}
                disabled={!value.aiEnabled}
                onChange={(e) => onChange({ aiAddresses: { ...value.aiAddresses, [provider]: e.target.value } })}
              />
            </>
          ) : null}
          {isKeyed(provider) ? (
            <>
              <label className="dialog-label" htmlFor={`${id}-key`}>
                API key:
              </label>
              <div className="settings-inline">
                <input
                  id={`${id}-key`}
                  className="dialog-field"
                  type="password"
                  autoComplete="off"
                  value={key}
                  disabled={!value.aiEnabled}
                  placeholder={
                    keyState === "saved"
                      ? "Saved (type a new one to replace it)"
                      : keyState === "environment"
                        ? "From the environment"
                        : provider === "compatible"
                          ? "If the endpoint needs one"
                          : ""
                  }
                  onChange={(e) => setKey(e.target.value)}
                />
                <button
                  type="button"
                  className="dialog-btn"
                  disabled={!value.aiEnabled || (!key.trim() && keyState !== "saved")}
                  onClick={() => void saveKey()}
                >
                  {key.trim() || keyState !== "saved" ? "Save Key" : "Remove Key"}
                </button>
              </div>
            </>
          ) : null}
        </div>
      ) : null}
      {provider === "builtin" && value.aiEnabled ? (
        <p className="dialog-note">
          Inkspire's built-in model doesn't run in Cells yet. Choose Ollama or LM Studio to keep everything on this
          computer, or a service you have a key for.
        </p>
      ) : provider && value.aiEnabled ? (
        <p className="dialog-note">
          {LOCAL_PROVIDERS.includes(provider)
            ? "Runs on this computer: your cells stay here."
            : `The cells a tool is about are sent to ${PROVIDER_LABELS[provider]}.`}{" "}
          Inkspire uses the same choice and keys.
        </p>
      ) : null}
      {provider && provider !== "builtin" && value.aiEnabled ? (
        <div className="settings-inline">
          <button type="button" className="dialog-btn" disabled={testing} onClick={() => void test()}>
            {testing ? "Testing…" : "Test Connection"}
          </button>
        </div>
      ) : null}
      {message ? (
        <p className={message.error ? "dialog-note is-error" : "dialog-note"} role={message.error ? "alert" : "status"}>
          {message.text}
        </p>
      ) : null}
    </div>
  )
}
