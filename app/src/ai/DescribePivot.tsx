import type { PivotCache, PivotRecipe } from "@cells/core"
import { Sparkles } from "lucide-react"
import { useEffect, useRef, useState } from "react"
import { usePlatform } from "@/platform/context"
import { useSettings } from "@/settings/settings"
import { aiTarget, LOCAL_PROVIDERS, PROVIDER_LABELS } from "./client"
import { pivotRequest, readPivotReply } from "./pivot-prompt"

/** Is AI on and set up, so "Describe the PivotTable you want" can be offered? (It's opt-in: hidden otherwise.) */
export function useAiReady(): boolean {
  const platform = usePlatform()
  const settings = useSettings()
  return platform.ai.available && settings.aiEnabled && !("problem" in aiTarget(settings))
}

type Run = { state: "idle" } | { state: "running" } | { state: "failed"; message: string }

/**
 * "Describe the PivotTable you want": a line to type in and Ask. The AI's layout comes back through `onRecipe` (for a
 * preview; nothing is made until the person chooses OK or Apply). Says what is sent where. Renders nothing when AI is
 * off or not set up.
 */
export function DescribePivot({
  cache,
  onRecipe,
  autoFocus,
}: {
  cache: PivotCache
  onRecipe: (recipe: PivotRecipe) => void
  autoFocus?: boolean
}) {
  const platform = usePlatform()
  const settings = useSettings()
  const ready = useAiReady()
  const [text, setText] = useState("")
  const [run, setRun] = useState<Run>({ state: "idle" })
  const stop = useRef<AbortController | null>(null)
  useEffect(() => () => stop.current?.abort(), [])
  if (!ready) return null
  const target = aiTarget(settings)
  if ("problem" in target) return null
  const provider = settings.aiProvider as Exclude<typeof settings.aiProvider, "">
  const local = LOCAL_PROVIDERS.includes(provider)

  const ask = () => {
    if (!text.trim() || run.state === "running") return
    stop.current?.abort()
    const signal = new AbortController()
    stop.current = signal
    setRun({ state: "running" })
    platform.ai
      .generate({ ...pivotRequest(cache, text, cache.records.length), target }, () => {}, signal.signal)
      .then(
        (reply) => {
          if (signal.signal.aborted) return
          const read = readPivotReply(reply, cache)
          if ("problem" in read) return setRun({ state: "failed", message: read.problem })
          setRun({ state: "idle" })
          onRecipe(read.recipe)
        },
        (error: unknown) => {
          if (!signal.signal.aborted)
            setRun({ state: "failed", message: error instanceof Error ? error.message : String(error) })
        },
      )
  }
  const running = run.state === "running"

  return (
    <div className="describe-pivot">
      <label className="dialog-label describe-pivot-label" htmlFor="describe-pivot">
        <Sparkles size={14} strokeWidth={1.75} aria-hidden />
        Describe the PivotTable you want
      </label>
      <div className="describe-pivot-row">
        <input
          id="describe-pivot"
          className="dialog-field"
          data-autofocus={autoFocus || undefined}
          value={text}
          placeholder="e.g. sales by region per quarter, biggest first"
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            // Enter asks here, instead of the dialog's OK.
            if (e.key === "Enter") {
              e.preventDefault()
              e.stopPropagation()
              ask()
            }
          }}
        />
        {running ? (
          <button
            type="button"
            className="dialog-btn"
            onClick={() => {
              stop.current?.abort()
              setRun({ state: "idle" })
            }}
          >
            Stop
          </button>
        ) : (
          <button type="button" className="dialog-btn" disabled={!text.trim()} onClick={ask}>
            Ask
          </button>
        )}
      </div>
      {running ? (
        <p className="dialog-note" aria-live="polite">
          Asking {PROVIDER_LABELS[provider]}…
        </p>
      ) : null}
      {run.state === "failed" ? (
        <p className="dialog-note is-error" role="alert">
          {run.message}
        </p>
      ) : null}
      <p className="dialog-note describe-pivot-notice">
        Sends the column names and a few values from each to {PROVIDER_LABELS[provider]}
        {local ? "" : ", outside this computer"}. The AI only picks the layout: Cells works out the numbers.
      </p>
    </div>
  )
}
