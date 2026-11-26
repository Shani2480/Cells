import { useEffect, useState } from "react"
import { usePlatform } from "@/platform/context"
import { Dialog } from "@/shell/Dialog"
import { updateSettings, useSettings } from "./settings"
import { AiFields } from "./AiFields"
import { SetupDictionary } from "./ProofingFields"
import { PersonalFields, RegionalFields, ThemeField, separatorProblem } from "./SettingsFields"
import { detectSettings, initialsOf, type SharedSettings } from "./shared"

type Step = { id: string; title: string; text: string }

const STEPS: Step[] = [
  {
    id: "name",
    title: "Your name",
    text: "Notes and comments you add are signed with this name.",
  },
  {
    id: "region",
    title: "Dates and numbers",
    text: "How you write dates and numbers, and the language Spelling checks. You can change them later in File › Options.",
  },
  {
    id: "theme",
    title: "Theme",
    text: "Light, dark, or the same as your system.",
  },
  {
    id: "ai",
    title: "AI (optional)",
    text: "AI can write a formula from your words, explain one, or suggest a chart. It stays off unless you turn it on, and you choose which AI it uses.",
  },
]

/**
 * The first-run setup: four short steps (name, region, theme, the AI opt-in), each starting from what Inkspire already saved in
 * the shared settings or, failing that, from what the system says. Skip keeps those guesses. It shows once; File ›
 * Options › Run setup again brings it back.
 */
export function FirstRunSetup({ onDone }: { onDone?: () => void }) {
  const platform = usePlatform()
  const stored = useSettings()
  const [draft, setDraft] = useState<SharedSettings | null>(null)
  const [fromInkspire, setFromInkspire] = useState(false)
  const [step, setStep] = useState(0)
  const [problem, setProblem] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false
    void platform.settings.systemFacts().then(
      (facts) => {
        if (cancelled) return
        const found = detectSettings(facts)
        const s = useSettings.getState()
        // What the person already set (here or in Inkspire) wins over a guess.
        setFromInkspire(Boolean(s.name || s.locale))
        const name = s.name || found.name
        setDraft({
          ...s,
          name,
          initials: s.initials || initialsOf(name),
          locale: s.locale || found.locale,
          dateFormat: s.dateFormat || found.dateFormat,
          decimalSeparator: s.decimalSeparator || found.decimalSeparator,
          groupSeparator: s.locale ? s.groupSeparator : found.groupSeparator,
        })
      },
      () => !cancelled && setDraft({ ...useSettings.getState() }),
    )
    return () => {
      cancelled = true
    }
  }, [platform])

  if (!draft) return null
  const change = (patch: Partial<SharedSettings>) => {
    setDraft({ ...draft, ...patch })
    setProblem(null)
  }
  const finish = (values: SharedSettings) => {
    void updateSettings({ ...values, name: values.name.trim(), initials: values.initials.trim(), setupDone: true })
    onDone?.()
  }
  const last = step === STEPS.length - 1
  const next = () => {
    if (STEPS[step].id === "region") {
      const wrong = separatorProblem(draft)
      if (wrong) {
        setProblem(wrong)
        return
      }
    }
    if (last) finish(draft)
    else setStep(step + 1)
  }
  // Skip keeps what was there (and the guesses for what wasn't), so a skipped setup still signs notes sensibly.
  const skip = () => finish(separatorProblem(draft) ? { ...draft, ...pickRegion(stored) } : draft)
  const current = STEPS[step]

  return (
    <Dialog
      title="Welcome to Cells"
      onCancel={skip}
      onAccept={next}
      width={560}
      actions={
        <>
          <button type="button" className="dialog-btn setup-skip" onClick={skip}>
            Skip
          </button>
          {step > 0 ? (
            <button type="button" className="dialog-btn" onClick={() => setStep(step - 1)}>
              Back
            </button>
          ) : null}
          <button type="button" className="dialog-btn is-primary" onClick={next}>
            {last ? "Finish" : "Next"}
          </button>
        </>
      }
    >
      <div className="setup">
        <ol className="setup-steps" aria-label="Setup steps">
          {STEPS.map((s, i) => (
            <li key={s.id} className={i === step ? "is-current" : i < step ? "is-done" : undefined}>
              <span className="setup-step-number" aria-hidden>
                {i + 1}
              </span>
              <span>{s.title}</span>
            </li>
          ))}
        </ol>
        <div className="setup-page">
          <h2 className="setup-title">{current.title}</h2>
          <p className="setup-text">{current.text}</p>
          {current.id === "name" ? <PersonalFields value={draft} onChange={change} /> : null}
          {current.id === "region" ? (
            <>
              <RegionalFields value={draft} onChange={change} />
              <SetupDictionary locale={draft.locale || (typeof navigator === "undefined" ? "" : navigator.language)} />
            </>
          ) : null}
          {current.id === "theme" ? (
            <ThemeField
              value={draft}
              onChange={(patch) => {
                change(patch)
                // The choice shows at once, behind the dialog.
                void updateSettings(patch)
              }}
            />
          ) : null}
          {current.id === "ai" ? <AiFields value={draft} onChange={change} compact /> : null}
          {problem ? (
            <p className="dialog-note is-error" role="alert">
              {problem}
            </p>
          ) : null}
          {fromInkspire && step === 0 ? (
            <p className="dialog-note">Filled in from the settings you share with Inkspire.</p>
          ) : null}
        </div>
      </div>
    </Dialog>
  )
}

const pickRegion = (s: SharedSettings): Partial<SharedSettings> => ({
  locale: s.locale,
  dateFormat: s.dateFormat,
  decimalSeparator: s.decimalSeparator,
  groupSeparator: s.groupSeparator,
})
