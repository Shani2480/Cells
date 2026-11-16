import { useEffect, useState } from "react"
import { usePlatform } from "@/platform/context"
import { cn } from "@/lib/cn"
import { Dialog } from "@/shell/Dialog"
import { updateSettings, useSettings } from "./settings"
import { AiFields } from "./AiFields"
import { PersonalFields, RegionalFields, ThemeField, separatorProblem } from "./SettingsFields"
import type { SharedSettings } from "./shared"

export type OptionsPage = "general" | "advanced" | "ai"

/**
 * Excel Options (File › Options), with the pages Cells has settings for: General (user name, initials, Office
 * Theme) and Advanced (the date format and separators; Excel takes the date format from the system's region, Cells
 * keeps its own so it can share it with Inkspire), and AI (Cells' own page: Excel has no such page, its Copilot is an
 * account setting). Everything here is shared with Inkspire, except whether the AI tools are on in Cells.
 */
export function OptionsDialog({
  onClose,
  onRunSetup,
  page: first = "general",
}: {
  onClose: () => void
  onRunSetup: () => void
  page?: OptionsPage
}) {
  const platform = usePlatform()
  const stored = useSettings()
  const [opened] = useState<SharedSettings>(() => ({ ...stored }))
  const [draft, setDraft] = useState<SharedSettings>(opened)
  const [page, setPage] = useState<OptionsPage>(first)
  const [problem, setProblem] = useState<string | null>(null)
  const [where, setWhere] = useState("")
  useEffect(() => {
    void platform.settings.location().then(setWhere, () => setWhere(""))
  }, [platform])

  const change = (patch: Partial<SharedSettings>) => {
    setDraft((d) => ({ ...d, ...patch }))
    setProblem(null)
  }
  const accept = () => {
    const wrong = separatorProblem(draft)
    if (wrong) {
      setPage("advanced")
      setProblem(wrong)
      return
    }
    // Only what was changed here, so a change Inkspire made while the dialog was open isn't undone.
    const result = { ...draft, name: draft.name.trim(), initials: draft.initials.trim() }
    const changed = Object.fromEntries(
      Object.entries(result).filter(([key, value]) => opened[key as keyof SharedSettings] !== value),
    )
    if (Object.keys(changed).length) void updateSettings(changed)
    onClose()
  }

  return (
    <Dialog title="Excel Options" onCancel={onClose} onAccept={accept} width={640}>
      <div className="options-layout">
        <nav className="options-pages" aria-label="Options pages">
          {(
            [
              ["general", "General"],
              ["advanced", "Advanced"],
              ["ai", "AI"],
            ] as const
          ).map(([id, label]) => (
            <button
              key={id}
              type="button"
              className={cn("options-page", page === id && "is-current")}
              aria-current={page === id ? "page" : undefined}
              onClick={() => setPage(id)}
            >
              {label}
            </button>
          ))}
        </nav>
        <div className="options-body">
          {page === "general" ? (
            <>
              <p className="options-intro">General options for working with Cells.</p>
              <h3 className="options-heading">Personalize your copy of Cells</h3>
              <PersonalFields value={draft} onChange={change} />
              <span className="dialog-label">Office Theme:</span>
              <ThemeField value={draft} onChange={change} />
              <h3 className="options-heading">Start up options</h3>
              <button type="button" className="dialog-btn options-setup" onClick={onRunSetup}>
                Run setup again…
              </button>
            </>
          ) : page === "ai" ? (
            <>
              <p className="options-intro">
                AI that helps with formulas, charts and summaries. Off until you turn it on.
              </p>
              <h3 className="options-heading">AI tools</h3>
              <AiFields value={draft} onChange={change} />
            </>
          ) : (
            <>
              <p className="options-intro">Advanced options for working with Cells.</p>
              <h3 className="options-heading">Region, dates and numbers</h3>
              <RegionalFields value={draft} onChange={change} />
              <p className="dialog-note">
                Dates and numbers you type are read this way, and cells show them so. Formulas keep Excel's English form
                (1.5, commas between arguments), and saved files and CSV files open the same in Excel anywhere.
              </p>
              {problem ? (
                <p className="dialog-note is-error" role="alert">
                  {problem}
                </p>
              ) : null}
            </>
          )}
          <p className="dialog-note options-shared">
            These settings are shared with Inkspire{where ? ` (${where})` : ""}.
          </p>
        </div>
      </div>
    </Dialog>
  )
}
