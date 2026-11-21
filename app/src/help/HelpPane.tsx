import { ArrowLeft, ArrowRight, GraduationCap, Headset, MessageCircle, Search, Sparkles, X } from "lucide-react"
import { useEffect, useRef, useState } from "react"
import { hintTarget } from "@/hints/FeatureHints"
import { pointAt } from "@/hints/pointer"
import { usePlatform } from "@/platform/context"
import type { SheetController } from "@/sheet/controller"
import { FEATURE_TOPICS, keysText, searchHelp, tabOf, whereIs, type FeatureTopic } from "./help-content"

type Props = {
  controller: SheetController
  onDone: () => void
  /** Selects a ribbon tab (to show a feature's button before pointing at it). */
  onShowTab: (tab: string) => void
  /** Opens a dialog of the shell ("feedback", "contactSupport") or the File view's What's New. */
  onDialog: (name: string) => void
  onWhatsNew: () => void
}

/** How long a feature's tab has to draw before "Show me" looks for its button. */
const SHOW_DELAY_MS = 150

/**
 * Help › Help (F1): Cells' own help, offline. Search, the special features (each with where it is and "Show me",
 * which points at its button with the hint callout) and the keyboard shortcuts; then Show Training, What's New,
 * Contact Support and Feedback. Show Training is the same pane walking through the features one by one.
 */
export function HelpPane({ controller, onDone, onShowTab, onDialog, onWhatsNew }: Props) {
  const mac = usePlatform().os === "macos"
  const training = controller.helpPane === "training"
  // On a phone the pane fills the screen and there's no ribbon: no pointing, the text says where things are.
  const phone = document.documentElement.classList.contains("is-phone")
  const [query, setQuery] = useState("")
  const [missing, setMissing] = useState<string | null>(null)
  const [step, setStep] = useState(0)
  const searchBox = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (!training) searchBox.current?.focus()
  }, [training])

  // Nothing stays pointed at once the pane closes or changes.
  useEffect(() => () => pointAt(null), [])

  const show = (topic: FeatureTopic, stepText?: string) => {
    setMissing(null)
    if (phone) return
    const tab = topic.command ? tabOf(topic.command) : undefined
    if (tab) onShowTab(tab)
    window.setTimeout(() => {
      if (!hintTarget(topic.target)) {
        pointAt(null)
        setMissing(topic.id)
        return
      }
      pointAt({ id: topic.id, title: topic.title, text: topic.text, target: topic.target, step: stepText })
    }, SHOW_DELAY_MS)
  }

  const stepText = (i: number) => `${i + 1} of ${FEATURE_TOPICS.length}`
  useEffect(() => {
    if (!training) return
    show(FEATURE_TOPICS[step], stepText(step))
    // Pointing again only when the step changes or the tour starts.
  }, [training, step])

  const close = () => {
    pointAt(null)
    controller.setHelpPane(null)
    onDone()
  }
  const startTraining = () => {
    setStep(0)
    controller.setHelpPane("training")
  }
  const endTraining = () => {
    pointAt(null)
    setMissing(null)
    controller.setHelpPane("help")
  }

  const results = searchHelp(query)
  const searching = query.trim() !== ""
  const topic = FEATURE_TOPICS[step]

  return (
    <aside
      className="comments-pane help-pane"
      aria-label={training ? "Training" : "Help"}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>{training ? "Training" : "Help"}</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>

      {training ? (
        <div className="help-training">
          <p className="help-step">{stepText(step)}</p>
          <h3 className="help-training-title">{topic.title}</h3>
          <p className="help-text">{topic.text}</p>
          {whereIs(topic) ? <p className="help-where">Where: {whereIs(topic)}</p> : null}
          {missing === topic.id ? (
            <p className="help-note" role="status">
              It isn't on screen now, so there's nothing to point at.
            </p>
          ) : null}
          <div className="help-training-nav">
            <button
              type="button"
              className="dialog-btn"
              disabled={step === 0}
              onClick={() => setStep((s) => Math.max(0, s - 1))}
            >
              <ArrowLeft size={14} strokeWidth={1.75} aria-hidden /> Back
            </button>
            {step < FEATURE_TOPICS.length - 1 ? (
              <button type="button" className="dialog-btn is-primary" onClick={() => setStep((s) => s + 1)}>
                Next <ArrowRight size={14} strokeWidth={1.75} aria-hidden />
              </button>
            ) : (
              <button type="button" className="dialog-btn is-primary" onClick={endTraining}>
                Done
              </button>
            )}
          </div>
          <button type="button" className="help-link" onClick={endTraining}>
            Back to Help
          </button>
        </div>
      ) : (
        <div className="help-body">
          <div className="help-search">
            <Search size={14} strokeWidth={1.75} aria-hidden />
            <input
              type="search"
              aria-label="Search Help"
              placeholder="Search Help"
              ref={searchBox}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>

          <section aria-labelledby="help-features">
            <h3 id="help-features" className="help-heading">
              Features in Cells
            </h3>
            {results.topics.length === 0 ? <p className="help-note">No features match.</p> : null}
            {results.topics.map((t) => (
              <details key={t.id} className="help-topic" open={searching || undefined}>
                <summary>{t.title}</summary>
                <p className="help-text">{t.text}</p>
                {whereIs(t) ? <p className="help-where">Where: {whereIs(t)}</p> : null}
                {phone ? null : (
                  <button type="button" className="help-link" onClick={() => show(t)}>
                    Show me
                  </button>
                )}
                {missing === t.id ? (
                  <p className="help-note" role="status">
                    It isn't on screen now.{t.where ? ` ${t.where}.` : ""}
                  </p>
                ) : null}
              </details>
            ))}
          </section>

          <section aria-labelledby="help-shortcuts">
            <h3 id="help-shortcuts" className="help-heading">
              Keyboard shortcuts
            </h3>
            {results.shortcuts.length === 0 ? <p className="help-note">No shortcuts match.</p> : null}
            {results.shortcuts.map((group) => (
              <details key={group.title} className="help-topic" open={searching || undefined}>
                <summary>{group.title}</summary>
                <table className="help-keys">
                  <tbody>
                    {group.shortcuts.map((s) => (
                      <tr key={s.keys}>
                        <th scope="row">
                          <kbd>{keysText(s.keys, mac)}</kbd>
                        </th>
                        <td>{s.does}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </details>
            ))}
          </section>

          <section className="help-more" aria-label="More help">
            <button type="button" className="help-action" onClick={startTraining}>
              <GraduationCap size={16} strokeWidth={1.75} aria-hidden /> Show Training
            </button>
            <button type="button" className="help-action" onClick={onWhatsNew}>
              <Sparkles size={16} strokeWidth={1.75} aria-hidden /> What's New
            </button>
            <button type="button" className="help-action" onClick={() => onDialog("contactSupport")}>
              <Headset size={16} strokeWidth={1.75} aria-hidden /> Contact Support
            </button>
            <button type="button" className="help-action" onClick={() => onDialog("feedback")}>
              <MessageCircle size={16} strokeWidth={1.75} aria-hidden /> Feedback
            </button>
          </section>
        </div>
      )}
    </aside>
  )
}
