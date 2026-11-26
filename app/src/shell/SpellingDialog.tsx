import { useEffect, useRef, useState } from "react"
import { BUILT_IN_LANGUAGE, languageLabel, usableLanguages, useDictionaries } from "@/proofing/dictionaries"
import { loadSpeller } from "@/proofing/speller"
import { updateSettings, useSettings } from "@/settings/settings"
import type { SheetController } from "@/sheet/controller"
import { SPELLING_MESSAGES, SpellingSession, type SpellingIssue } from "@/sheet/spelling"
import { Dialog } from "./Dialog"

/**
 * Review › Spelling (F7): Excel's Spelling dialog. The word not in the dictionary (which can be edited), the
 * suggestions, and Ignore Once, Ignore All, Add to Dictionary, Change and Change All. It ends with Excel's message.
 * Dictionary language lists the languages on this device; picking another checks on from the current word in it
 * (and keeps it for next time, as Excel does).
 */
export function SpellingDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const session = useRef<SpellingSession | null>(null)
  const [issue, setIssue] = useState<SpellingIssue | null>(null)
  const [failed, setFailed] = useState(false)
  const [word, setWord] = useState("")
  const [chosen, setChosen] = useState("")
  const [busy, setBusy] = useState(false)
  const installed = useDictionaries((d) => d.installed)
  const settings = useSettings()
  const languages = usableLanguages(installed)
  // A language whose dictionary isn't here (removed, or chosen in Inkspire) falls back to the built-in one.
  const [language, setLanguage] = useState(() =>
    languages.some((l) => l.id === settings.spellLanguage) ? settings.spellLanguage : BUILT_IN_LANGUAGE.id,
  )

  const show = async (pending: Promise<SpellingIssue | null>) => {
    setBusy(true)
    let next: SpellingIssue | null
    try {
      next = await pending
    } finally {
      setBusy(false)
    }
    if (!next) {
      onClose()
      controller.shell.alert(SPELLING_MESSAGES.done)
      return
    }
    setIssue(next)
    setWord(next.word)
    setChosen(next.suggestions[0] ?? "")
  }

  useEffect(() => {
    let live = true
    const st = useSettings.getState()
    loadSpeller(language, st.dictionaryWords)
      .then((speller) => {
        if (!live) return
        session.current = new SpellingSession(
          controller,
          speller,
          (w) => {
            const words = useSettings.getState().dictionaryWords
            if (!words.includes(w)) void updateSettings({ dictionaryWords: [...words, w] })
          },
          { ignoreUppercase: st.ignoreUppercase, ignoreNumbers: st.ignoreNumbers, ignoreAddresses: st.ignoreAddresses },
        )
        void show(session.current.next())
      })
      .catch(() => live && setFailed(true))
    return () => {
      live = false
    }
  }, [])

  const s = session.current
  // Change uses the edited word when it was edited, else the suggestion picked.
  const replacement = issue && word !== issue.word ? word : chosen
  const button = (label: React.ReactNode, run: () => void, disabled = false) => (
    <button type="button" className="dialog-btn" disabled={!s || !issue || busy || disabled} onClick={run}>
      {label}
    </button>
  )

  return (
    <Dialog
      title={`Spelling: ${languageLabel(language)}`}
      width={480}
      onCancel={onClose}
      onAccept={() => s && issue && replacement && !busy && void show(s.change(replacement))}
      actions={
        <button type="button" className="dialog-btn" onClick={onClose}>
          Cancel
        </button>
      }
    >
      {failed ? (
        <p className="dialog-note">The dictionary couldn't be read, so spelling can't be checked now.</p>
      ) : !issue ? (
        <p className="dialog-note">Reading the dictionary…</p>
      ) : (
        <div className="spelling">
          <label className="dialog-label" htmlFor="spelling-word">
            Not in Dictionary:{s?.where() ? <span className="spelling-where"> {s.where()}</span> : null}
          </label>
          <div />
          <input id="spelling-word" className="dialog-field" value={word} onChange={(e) => setWord(e.target.value)} />
          <div className="spelling-buttons">
            {button(
              <>
                <u>I</u>gnore Once
              </>,
              () => void show(s!.ignoreOnce()),
            )}
            {button(
              <>
                I<u>g</u>nore All
              </>,
              () => void show(s!.ignoreAll()),
            )}
            {button(
              <>
                <u>A</u>dd to Dictionary
              </>,
              () => void show(s!.addToDictionary()),
            )}
          </div>
          <span className="dialog-label">
            Suggestio<u>n</u>s:
          </span>
          <div />
          <div className="dialog-list spelling-list" role="listbox" aria-label="Suggestions">
            {issue.suggestions.length ? (
              issue.suggestions.map((sug) => (
                <button
                  type="button"
                  role="option"
                  aria-selected={sug === chosen}
                  key={sug}
                  className={`dialog-list-item${sug === chosen ? " is-selected" : ""}`}
                  onClick={() => (setChosen(sug), setWord(issue.word))}
                  onDoubleClick={() => void show(s!.change(sug))}
                >
                  {sug}
                </button>
              ))
            ) : (
              <span className="spelling-none">(No Suggestions)</span>
            )}
          </div>
          <div className="spelling-buttons">
            {button(
              <>
                <u>C</u>hange
              </>,
              () => void show(s!.change(replacement)),
              !replacement,
            )}
            {button(
              <>
                Chang<u>e</u> All
              </>,
              () => void show(s!.changeEvery(replacement)),
              !replacement,
            )}
          </div>
          <label className="dialog-label" htmlFor="spelling-language">
            Dictionary <u>l</u>anguage:
          </label>
          <div />
          <select
            id="spelling-language"
            className="dialog-field"
            disabled={!s || busy}
            value={language}
            onChange={(e) => {
              const next = e.target.value
              setLanguage(next)
              void updateSettings({ spellLanguage: next })
              void show(
                loadSpeller(next, useSettings.getState().dictionaryWords).then((sp) => session.current!.useSpeller(sp)),
              ).catch(() => setFailed(true))
            }}
          >
            {languages.map((l) => (
              <option key={l.id} value={l.id}>
                {l.label}
              </option>
            ))}
          </select>
        </div>
      )}
    </Dialog>
  )
}
