import { useState } from "react"
import { Download, Trash2 } from "lucide-react"
import { usePlatform } from "@/platform/context"
import { updateSettings } from "./settings"
import {
  BUILT_IN_LANGUAGE,
  dictionaryEntry,
  dictionaryForLocale,
  DOWNLOADABLE_DICTIONARIES,
  downloadDictionary,
  downloadSize,
  removeDictionary,
  usableLanguages,
  useDictionaries,
} from "@/proofing/dictionaries"
import { AutoCorrectDialog } from "./AutoCorrectDialog"
import type { SharedSettings } from "./shared"

type Props = { value: SharedSettings; onChange: (patch: Partial<SharedSettings>) => void }

/** Excel's Proofing options: the words Spelling passes over, and which dictionary it uses. */
export function ProofingFields({ value, onChange }: Props) {
  const [autoCorrect, setAutoCorrect] = useState(false)
  const installed = useDictionaries((d) => d.installed)
  const languages = usableLanguages(installed)
  const check = (key: "ignoreUppercase" | "ignoreNumbers" | "ignoreAddresses", label: string) => (
    <label className="dialog-radio">
      <input type="checkbox" checked={value[key]} onChange={(e) => onChange({ [key]: e.target.checked })} />
      {label}
    </label>
  )
  return (
    <>
      <h3 className="options-heading">AutoCorrect options</h3>
      <span className="options-row">
        Change how Cells corrects and formats text as you type:
        <button type="button" className="dialog-btn" onClick={() => setAutoCorrect(true)}>
          AutoCorrect Options…
        </button>
      </span>
      {autoCorrect ? (
        <AutoCorrectDialog
          value={value.autoCorrect}
          onCancel={() => setAutoCorrect(false)}
          onDone={(next) => {
            setAutoCorrect(false)
            onChange({ autoCorrect: next })
          }}
        />
      ) : null}
      <h3 className="options-heading">When correcting spelling</h3>
      {check("ignoreUppercase", "Ignore words in UPPERCASE")}
      {check("ignoreNumbers", "Ignore words that contain numbers")}
      {check("ignoreAddresses", "Ignore Internet and file addresses")}
      <label className="dialog-label" htmlFor="options-dictionary">
        Dictionary language:
      </label>
      <select
        id="options-dictionary"
        className="dialog-field options-dictionary"
        value={languages.some((l) => l.id === value.spellLanguage) ? value.spellLanguage : BUILT_IN_LANGUAGE.id}
        onChange={(e) => onChange({ spellLanguage: e.target.value })}
      >
        {languages.map((l) => (
          <option key={l.id} value={l.id}>
            {l.label}
          </option>
        ))}
      </select>
      <DictionaryList />
    </>
  )
}

/**
 * The languages that can be downloaded: each with its size and its word list's licence, and Download or Remove.
 * Downloads happen at once (not on OK) and stay on the device.
 */
export function DictionaryList({ only }: { only?: string[] }) {
  const platform = usePlatform()
  const { installed, downloading, failed } = useDictionaries()
  const list = DOWNLOADABLE_DICTIONARIES.filter((d) => !only || only.includes(d.id))
  return (
    <>
      {!only ? <h3 className="options-heading">More spelling languages</h3> : null}
      <ul className="dictionary-list" aria-label="Spelling languages">
        {list.map((d) => {
          const have = installed.includes(d.id)
          const share = downloading[d.id]
          return (
            <li key={d.id} className="dictionary-item">
              <span className="dictionary-name">{d.label}</span>
              <span className="dictionary-meta">
                {downloadSize(d)} · {d.license.replace(/[()]/g, "")}
              </span>
              {share !== undefined ? (
                <span className="dictionary-progress" role="status">
                  Downloading… {Math.round(share * 100)}%
                </span>
              ) : have ? (
                <button
                  type="button"
                  className="dialog-btn dictionary-btn"
                  aria-label={`Remove ${d.label}`}
                  onClick={() => void removeDictionary(d.id)}
                >
                  <Trash2 size={14} aria-hidden /> Remove
                </button>
              ) : (
                <button
                  type="button"
                  className="dialog-btn dictionary-btn"
                  aria-label={`Download ${d.label}`}
                  onClick={() => void downloadDictionary(d.id, (url) => platform.web.fetch(url))}
                >
                  <Download size={14} aria-hidden /> Download
                </button>
              )}
              {failed?.id === d.id ? (
                <span className="dialog-note is-error dictionary-error" role="alert">
                  {failed.message}
                </span>
              ) : null}
            </li>
          )
        })}
      </ul>
      <p className="dialog-note">
        Dictionaries come from the open-source Hunspell word lists (via jsDelivr) and are checked before they're kept.
        Spelling works offline once a language is here.
      </p>
    </>
  )
}

/**
 * The setup's offer, under Dates and numbers: the dictionary for the region picked, when Cells can download one.
 * Downloaded, it becomes Spelling's language.
 */
export function SetupDictionary({ locale }: { locale: string }) {
  const platform = usePlatform()
  const { installed, downloading, failed } = useDictionaries()
  const id = dictionaryForLocale(locale)
  const entry = id ? dictionaryEntry(id) : undefined
  if (!entry) return null
  const share = downloading[entry.id]
  return (
    <div className="setup-dictionary">
      <span>
        Spelling in {entry.label}: a {downloadSize(entry)} download.
      </span>
      {installed.includes(entry.id) ? (
        <span role="status">Downloaded</span>
      ) : share !== undefined ? (
        <span role="status">Downloading… {Math.round(share * 100)}%</span>
      ) : (
        <button
          type="button"
          className="dialog-btn"
          onClick={() =>
            void downloadDictionary(entry.id, (url) => platform.web.fetch(url)).then((ok) => {
              if (ok) void updateSettings({ spellLanguage: entry.id })
            })
          }
        >
          <Download size={14} aria-hidden /> Download
        </button>
      )}
      {failed?.id === entry.id ? (
        <span className="dialog-note is-error" role="alert">
          {failed.message}
        </span>
      ) : null}
    </div>
  )
}
