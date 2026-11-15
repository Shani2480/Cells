import { useId } from "react"
import { cn } from "@/lib/cn"
import {
  DATE_FORMAT_OPTIONS,
  LOCALE_OPTIONS,
  datePatternFor,
  initialsOf,
  localeLabel,
  separatorsFor,
  type SharedSettings,
  type ThemeChoice,
} from "./shared"

/** The fields the first-run setup and File › Options share, so both say and do the same. */
export type FieldsProps = {
  value: SharedSettings
  onChange: (patch: Partial<SharedSettings>) => void
}

const SAMPLE_DAY = new Date(2026, 0, 31)

/** A day written to a pattern such as `dd/MM/yyyy` (`yyyy`, `yy`, `MM`, `M`, `dd`, `d`). */
export function formatWithPattern(date: Date, pattern: string): string {
  const pad = (n: number) => String(n).padStart(2, "0")
  return pattern.replace(/yyyy|yy|MM|M|dd|d/g, (token) => {
    switch (token) {
      case "yyyy":
        return String(date.getFullYear())
      case "yy":
        return pad(date.getFullYear() % 100)
      case "MM":
        return pad(date.getMonth() + 1)
      case "M":
        return String(date.getMonth() + 1)
      case "dd":
        return pad(date.getDate())
      default:
        return String(date.getDate())
    }
  })
}

/** User name and initials; the initials follow the name until they're typed over. */
export function PersonalFields({ value, onChange }: FieldsProps) {
  const id = useId()
  return (
    <div className="settings-grid">
      <label className="dialog-label" htmlFor={`${id}-name`}>
        User name:
      </label>
      <input
        id={`${id}-name`}
        className="dialog-field"
        value={value.name}
        autoComplete="name"
        onChange={(e) => {
          const name = e.target.value
          const follows = value.initials === initialsOf(value.name)
          onChange(follows ? { name, initials: initialsOf(name) } : { name })
        }}
      />
      <label className="dialog-label" htmlFor={`${id}-initials`}>
        Initials:
      </label>
      <input
        id={`${id}-initials`}
        className="dialog-field is-short"
        value={value.initials}
        maxLength={6}
        onChange={(e) => onChange({ initials: e.target.value })}
      />
    </div>
  )
}

/**
 * Region, date order and the separators. Choosing a region fills in its date format and separators, which can then
 * be changed on their own; the samples show the result.
 */
export function RegionalFields({ value, onChange }: FieldsProps) {
  const id = useId()
  const locales =
    LOCALE_OPTIONS.includes(value.locale) || !value.locale ? LOCALE_OPTIONS : [value.locale, ...LOCALE_OPTIONS]
  const dates =
    DATE_FORMAT_OPTIONS.includes(value.dateFormat) || !value.dateFormat
      ? DATE_FORMAT_OPTIONS
      : [value.dateFormat, ...DATE_FORMAT_OPTIONS]
  const decimal = value.decimalSeparator || "."
  const group = value.groupSeparator
  const number = `1${group}234${group}567${decimal}89`
  return (
    <div className="settings-grid">
      <label className="dialog-label" htmlFor={`${id}-locale`}>
        Region:
      </label>
      <select
        id={`${id}-locale`}
        className="dialog-field"
        value={value.locale}
        onChange={(e) => {
          const locale = e.target.value
          const separators = separatorsFor(locale)
          onChange({
            locale,
            dateFormat: datePatternFor(locale),
            decimalSeparator: separators.decimal,
            groupSeparator: separators.group,
          })
        }}
      >
        {value.locale ? null : <option value="">United States (Excel's default)</option>}
        {locales.map((tag) => (
          <option key={tag} value={tag}>
            {localeLabel(tag)}
          </option>
        ))}
      </select>
      <label className="dialog-label" htmlFor={`${id}-date`}>
        Short date:
      </label>
      <div className="settings-inline">
        <select
          id={`${id}-date`}
          className="dialog-field"
          value={value.dateFormat || "M/d/yyyy"}
          onChange={(e) => onChange({ dateFormat: e.target.value })}
        >
          {dates.map((pattern) => (
            <option key={pattern} value={pattern}>
              {pattern}
            </option>
          ))}
        </select>
        <span className="settings-sample" aria-label="Sample date">
          {formatWithPattern(SAMPLE_DAY, value.dateFormat || "M/d/yyyy")}
        </span>
      </div>
      <label className="dialog-label" htmlFor={`${id}-decimal`}>
        Decimal separator:
      </label>
      <div className="settings-inline">
        <input
          id={`${id}-decimal`}
          className="dialog-field is-tiny"
          value={value.decimalSeparator || "."}
          maxLength={1}
          onChange={(e) => onChange({ decimalSeparator: e.target.value })}
        />
        <label className="dialog-label" htmlFor={`${id}-group`}>
          Thousands separator:
        </label>
        <input
          id={`${id}-group`}
          className="dialog-field is-tiny"
          value={value.groupSeparator}
          maxLength={1}
          onChange={(e) => onChange({ groupSeparator: e.target.value })}
        />
        <span className="settings-sample" aria-label="Sample number">
          {number}
        </span>
      </div>
    </div>
  )
}

const THEMES: { id: ThemeChoice; label: string }[] = [
  { id: "system", label: "Use system setting" },
  { id: "light", label: "White" },
  { id: "dark", label: "Black" },
]

/** Office Theme as Excel names its choices, each with a small picture of the window. */
export function ThemeField({ value, onChange }: FieldsProps) {
  return (
    <div className="settings-themes" role="radiogroup" aria-label="Office Theme">
      {THEMES.map((t) => (
        <label key={t.id} className={cn("settings-theme", value.theme === t.id && "is-current")}>
          <input
            type="radio"
            name="office-theme"
            checked={value.theme === t.id}
            onChange={() => onChange({ theme: t.id })}
          />
          <span className={cn("settings-theme-picture", `is-${t.id}`)} aria-hidden>
            <span className="bar" />
            {t.id === "system" ? (
              <span className="halves">
                <span className="grid" />
                <span className="grid is-dark" />
              </span>
            ) : (
              <span className={cn("grid", t.id === "dark" && "is-dark")} />
            )}
          </span>
          <span>{t.label}</span>
        </label>
      ))}
    </div>
  )
}

/** Checks a separator choice: one character each, and not the same for both. */
export function separatorProblem(value: SharedSettings): string | null {
  const decimal = value.decimalSeparator || "."
  if (/[\d\s-]/.test(decimal)) return "The decimal separator can't be a digit, a space or a minus sign."
  if (/[\d-]/.test(value.groupSeparator)) return "The thousands separator can't be a digit or a minus sign."
  if (decimal === value.groupSeparator) return "The decimal and thousands separators must be different."
  return null
}
