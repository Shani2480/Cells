import { useMemo, useRef, useState } from "react"
import type { SheetController } from "@/sheet/controller"
import { Dialog } from "./Dialog"

/** Unicode blocks the Subset list offers, as Excel's dialog names them. */
export const SUBSETS: { name: string; from: number; to: number }[] = [
  { name: "Basic Latin", from: 0x20, to: 0x7e },
  { name: "Latin-1 Supplement", from: 0xa0, to: 0xff },
  { name: "Latin Extended-A", from: 0x100, to: 0x17f },
  { name: "Latin Extended-B", from: 0x180, to: 0x24f },
  { name: "IPA Extensions", from: 0x250, to: 0x2af },
  { name: "Spacing Modifier Letters", from: 0x2b0, to: 0x2ff },
  { name: "Combining Diacritical Marks", from: 0x300, to: 0x36f },
  { name: "Greek and Coptic", from: 0x370, to: 0x3ff },
  { name: "Cyrillic", from: 0x400, to: 0x4ff },
  { name: "Hebrew", from: 0x590, to: 0x5ff },
  { name: "Arabic", from: 0x600, to: 0x6ff },
  { name: "Latin Extended Additional", from: 0x1e00, to: 0x1eff },
  { name: "General Punctuation", from: 0x2000, to: 0x206f },
  { name: "Superscripts and Subscripts", from: 0x2070, to: 0x209f },
  { name: "Currency Symbols", from: 0x20a0, to: 0x20cf },
  { name: "Letterlike Symbols", from: 0x2100, to: 0x214f },
  { name: "Number Forms", from: 0x2150, to: 0x218f },
  { name: "Arrows", from: 0x2190, to: 0x21ff },
  { name: "Mathematical Operators", from: 0x2200, to: 0x22ff },
  { name: "Miscellaneous Technical", from: 0x2300, to: 0x23ff },
  { name: "Enclosed Alphanumerics", from: 0x2460, to: 0x24ff },
  { name: "Box Drawing", from: 0x2500, to: 0x257f },
  { name: "Block Elements", from: 0x2580, to: 0x259f },
  { name: "Geometric Shapes", from: 0x25a0, to: 0x25ff },
  { name: "Miscellaneous Symbols", from: 0x2600, to: 0x26ff },
  { name: "Dingbats", from: 0x2700, to: 0x27bf },
  { name: "Supplemental Arrows-B", from: 0x2900, to: 0x297f },
  { name: "Miscellaneous Mathematical Symbols-B", from: 0x2980, to: 0x29ff },
  { name: "CJK Symbols and Punctuation", from: 0x3000, to: 0x303f },
  { name: "Alphabetic Presentation Forms", from: 0xfb00, to: 0xfb4f },
  { name: "Halfwidth and Fullwidth Forms", from: 0xff00, to: 0xffef },
  { name: "Miscellaneous Symbols and Pictographs", from: 0x1f300, to: 0x1f5ff },
  { name: "Emoticons", from: 0x1f600, to: 0x1f64f },
  { name: "Transport and Map Symbols", from: 0x1f680, to: 0x1f6ff },
]

/** The Special Characters tab: Excel's list, with the names it gives them. */
export const SPECIAL: { char: string; name: string }[] = [
  { char: "—", name: "Em Dash" },
  { char: "–", name: "En Dash" },
  { char: "‑", name: "Nonbreaking Hyphen" },
  { char: "­", name: "Optional Hyphen" },
  { char: " ", name: "Em Space" },
  { char: " ", name: "En Space" },
  { char: " ", name: "1/4 Em Space" },
  { char: " ", name: "Nonbreaking Space" },
  { char: "©", name: "Copyright" },
  { char: "®", name: "Registered" },
  { char: "™", name: "Trademark" },
  { char: "§", name: "Section" },
  { char: "¶", name: "Paragraph" },
  { char: "…", name: "Ellipsis" },
  { char: "‘", name: "Single Opening Quote" },
  { char: "’", name: "Single Closing Quote" },
  { char: "“", name: "Double Opening Quote" },
  { char: "”", name: "Double Closing Quote" },
  { char: "​", name: "No-Width Optional Break" },
  { char: "⁠", name: "No-Width Non Break" },
]

const COLUMNS = 16
/** Spaces, soft hyphens and zero-width marks show nothing in the list. */
const INVISIBLE = /[\s\u00ad\u200b\u2060]/
const RECENT_KEY = "cells.symbols.recent"
/** Code points the grid leaves out: unassigned, controls, surrogates and private use. */
const HIDDEN = /[\p{Cn}\p{Cc}\p{Cs}\p{Co}]/u

/** The code points of a subset that have a character. */
export function subsetChars(from: number, to: number): number[] {
  const out: number[] = []
  for (let cp = from; cp <= to; cp++) if (!HIDDEN.test(String.fromCodePoint(cp))) out.push(cp)
  return out
}

/** "U+00A9" style hex for the Character code field. */
export const hexOf = (cp: number) => cp.toString(16).toUpperCase().padStart(4, "0")

/** The subset that holds a code point (for Character code typed in), or -1. */
export const subsetOf = (cp: number) => SUBSETS.findIndex((s) => cp >= s.from && cp <= s.to)

function loadRecent(): number[] {
  try {
    const list = JSON.parse(window.localStorage.getItem(RECENT_KEY) ?? "[]") as unknown
    return Array.isArray(list) ? list.filter((n): n is number => typeof n === "number").slice(0, COLUMNS) : []
  } catch {
    return []
  }
}

function saveRecent(list: number[]) {
  try {
    window.localStorage.setItem(RECENT_KEY, JSON.stringify(list))
  } catch {
    // Private window or blocked storage: the list just isn't kept.
  }
}

/**
 * Insert › Symbol: Excel's Symbol dialog, with its Symbols tab (a subset of Unicode in a grid, the recently used
 * symbols, the character code) and Special Characters tab. Insert puts the character in the cell being edited, or
 * opens the active cell for editing with it at the end; the dialog stays open for more, and Cancel becomes Close.
 */
export function SymbolDialog({ controller, onClose }: { controller: SheetController; onClose: () => void }) {
  const [tab, setTab] = useState<"symbols" | "special">("symbols")
  const [subset, setSubset] = useState(1)
  const [picked, setPicked] = useState(0xa9)
  const [special, setSpecial] = useState(0)
  const [recent, setRecent] = useState(loadRecent)
  const [code, setCode] = useState(hexOf(0xa9))
  const [inserted, setInserted] = useState(false)
  const gridRef = useRef<HTMLDivElement>(null)
  const chars = useMemo(() => subsetChars(SUBSETS[subset].from, SUBSETS[subset].to), [subset])

  const pick = (cp: number, focus = false) => {
    setPicked(cp)
    setCode(hexOf(cp))
    const at = subsetOf(cp)
    if (at >= 0 && at !== subset) setSubset(at)
    if (focus) requestAnimationFrame(() => gridRef.current?.querySelector<HTMLElement>(`[data-cp="${cp}"]`)?.focus())
  }

  const insert = () => {
    const text = tab === "symbols" ? String.fromCodePoint(picked) : SPECIAL[special].char
    controller.insertText(text)
    // A locked cell on a protected sheet says so and takes nothing.
    if (!controller.edit) return
    setInserted(true)
    if (tab === "symbols") {
      const next = [picked, ...recent.filter((n) => n !== picked)].slice(0, COLUMNS)
      setRecent(next)
      saveRecent(next)
    }
  }

  const onGridKey = (event: React.KeyboardEvent) => {
    const steps: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -COLUMNS, ArrowDown: COLUMNS }
    const step = steps[event.key]
    if (step === undefined) return
    event.preventDefault()
    const i = chars.indexOf(picked)
    const next = chars[Math.max(0, Math.min(chars.length - 1, (i < 0 ? 0 : i) + step))]
    if (next !== undefined) pick(next, true)
  }

  return (
    <Dialog
      title="Symbol"
      onCancel={onClose}
      onAccept={insert}
      width={560}
      actions={
        <>
          <button type="button" className="dialog-btn is-primary" onClick={insert}>
            <u>I</u>nsert
          </button>
          <button type="button" className="dialog-btn" onClick={onClose}>
            {inserted ? "Close" : "Cancel"}
          </button>
        </>
      }
    >
      <div className="fc-tabs" role="tablist">
        {(["symbols", "special"] as const).map((id) => (
          <button
            key={id}
            type="button"
            role="tab"
            aria-selected={tab === id}
            className="fc-tab"
            onClick={() => setTab(id)}
          >
            {id === "symbols" ? "Symbols" : "Special Characters"}
          </button>
        ))}
      </div>
      {tab === "symbols" ? (
        <div className="symbol-page">
          <div className="symbol-row">
            <label className="dialog-label" htmlFor="symbol-font">
              <u>F</u>ont:
            </label>
            <select
              id="symbol-font"
              className="dialog-field"
              disabled
              title="A cell shows its text in one font, so a symbol takes the cell's font."
            >
              <option>(normal text)</option>
            </select>
            <label className="dialog-label" htmlFor="symbol-subset">
              Su<u>b</u>set:
            </label>
            <select
              id="symbol-subset"
              className="dialog-field"
              value={subset}
              onChange={(e) => {
                const at = Number(e.target.value)
                setSubset(at)
                const first = subsetChars(SUBSETS[at].from, SUBSETS[at].to)[0]
                if (first !== undefined) {
                  setPicked(first)
                  setCode(hexOf(first))
                }
              }}
            >
              {SUBSETS.map((s, i) => (
                <option key={s.name} value={i}>
                  {s.name}
                </option>
              ))}
            </select>
          </div>
          <div
            ref={gridRef}
            className="symbol-grid"
            role="grid"
            aria-label={SUBSETS[subset].name}
            onKeyDown={onGridKey}
          >
            {chars.map((cp) => (
              <button
                key={cp}
                type="button"
                data-cp={cp}
                className={cp === picked ? "symbol-cell is-selected" : "symbol-cell"}
                tabIndex={cp === picked ? 0 : -1}
                aria-label={`U+${hexOf(cp)}`}
                aria-selected={cp === picked}
                onClick={() => pick(cp)}
                onDoubleClick={() => {
                  pick(cp)
                  insert()
                }}
              >
                {String.fromCodePoint(cp)}
              </button>
            ))}
          </div>
          <div className="dialog-label">
            R<u>e</u>cently used symbols:
          </div>
          <div className="symbol-recent">
            {Array.from({ length: COLUMNS }, (_, i) => recent[i]).map((cp, i) =>
              cp === undefined ? (
                <span key={i} className="symbol-cell is-empty" />
              ) : (
                <button
                  key={i}
                  type="button"
                  className={cp === picked ? "symbol-cell is-selected" : "symbol-cell"}
                  aria-label={`U+${hexOf(cp)}`}
                  onClick={() => pick(cp)}
                >
                  {String.fromCodePoint(cp)}
                </button>
              ),
            )}
          </div>
          <div className="symbol-row">
            <span className="symbol-name">Unicode character U+{hexOf(picked)}</span>
            <label className="dialog-label" htmlFor="symbol-code">
              <u>C</u>haracter code:
            </label>
            <input
              id="symbol-code"
              className="dialog-field symbol-code"
              value={code}
              maxLength={5}
              onChange={(e) => {
                const text = e.target.value.toUpperCase().replace(/[^0-9A-F]/g, "")
                setCode(text)
                const cp = parseInt(text, 16)
                if (text && cp <= 0x10ffff && !HIDDEN.test(String.fromCodePoint(cp))) {
                  setPicked(cp)
                  const at = subsetOf(cp)
                  if (at >= 0) setSubset(at)
                }
              }}
            />
            <span className="dialog-label">from: Unicode (hex)</span>
          </div>
        </div>
      ) : (
        <div className="dialog-list symbol-special" role="listbox" aria-label="Special characters">
          {SPECIAL.map((s, i) => (
            <button
              key={s.name}
              type="button"
              role="option"
              aria-selected={i === special}
              className={i === special ? "dialog-list-item is-selected" : "dialog-list-item"}
              onClick={() => setSpecial(i)}
              onDoubleClick={() => {
                setSpecial(i)
                controller.insertText(s.char)
                setInserted(true)
              }}
            >
              <span className="symbol-special-char">{INVISIBLE.test(s.char) ? "" : s.char}</span>
              {s.name}
            </button>
          ))}
        </div>
      )}
    </Dialog>
  )
}
