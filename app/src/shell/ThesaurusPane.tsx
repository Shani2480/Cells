import { ArrowLeft, Search, X } from "lucide-react"
import { useEffect, useState } from "react"
import { loadThesaurus, type Meaning, type Thesaurus } from "@/proofing/thesaurus"
import { matchCase, wordsIn } from "@/proofing/words"
import type { SheetController } from "@/sheet/controller"

/** The word the Thesaurus starts with: the active cell's text when it's one word. */
export function startWord(controller: SheetController): string {
  const { row, col } = controller.selection.active
  const v = controller.sheet.cells.value(row, col)
  if (typeof v !== "string") return ""
  const words = wordsIn(v, { ignoreUppercase: false, ignoreNumbers: true, ignoreAddresses: true })
  return words.length === 1 ? words[0].word : ""
}

/**
 * Review › Thesaurus (Shift+F7): Excel's Thesaurus pane. A word's meanings, each with its synonyms (and opposites);
 * a click looks a word up, Insert puts it in the active cell in place of the word looked up.
 */
export function ThesaurusPane({ controller, onDone }: { controller: SheetController; onDone: () => void }) {
  const [query, setQuery] = useState(() => startWord(controller))
  const [shown, setShown] = useState(query)
  const [history, setHistory] = useState<string[]>([])
  const [thesaurus, setThesaurus] = useState<Thesaurus | null>(null)
  const [failed, setFailed] = useState(false)
  useEffect(() => {
    let live = true
    loadThesaurus()
      .then((t) => live && setThesaurus(t))
      .catch(() => live && setFailed(true))
    return () => {
      live = false
    }
  }, [])

  const look = (word: string) => {
    if (shown) setHistory((h) => [...h, shown])
    setQuery(word)
    setShown(word)
  }
  const close = () => {
    controller.setThesaurusPane(false)
    onDone()
  }
  const insert = (word: string) => {
    const { row, col } = controller.selection.active
    const v = controller.sheet.cells.value(row, col)
    const text = typeof v === "string" ? v : ""
    const target = history[0] ?? shown
    const found = wordsIn(text, { ignoreUppercase: false, ignoreNumbers: false, ignoreAddresses: true }).find(
      (w) => w.word.toLowerCase() === target.toLowerCase(),
    )
    const next = found ? text.slice(0, found.start) + matchCase(found.word, word) + text.slice(found.end) : word
    controller.typeInto({ row, col }, next, "Thesaurus")
  }
  const meanings: Meaning[] = thesaurus && shown ? thesaurus.lookup(shown) : []

  return (
    <aside
      className="comments-pane thesaurus-pane"
      aria-label="Thesaurus"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault()
          close()
        }
      }}
    >
      <header className="comments-pane-head">
        <h2>Thesaurus</h2>
        <button type="button" className="comments-pane-close" aria-label="Close" title="Close" onClick={close}>
          <X size={16} strokeWidth={1.75} />
        </button>
      </header>
      <form
        className="thesaurus-search"
        onSubmit={(e) => {
          e.preventDefault()
          if (query.trim()) look(query.trim())
        }}
      >
        <button
          type="button"
          className="comments-pane-close"
          aria-label="Back"
          title="Back"
          disabled={!history.length}
          onClick={() => {
            const prev = history[history.length - 1]
            setHistory((h) => h.slice(0, -1))
            setQuery(prev)
            setShown(prev)
          }}
        >
          <ArrowLeft size={16} strokeWidth={1.75} />
        </button>
        <input
          className="dialog-field"
          aria-label="Search for"
          placeholder="Search for a word"
          value={query}
          autoFocus
          onChange={(e) => setQuery(e.target.value)}
        />
        <button type="submit" className="comments-pane-close" aria-label="Search" title="Search">
          <Search size={16} strokeWidth={1.75} />
        </button>
      </form>
      <div className="comments-pane-list thesaurus-list">
        {failed ? (
          <p className="comments-pane-empty">The thesaurus couldn't be read.</p>
        ) : !thesaurus ? (
          <p className="comments-pane-empty">Reading the thesaurus…</p>
        ) : !shown ? (
          <p className="comments-pane-empty">Type a word to look for its synonyms.</p>
        ) : !meanings.length ? (
          <p className="comments-pane-empty">
            No results found for “{shown}”. Check the spelling, or try another word.
          </p>
        ) : (
          meanings.map((m, i) => (
            <section key={i} className="thesaurus-meaning">
              <h3>
                {m.heading} <span className="thesaurus-pos">({m.pos})</span>
              </h3>
              <ul>
                {m.synonyms.map((w) => (
                  <WordRow key={w} word={w} onLook={look} onInsert={insert} />
                ))}
              </ul>
              {m.antonyms.length ? (
                <>
                  <h4>Antonyms</h4>
                  <ul>
                    {m.antonyms.map((w) => (
                      <WordRow key={w} word={w} onLook={look} onInsert={insert} />
                    ))}
                  </ul>
                </>
              ) : null}
            </section>
          ))
        )}
      </div>
      <footer className="thesaurus-foot">English (United States) · WordNet</footer>
    </aside>
  )
}

function WordRow({
  word,
  onLook,
  onInsert,
}: {
  word: string
  onLook: (w: string) => void
  onInsert: (w: string) => void
}) {
  return (
    <li className="thesaurus-word">
      <button type="button" className="thesaurus-look" onClick={() => onLook(word)}>
        {word}
      </button>
      <button type="button" className="thesaurus-insert" onClick={() => onInsert(word)}>
        Insert
      </button>
    </li>
  )
}
