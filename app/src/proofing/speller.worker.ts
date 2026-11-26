import nspell from "nspell"

/** Reads a Hunspell dictionary away from the page (big ones take seconds) and answers about words. */
let s: ReturnType<typeof nspell> | null = null

type Message =
  | { kind: "load"; id: number; aff: string; dic: string; words: string[] }
  | { kind: "check"; id: number; words: string[] }
  | { kind: "suggest"; id: number; word: string }
  | { kind: "add"; id: number; word: string }

/** The worker's own scope (the app is typed against the DOM, not the worker library). */
const scope = self as unknown as {
  onmessage: ((e: MessageEvent<Message>) => void) | null
  postMessage(message: unknown): void
}

const correct = (word: string) => !!s && (s.correct(word) || s.correct(word.replace(/’/g, "'")))

scope.onmessage = (e: MessageEvent<Message>) => {
  const m = e.data
  try {
    if (m.kind === "load") {
      s = nspell({ aff: m.aff, dic: m.dic })
      for (const w of m.words) s.add(w)
      scope.postMessage({ id: m.id, result: null })
    } else if (m.kind === "check") scope.postMessage({ id: m.id, result: m.words.map(correct) })
    else if (m.kind === "suggest") scope.postMessage({ id: m.id, result: s ? s.suggest(m.word).slice(0, 8) : [] })
    else s?.add(m.word)
  } catch (error) {
    scope.postMessage({ id: m.id, result: null, error: String(error) })
  }
}
