/**
 * Flash Fill: learn how a column is made from the columns beside it, from a few examples, and fill in the rest.
 *
 * A program is a row of pieces, each either fixed text or a span of words of one source column (counted from the
 * start or the end, so "the last word" works on names of any length), optionally only its first letters and in
 * another case. The search runs over all examples at once (a state is how far each example's output is made), cheapest
 * program first, so the program found fits every example; the cheapest prefers few pieces, taken text over typed text.
 */

export type FlashCase = "same" | "lower" | "upper" | "proper"

export type FlashPiece =
  | { kind: "text"; text: string }
  | {
      kind: "span"
      col: number
      /** First and last word; 0 is the first, -1 the last. */
      from: number
      to: number
      /** Only the first `first` characters (initials). */
      first?: number
      case: FlashCase
    }

/** Words, numbers, runs of spaces and single other characters ("McDonald" is two words, "NASA" one). */
export function flashTokens(text: string): string[] {
  return text.match(/\p{Lu}?\p{Ll}+|\p{Lu}+(?!\p{Ll})|\p{L}+|\d+|\s+|[^\p{L}\d\s]/gu) ?? []
}

function changeCase(text: string, c: FlashCase): string {
  if (c === "lower") return text.toLowerCase()
  if (c === "upper") return text.toUpperCase()
  if (c === "proper")
    return text.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (_, a: string, b: string) => a + b.toUpperCase())
  return text
}

function spanText(p: Extract<FlashPiece, { kind: "span" }>, tokens: string[][]): string | null {
  const t = tokens[p.col]
  if (!t) return null
  const i = p.from < 0 ? t.length + p.from : p.from
  const j = p.to < 0 ? t.length + p.to : p.to
  if (i < 0 || j >= t.length || i > j) return null
  let text = t.slice(i, j + 1).join("")
  if (p.first !== undefined) {
    if (text.length < p.first) return null
    text = text.slice(0, p.first)
  }
  return changeCase(text, p.case)
}

/** A program's output for one row's sources (null when a piece doesn't apply, e.g. the row has too few words). */
export function runFlashProgram(program: FlashPiece[], sources: string[]): string | null {
  const tokens = sources.map(flashTokens)
  let out = ""
  for (const p of program) {
    const text = p.kind === "text" ? p.text : spanText(p, tokens)
    if (text === null) return null
    out += text
  }
  return out
}

const CASES: FlashCase[] = ["same", "lower", "upper", "proper"]
const TEXT_COST = 2
const NODE_LIMIT = 20_000

type Node = { pos: number[]; cost: number; program: FlashPiece[] }

/**
 * The cheapest program that turns each example's sources into its output, or null. `examples` are rows of
 * [sources, output].
 */
export function learnFlashProgram(examples: { sources: string[]; output: string }[]): FlashPiece[] | null {
  if (!examples.length || examples.some((e) => e.output === "")) return null
  const tokens = examples.map((e) => e.sources.map(flashTokens))
  const outs = examples.map((e) => e.output)
  const done = (pos: number[]) => pos.every((p, k) => p === outs[k].length)
  const seen = new Map<string, number>()
  const queue: Node[] = [{ pos: outs.map(() => 0), cost: 0, program: [] }]
  let expanded = 0
  while (queue.length && expanded++ < NODE_LIMIT) {
    // The cheapest open state (the queue stays small: states are positions in the outputs).
    let best = 0
    for (let i = 1; i < queue.length; i++) if (queue[i].cost < queue[best].cost) best = i
    const node = queue.splice(best, 1)[0]
    if (done(node.pos)) return mergeText(node.program)
    const key = node.pos.join(",")
    if ((seen.get(key) ?? Infinity) < node.cost) continue
    const push = (piece: FlashPiece, cost: number, lengths: number[]) => {
      const pos = node.pos.map((p, k) => p + lengths[k])
      const total = node.cost + cost
      const k = pos.join(",")
      if ((seen.get(k) ?? Infinity) <= total) return
      seen.set(k, total)
      queue.push({ pos, cost: total, program: [...node.program, piece] })
    }
    // Typed text: the same next character in every example.
    const ch = outs[0][node.pos[0]]
    if (ch !== undefined && outs.every((o, k) => o[node.pos[k]] === ch))
      push(
        { kind: "text", text: ch },
        TEXT_COST,
        outs.map(() => 1),
      )
    // Spans of the first example's sources that make its output here, kept when they fit every other example too.
    const p0 = node.pos[0]
    const tried = new Set<string>()
    tokens[0].forEach((t, col) => {
      const n = t.length
      for (let i = 0; i < n; i++) {
        let text = ""
        for (let j = i; j < n; j++) {
          text += t[j]
          for (const c of CASES) {
            const whole = changeCase(text, c)
            const firsts: (number | undefined)[] = [undefined]
            if (i === j) for (let f = 1; f < text.length && f <= 3; f++) firsts.push(f)
            for (const first of firsts) {
              const made = first === undefined ? whole : changeCase(text.slice(0, first), c)
              if (!made || !outs[0].startsWith(made, p0)) continue
              // Counted from the start, or from the end when the span reaches the last word.
              const anchors: [number, number, number][] =
                j === n - 1 && i > 0
                  ? [
                      [i - n, -1, 0],
                      [i, -1, 0.2],
                      [i, j, 0.4],
                    ]
                  : [
                      [i, j, 0],
                      [i, j - n, 0.3],
                      [i - n, j - n, 0.4],
                    ]
              for (const [from, to, extra] of anchors) {
                const piece: FlashPiece = { kind: "span", col, from, to, case: c, ...(first ? { first } : {}) }
                const id = JSON.stringify(piece)
                if (tried.has(id)) continue
                tried.add(id)
                const lengths: number[] = []
                let fits = true
                for (let k = 0; k < outs.length && fits; k++) {
                  const s = k === 0 ? made : spanText(piece, tokens[k])
                  if (s === null || s === "" || !outs[k].startsWith(s, node.pos[k])) fits = false
                  else lengths.push(s.length)
                }
                if (!fits) continue
                const cost = 1 + extra + (c === "same" ? 0 : 0.1) + (first ? 0.5 : 0)
                push(piece, cost, lengths)
              }
            }
          }
        }
      }
    })
  }
  return null
}

function mergeText(program: FlashPiece[]): FlashPiece[] {
  const out: FlashPiece[] = []
  for (const p of program) {
    const last = out[out.length - 1]
    if (p.kind === "text" && last?.kind === "text") last.text += p.text
    else out.push(p.kind === "text" ? { ...p } : p)
  }
  return out
}

/**
 * Flash Fill over rows: `sources` are each row's source texts, `outputs` the column being filled (null: empty, to
 * fill). Returns the texts for the empty rows (null where the pattern doesn't apply), or null when no pattern fits the
 * examples.
 */
export function flashFill(sources: string[][], outputs: (string | null)[]): (string | null)[] | null {
  const examples = outputs.flatMap((o, i) => (o === null ? [] : [{ sources: sources[i], output: o }]))
  const program = learnFlashProgram(examples)
  if (!program) return null
  return outputs.map((o, i) =>
    o !== null || sources[i].every((s) => s === "") ? null : runFlashProgram(program, sources[i]),
  )
}
