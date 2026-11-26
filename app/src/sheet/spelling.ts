import { isSingleCellOn, selectCell, type CellAddress, type Op, type Range } from "@cells/core"
import type { Speller } from "@/proofing/speller"
import { DEFAULT_SPELL_OPTIONS, matchCase, wordsIn, type SpellOptions } from "@/proofing/words"
import type { SheetController } from "./controller"

/** Excel's messages at the end of a check. */
export const SPELLING_MESSAGES = {
  done: "Spell check complete. You're good to go!",
}

/** A word Spelling stopped at: the cell, where the word is in its text, and the dictionary's suggestions. */
export type SpellingIssue = CellAddress & { start: number; end: number; word: string; suggestions: string[] }

/**
 * Review › Spelling (F7) on the sheet shown: the selection when it's more than one cell, otherwise the whole sheet
 * from the active cell on, round to where it began. Formulas aren't checked, only text typed into cells. Ignore All
 * and Change All last for this check; Add to Dictionary is kept (and shared with Inkspire).
 */
export class SpellingSession {
  private readonly cells: CellAddress[]
  private at = 0
  /** Where in the current cell's text to look next. */
  private offset = 0
  private readonly ignored = new Set<string>()
  private readonly changeAll = new Map<string, string>()
  current: SpellingIssue | null = null

  constructor(
    private readonly c: SheetController,
    private speller: Speller,
    private readonly onAdd: (word: string) => void = () => undefined,
    private readonly options: SpellOptions = DEFAULT_SPELL_OPTIONS,
  ) {
    const sel = c.selection
    // One cell (a merged cell counts as one) checks the whole sheet.
    const one = isSingleCellOn(c.sheet, sel)
    const used = c.sheet.cells.usedBounds()
    const ranges: readonly Range[] = one ? (used ? [used] : []) : sel.ranges
    const list: CellAddress[] = []
    for (const r of ranges) {
      const area = used
        ? {
            r0: Math.max(r.r0, used.r0),
            r1: Math.min(r.r1, used.r1),
            c0: Math.max(r.c0, used.c0),
            c1: Math.min(r.c1, used.c1),
          }
        : r
      if (area.r0 > area.r1 || area.c0 > area.c1) continue
      c.sheet.cells.forEach(
        area,
        (row, col, data) => {
          if (typeof data.value === "string" && !data.input?.startsWith("=")) list.push({ row, col })
        },
        true,
      )
    }
    list.sort((a, b) => a.row - b.row || a.col - b.col)
    // The whole sheet starts at the active cell and comes round to it (Excel goes on from the top).
    if (one) {
      const { row, col } = sel.active
      const from = list.findIndex((p) => p.row > row || (p.row === row && p.col >= col))
      if (from > 0) list.push(...list.splice(0, from))
    }
    this.cells = list
  }

  private text(p: CellAddress): string {
    const v = this.c.sheet.cells.value(p.row, p.col)
    return typeof v === "string" ? v : ""
  }

  /** Another language's dictionary (the Spelling dialog's picker): the check goes on from the current word. */
  useSpeller(speller: Speller): Promise<SpellingIssue | null> {
    this.speller = speller
    if (this.current) this.offset = this.current.start
    return this.next()
  }

  /** Moves to the next word not in the dictionary (null: the check is done). Change All words change on the way. */
  async next(): Promise<SpellingIssue | null> {
    while (this.at < this.cells.length) {
      const p = this.cells[this.at]
      const words = wordsIn(this.text(p), this.options).filter(
        (w) => w.start >= this.offset && !this.ignored.has(w.word.toLowerCase()),
      )
      const known = words.length ? await this.speller.check(words.map((w) => w.word)) : []
      let found: SpellingIssue | null = null
      for (const [i, w] of words.entries()) {
        if (known[i]) continue
        const to = this.changeAll.get(w.word.toLowerCase())
        if (to !== undefined) {
          this.replace(p, w.start, w.end, matchCase(w.word, to))
          this.offset = w.start + to.length
          return this.next()
        }
        found = { ...p, start: w.start, end: w.end, word: w.word, suggestions: await this.speller.suggest(w.word) }
        break
      }
      if (found) {
        this.current = found
        this.c.setSelection(selectCell({ row: found.row, col: found.col }), { row: found.row, col: found.col })
        return found
      }
      this.at++
      this.offset = 0
    }
    this.current = null
    return null
  }

  ignoreOnce(): Promise<SpellingIssue | null> {
    if (this.current) this.offset = this.current.end
    return this.next()
  }

  ignoreAll(): Promise<SpellingIssue | null> {
    if (this.current) this.ignored.add(this.current.word.toLowerCase())
    return this.ignoreOnce()
  }

  addToDictionary(): Promise<SpellingIssue | null> {
    if (this.current) {
      this.speller.add(this.current.word)
      this.onAdd(this.current.word)
    }
    return this.ignoreOnce()
  }

  change(to: string): Promise<SpellingIssue | null> {
    const cur = this.current
    if (!cur) return this.next()
    this.replace(cur, cur.start, cur.end, to)
    this.offset = cur.start + to.length
    return this.next()
  }

  changeEvery(to: string): Promise<SpellingIssue | null> {
    if (this.current) this.changeAll.set(this.current.word.toLowerCase(), to)
    return this.change(to)
  }

  private replace(p: CellAddress, start: number, end: number, to: string) {
    const text = this.text(p)
    const next = text.slice(0, start) + to + text.slice(end)
    const keep = this.c.sheet.cells.get(p.row, p.col)
    // The cell keeps its format, and a leading apostrophe that kept text as text.
    const data = {
      value: next,
      ...(keep?.input?.startsWith("'") ? { input: "'" + next } : {}),
      ...(keep?.style ? { style: keep.style } : {}),
    }
    const op: Op = { kind: "setCells", sheet: this.c.sheetIndex, cells: [{ row: p.row, col: p.col, data }] }
    this.c.transact([op], { label: "Spelling", before: this.c.selection, after: this.c.selection })
  }
}
