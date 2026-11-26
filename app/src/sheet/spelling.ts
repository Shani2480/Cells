import {
  allObjects,
  cellName,
  isChart,
  isSingleCellOn,
  replaceText,
  selectCell,
  textPositions,
  type CellAddress,
  type HeaderFooterText,
  type Op,
  type PageSetup,
  type Range,
  type SheetShape,
} from "@cells/core"
import type { Speller } from "@/proofing/speller"
import { DEFAULT_SPELL_OPTIONS, matchCase, wordsIn, type SpellOptions } from "@/proofing/words"
import type { SheetController } from "./controller"

/** Excel's messages at the end of a check. */
export const SPELLING_MESSAGES = {
  done: "Spell check complete. You're good to go!",
}

/** The header and footer fields of a sheet's page setup, as Spelling goes through them. */
const HEADER_FIELDS = ["header", "footer", "firstHeader", "firstFooter", "evenHeader", "evenFooter"] as const
type HeaderField = (typeof HEADER_FIELDS)[number]

/** Where text Spelling checks lives: a cell, a note, a comment's post, a shape, a header or footer section. */
export type SpellPlace =
  | ({ kind: "cell" } & CellAddress)
  | ({ kind: "note" } & CellAddress)
  | ({ kind: "comment"; post: number } & CellAddress)
  | { kind: "shape"; id: number }
  | { kind: "header"; field: HeaderField; section: keyof HeaderFooterText }

/**
 * A word Spelling stopped at: where it is (with the cell, for a cell, note or comment), where the word is in its
 * text, and the dictionary's suggestions.
 */
export type SpellingIssue = Partial<CellAddress> & {
  place: SpellPlace
  start: number
  end: number
  word: string
  suggestions: string[]
}

const HEADER_NAMES: Record<HeaderField, string> = {
  header: "header",
  footer: "footer",
  firstHeader: "first page header",
  firstFooter: "first page footer",
  evenHeader: "even page header",
  evenFooter: "even page footer",
}

/** Where a word is, for the Spelling dialog when it isn't in a cell ("" for a cell: the cell is selected). */
export function placeLabel(p: SpellPlace, shapeName: (id: number) => string): string {
  const cell = (q: CellAddress) => cellName(q)
  switch (p.kind) {
    case "cell":
      return ""
    case "note":
      return `In the note on ${cell(p)}`
    case "comment":
      return `In the comment on ${cell(p)}`
    case "shape":
      return `In ${shapeName(p.id)}`
    case "header":
      return `In the ${p.section} section of the ${HEADER_NAMES[p.field]}`
  }
}

/** A header's codes (&P, &"Arial,Bold", &12, &K00FF00…) blanked out, so only its words are checked. */
const HEADER_CODE = /&(?:"[^"]*"|K[0-9A-Fa-f]{6}|K\d{2}[+-]\d{3}|\d+|[\s\S])/g
const headerWords = (text: string) => text.replace(HEADER_CODE, (m) => " ".repeat(m.length))

/**
 * Review › Spelling (F7) on the sheet shown, as Excel checks it: the selected shapes' text when shapes are selected;
 * the selection's cells when it's more than one cell; otherwise the whole sheet from the active cell on, round to
 * where it began, then its notes and comments, the text in its shapes and text boxes, and its headers and footers.
 * Formulas aren't checked, only text typed in. Ignore All and Change All last for this check; Add to Dictionary is
 * kept (and shared with Inkspire).
 */
export class SpellingSession {
  private readonly places: SpellPlace[]
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
    const chosen = c.objects.drawings().filter((o): o is SheetShape => o.kind === "shape" && !!o.text)
    if (chosen.length) {
      this.places = chosen.map((o) => ({ kind: "shape", id: o.id }))
      return
    }
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
    const places: SpellPlace[] = list.map((p) => ({ kind: "cell", ...p }))
    if (one) {
      const sheet = c.sheet
      const notes = sheet.annotations.all().sort((a, b) => a.row - b.row || a.col - b.col)
      for (const { row, col, value } of notes) {
        if (value.kind === "note") places.push({ kind: "note", row, col })
        else value.posts.forEach((_, post) => places.push({ kind: "comment", row, col, post }))
      }
      for (const o of allObjects(sheet.objects))
        if (!isChart(o) && o.kind === "shape" && o.text && !o.hidden) places.push({ kind: "shape", id: o.id })
      for (const field of HEADER_FIELDS)
        for (const section of ["left", "center", "right"] as const)
          if (sheet.pageSetup[field][section]) places.push({ kind: "header", field, section })
    }
    this.places = places
  }

  /** The place's text as Spelling reads it (a header's codes blanked, so they're never words). */
  private text(p: SpellPlace, raw = false): string {
    const sheet = this.c.sheet
    switch (p.kind) {
      case "cell": {
        const v = sheet.cells.value(p.row, p.col)
        return typeof v === "string" ? v : ""
      }
      case "note": {
        const a = sheet.annotations.get(p.row, p.col)
        return a?.kind === "note" ? a.text : ""
      }
      case "comment": {
        const a = sheet.annotations.get(p.row, p.col)
        return a?.kind === "comment" ? (a.posts[p.post]?.text ?? "") : ""
      }
      case "shape": {
        const o = this.shape(p.id)
        return o?.text ? textPositions(o.text) : ""
      }
      case "header": {
        const t = sheet.pageSetup[p.field][p.section]
        return raw ? t : headerWords(t)
      }
    }
  }

  /** Where the current word is, for the dialog ("" in a cell). */
  where(): string {
    return this.current ? placeLabel(this.current.place, (id) => this.shape(id)?.name ?? "a shape") : ""
  }

  private shape(id: number): SheetShape | null {
    for (const o of allObjects(this.c.sheet.objects)) if (o.id === id && !isChart(o) && o.kind === "shape") return o
    return null
  }

  /** The place shown while its word is asked about: its cell selected, or the shape. */
  private reveal(p: SpellPlace) {
    if (p.kind === "shape") this.c.objects.select(p.id)
    else if (p.kind !== "header") {
      if (this.c.objects.selected.length) this.c.objects.select(null)
      this.c.setSelection(selectCell({ row: p.row, col: p.col }), { row: p.row, col: p.col })
    }
  }

  /** Another language's dictionary (the Spelling dialog's picker): the check goes on from the current word. */
  useSpeller(speller: Speller): Promise<SpellingIssue | null> {
    this.speller = speller
    if (this.current) this.offset = this.current.start
    return this.next()
  }

  /** Moves to the next word not in the dictionary (null: the check is done). Change All words change on the way. */
  async next(): Promise<SpellingIssue | null> {
    while (this.at < this.places.length) {
      const p = this.places[this.at]
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
        const cell = p.kind === "shape" || p.kind === "header" ? {} : { row: p.row, col: p.col }
        found = {
          ...cell,
          place: p,
          start: w.start,
          end: w.end,
          word: w.word,
          suggestions: await this.speller.suggest(w.word),
        }
        break
      }
      if (found) {
        this.current = found
        this.reveal(p)
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
    this.replace(cur.place, cur.start, cur.end, to)
    this.offset = cur.start + to.length
    return this.next()
  }

  changeEvery(to: string): Promise<SpellingIssue | null> {
    if (this.current) this.changeAll.set(this.current.word.toLowerCase(), to)
    return this.change(to)
  }

  /** A word replaced in its place, as one undo step. */
  private replace(p: SpellPlace, start: number, end: number, to: string) {
    const c = this.c
    const sheet = c.sheet
    const swap = (text: string) => text.slice(0, start) + to + text.slice(end)
    const after = { label: "Spelling", before: c.selection, after: c.selection }
    switch (p.kind) {
      case "cell": {
        const next = swap(this.text(p))
        const keep = sheet.cells.get(p.row, p.col)
        // The cell keeps its format, and a leading apostrophe that kept text as text.
        const data = {
          value: next,
          ...(keep?.input?.startsWith("'") ? { input: "'" + next } : {}),
          ...(keep?.style ? { style: keep.style } : {}),
        }
        const op: Op = { kind: "setCells", sheet: c.sheetIndex, cells: [{ row: p.row, col: p.col, data }] }
        c.transact([op], after)
        return
      }
      case "note":
      case "comment": {
        const a = sheet.annotations.get(p.row, p.col)
        if (!a) return
        const value =
          a.kind === "note"
            ? { ...a, text: swap(a.text) }
            : {
                ...a,
                posts: a.posts.map((x, i) => (i === (p as { post: number }).post ? { ...x, text: swap(x.text) } : x)),
              }
        c.transact([{ kind: "annotations", sheet: c.sheetIndex, writes: [{ row: p.row, col: p.col, value }] }], after)
        return
      }
      case "shape":
        c.objects.update(
          "Spelling",
          (o) => {
            const s = o as SheetShape
            if (s.text) s.text = replaceText(s.text, start, end, to)
          },
          [p.id],
        )
        return
      case "header": {
        const setup: PageSetup = { ...sheet.pageSetup }
        setup[p.field] = { ...setup[p.field], [p.section]: swap(this.text(p, true)) }
        c.transact([{ kind: "pageSetup", sheet: c.sheetIndex, setup }], after)
        return
      }
    }
  }
}
