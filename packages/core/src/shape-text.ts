import { copyText, type ShapeColor, type ShapeText, type TextParagraph, type TextRun } from "./shapes"

/**
 * Formatting part of a shape's text, as Excel does when text in a shape or text box is selected (Home › Font and
 * the mini toolbar): runs split where the selection starts and ends, and the look set on what's between. Positions
 * count characters of the plain text, a paragraph break counting as one ("\n"), as `plainText` gives it.
 */

/** How a run looks: everything but its text. */
export type TextLook = Omit<TextRun, "text">

/** A change of look: a value sets it, null takes it away (back to the shape's default). */
export type TextLookPatch = { [K in keyof TextLook]?: TextLook[K] | null }

/** The look's fields, in one place for comparing and copying. */
const LOOK_KEYS = ["bold", "italic", "underline", "strike", "size", "font", "color"] as const

export function lookOf(run: TextRun | undefined): TextLook {
  const look: TextLook = {}
  if (!run) return look
  for (const k of LOOK_KEYS) if (run[k] !== undefined) (look as Record<string, unknown>)[k] = run[k]
  return look
}

export function sameTextLook(a: TextLook, b: TextLook): boolean {
  return LOOK_KEYS.every((k) =>
    k === "color" ? JSON.stringify(a.color ?? null) === JSON.stringify(b.color ?? null) : a[k] === b[k],
  )
}

function patched(run: TextRun, patch: TextLookPatch): TextRun {
  const out: TextRun = { ...run }
  for (const k of LOOK_KEYS) {
    if (!(k in patch)) continue
    const v = patch[k]
    if (v === null || v === undefined || v === false) delete out[k]
    else (out as Record<string, unknown>)[k] = k === "color" ? { ...(v as ShapeColor) } : v
  }
  return out
}

/** Runs next to each other with the same look as one run; empty runs dropped (a paragraph keeps one for its look). */
function tidy(runs: TextRun[]): TextRun[] {
  const out: TextRun[] = []
  for (const r of runs) {
    const last = out[out.length - 1]
    if (!r.text && runs.some((x) => x.text)) continue
    if (last && sameTextLook(lookOf(last), lookOf(r))) last.text += r.text
    else out.push({ ...r })
  }
  return out
}

/** Each paragraph's start in the plain text (equations count as their linear text, which isn't formatted here). */
function paragraphStarts(text: ShapeText): number[] {
  const starts: number[] = []
  let at = 0
  for (const p of text.paragraphs) {
    starts.push(at)
    at += paragraphLength(p) + 1
  }
  return starts
}

function paragraphLength(p: TextParagraph): number {
  return p.math ? 0 : p.runs.reduce((n, r) => n + r.text.length, 0)
}

/** The plain text's length. */
export function textLength(text: ShapeText): number {
  return text.paragraphs.reduce((n, p) => n + paragraphLength(p), 0) + Math.max(0, text.paragraphs.length - 1)
}

/**
 * The text with `patch` applied from `from` to `to`. An empty range changes nothing (the caret's look is the
 * editor's business); the whole text is 0 to `textLength`.
 */
export function formatText(text: ShapeText, from: number, to: number, patch: TextLookPatch): ShapeText {
  const out = copyText(text)
  const starts = paragraphStarts(out)
  out.paragraphs.forEach((p, i) => {
    if (p.math) return
    const len = paragraphLength(p)
    const a = Math.max(0, from - starts[i])
    const b = Math.min(len, to - starts[i])
    // An empty paragraph inside the range takes the look too (what's typed there later has it).
    if (len === 0) {
      if (from < to && from <= starts[i] && to >= starts[i]) p.runs = [patched(p.runs[0] ?? { text: "" }, patch)]
      return
    }
    if (a >= b) return
    const runs: TextRun[] = []
    let at = 0
    for (const r of p.runs) {
      const s = at
      const e = at + r.text.length
      at = e
      if (e <= a || s >= b) {
        runs.push(r)
        continue
      }
      const cutA = Math.max(a, s) - s
      const cutB = Math.min(b, e) - s
      if (cutA > 0) runs.push({ ...r, text: r.text.slice(0, cutA) })
      runs.push(patched({ ...r, text: r.text.slice(cutA, cutB) }, patch))
      if (cutB < r.text.length) runs.push({ ...r, text: r.text.slice(cutB) })
    }
    p.runs = tidy(runs)
  })
  return out
}

/**
 * The look the ribbon shows for a range: the first character's (Excel shows the start of the selection), or for a
 * caret the character before it (the paragraph's own look at its start).
 */
export function lookAt(text: ShapeText, from: number, to = from): TextLook {
  const starts = paragraphStarts(text)
  let i = starts.length - 1
  while (i > 0 && starts[i] > from) i--
  const p = text.paragraphs[i]
  if (!p) return {}
  const off = from - starts[i]
  // A selection looks at its first character; a caret at the one before it.
  const pos = to > from ? off + 1 : off
  let at = 0
  for (const r of p.runs) {
    at += r.text.length
    if (pos <= at && r.text) return lookOf(r)
  }
  return lookOf(p.runs[0])
}

/** A Home alignment for the paragraphs the range touches. */
export function alignText(text: ShapeText, from: number, to: number, align: TextParagraph["align"] | null): ShapeText {
  const out = copyText(text)
  const starts = paragraphStarts(out)
  out.paragraphs.forEach((p, i) => {
    const end = starts[i] + paragraphLength(p)
    if (end < from || starts[i] > to) return
    if (align) p.align = align
    else delete p.align
  })
  return out
}

/** The alignment of the paragraph a position is in. */
export function alignAt(text: ShapeText, at: number): TextParagraph["align"] {
  const starts = paragraphStarts(text)
  let i = starts.length - 1
  while (i > 0 && starts[i] > at) i--
  return text.paragraphs[i]?.align
}

/** Runs tidied in every paragraph (after an editor rebuilt the text). */
export function tidyText(text: ShapeText): ShapeText {
  return { ...text, paragraphs: text.paragraphs.map((p) => (p.math ? p : { ...p, runs: tidy(p.runs) })) }
}
