import {
  bookNameOf,
  copyExternalBooks,
  linkKey,
  MAX_COLS,
  MAX_ROWS,
  parseRefText,
  quoteSheet,
  splitBookSheet,
  tokenize,
  type CellValue,
  type CellWrite,
  type ExternalBook,
  type Op,
  type Range,
  type Workbook,
} from "@cells/core"

/**
 * Data › Workbook Links: which other workbooks this one's formulas use, and the steps that change them — Update
 * Values (read the file again), Change Source (point the links at another file) and Break Link (formulas become their
 * values). Each is one step to undo.
 */

type Use = { name: string; sheets: Map<string, { name: string; areas: Range[] }>; names: Set<string>; cells: number }

/** A ref or name token's prefix and what follows its "!". */
function split(text: string): { prefix: string; rest: string } | null {
  if (text.startsWith("'")) {
    let i = 1
    for (; i < text.length; i++) {
      if (text[i] !== "'") continue
      if (text[i + 1] === "'") i++
      else break
    }
    return text[i + 1] === "!" ? { prefix: text.slice(1, i).replace(/''/g, "'"), rest: text.slice(i + 2) } : null
  }
  const bang = text.indexOf("!")
  return bang < 0 ? null : { prefix: text.slice(0, bang), rest: text.slice(bang + 1) }
}

/** The workbook a token links to ("[Book]Sheet" or a known book's name alone), with its sheet. */
function linkOf(
  token: { kind: string; text: string },
  books: readonly ExternalBook[],
  sheets: Set<string>,
): { book: string; sheet: string | null; rest: string } | null {
  if (token.kind !== "ref" && token.kind !== "name") return null
  const p = split(token.text)
  if (!p) return null
  const parts = splitBookSheet(p.prefix)
  if (parts) return { book: parts.book, sheet: parts.sheet, rest: p.rest }
  const upper = p.prefix.toUpperCase()
  if (token.kind === "name" && !sheets.has(upper) && books.some((b) => b.name.toUpperCase() === upper))
    return { book: books.find((b) => b.name.toUpperCase() === upper)!.name, sheet: null, rest: p.rest }
  return null
}

/** Every formula's (and name's) text in the workbook. */
function eachFormula(
  wb: Workbook,
  visit: (text: string, at: { sheet: number; row: number; col: number } | null) => void,
) {
  wb.sheets.forEach((sheet, index) =>
    sheet.cells.forEachFormula((row, col, input) => {
      if (input.includes("!")) visit(input, { sheet: index, row, col })
    }),
  )
  for (const n of wb.names.all()) if (n.ref.includes("!")) visit("=" + n.ref, null)
}

/** The workbooks formulas link to, by upper-case name: sheets and areas used, names used, formula cells. */
export function linkUses(wb: Workbook): Map<string, Use> {
  const uses = new Map<string, Use>()
  const books = wb.links.list
  const sheets = new Set(wb.sheets.map((s) => s.name.toUpperCase()))
  eachFormula(wb, (text, at) => {
    const seen = new Set<string>()
    for (const t of tokenize(text)) {
      const link = linkOf(t, books, sheets)
      if (!link) continue
      const key = link.book.toUpperCase()
      let use = uses.get(key)
      if (!use) uses.set(key, (use = { name: link.book, sheets: new Map(), names: new Set(), cells: 0 }))
      if (at && !seen.has(key)) {
        use.cells++
        seen.add(key)
      }
      if (t.kind === "name" || link.sheet === null) {
        use.names.add(link.rest)
        continue
      }
      const ref = parseRefText("X!" + link.rest)
      const sheetKey = link.sheet.toUpperCase()
      let s = use.sheets.get(sheetKey)
      if (!s) use.sheets.set(sheetKey, (s = { name: link.sheet, areas: [] }))
      if (!ref) continue
      const to = ref.to ?? ref.from
      s.areas.push({
        r0: Math.min(ref.from.row ?? 0, to.row ?? MAX_ROWS - 1),
        r1: Math.max(ref.from.row ?? 0, to.row ?? MAX_ROWS - 1),
        c0: Math.min(ref.from.col ?? 0, to.col ?? MAX_COLS - 1),
        c1: Math.max(ref.from.col ?? 0, to.col ?? MAX_COLS - 1),
      })
    }
  })
  return uses
}

/** A row of Workbook Links: a link Cells keeps, or a workbook formulas name whose values were never read. */
export type LinkRow = { name: string; location: string; known: boolean; cells: number }

export function linkRows(wb: Workbook): LinkRow[] {
  const uses = linkUses(wb)
  const rows: LinkRow[] = wb.links.list.map((b) => ({
    name: b.name,
    location: b.absolute ?? b.target,
    known: true,
    cells: uses.get(b.name.toUpperCase())?.cells ?? 0,
  }))
  for (const [key, use] of uses)
    if (!wb.links.list.some((b) => b.name.toUpperCase() === key))
      rows.push({ name: use.name, location: use.name, known: false, cells: use.cells })
  return rows
}

/** Where to look for a link's file: its stored path against this workbook's folder, then the absolute path. */
export function linkPaths(book: { target: string; absolute?: string } | undefined, own: string | undefined) {
  const out: string[] = []
  if (!book) return out
  const absolute = (p: string) => /^(?:[A-Za-z]:[\\/]|[\\/]|content:)/.test(p)
  if (absolute(book.target)) out.push(book.target)
  else if (own && !own.startsWith("content:")) {
    const sep = own.includes("\\") && !own.includes("/") ? "\\" : "/"
    const dir = own.slice(0, Math.max(own.lastIndexOf("/"), own.lastIndexOf("\\")) + 1)
    out.push(dir + book.target.replace(/[\\/]/g, sep))
  }
  if (book.absolute && !out.includes(book.absolute)) out.push(book.absolute)
  return out
}

/** A source workbook as read: its sheets' values and its names. */
export type SourceBook = Pick<Workbook, "sheets" | "names">

/** The link with the values this workbook's formulas read from the source, as now in the file. */
export function readLink(
  wb: Workbook,
  name: string,
  source: SourceBook,
  location: { target: string; absolute?: string },
) {
  const use = linkUses(wb).get(name.toUpperCase())
  const sheetOf = (n: string) => source.sheets.findIndex((s) => s.name.toUpperCase() === n.toUpperCase())
  const book: ExternalBook = {
    name: bookNameOf(location.absolute ?? location.target),
    target: location.target,
    ...(location.absolute ? { absolute: location.absolute } : {}),
    sheets: [],
    names: [],
  }
  const want = new Map<number, Range[]>()
  const add = (sheet: number, area: Range) => {
    let list = want.get(sheet)
    if (!list) want.set(sheet, (list = []))
    list.push(area)
  }
  for (const s of use?.sheets.values() ?? []) {
    const i = sheetOf(s.name)
    if (i >= 0) for (const a of s.areas) add(i, a)
  }
  for (const n of use?.names ?? []) {
    const def = source.names.all().find((d) => d.name.toUpperCase() === n.toUpperCase())
    if (!def) continue
    book.names.push({ name: def.name, refersTo: "=" + def.ref })
    for (const t of tokenize("=" + def.ref)) {
      if (t.kind !== "ref") continue
      const ref = parseRefText(t.text)
      if (!ref?.sheet) continue
      const i = sheetOf(ref.sheet)
      const to = ref.to ?? ref.from
      if (i >= 0)
        add(i, {
          r0: ref.from.row ?? 0,
          r1: to.row ?? MAX_ROWS - 1,
          c0: ref.from.col ?? 0,
          c1: to.col ?? MAX_COLS - 1,
        })
    }
  }
  // Every sheet keeps its place (a name's refersTo and the file's sheet ids count them); values only where used.
  book.sheets = source.sheets.map((sheet, i) => {
    const cells = new Map<number, CellValue>()
    for (const area of want.get(i) ?? []) {
      const used = sheet.cells.usedBounds()
      if (!used) break
      const r: Range = {
        r0: Math.max(area.r0, used.r0),
        r1: Math.min(area.r1, used.r1),
        c0: Math.max(area.c0, used.c0),
        c1: Math.min(area.c1, used.c1),
      }
      if (r.r0 > r.r1 || r.c0 > r.c1) continue
      sheet.cells.forEach(r, (row, col, data) => {
        if (data.value !== undefined) cells.set(linkKey(row, col), data.value)
      })
    }
    return { name: sheet.name, cells }
  })
  return book
}

/** Update Values (and Change Source): the link replaced, formulas renamed to the new file when it changed. */
export function updateLinkOps(wb: Workbook, old: string, book: ExternalBook): Op[] {
  const list = copyExternalBooks(wb.links.list)
  const at = list.findIndex((b) => b.name.toUpperCase() === old.toUpperCase())
  // A new file's name may already be another link's: the two become one.
  const clash = list.findIndex((b, i) => i !== at && b.name.toUpperCase() === book.name.toUpperCase())
  if (clash >= 0) list.splice(clash, 1)
  const index = list.findIndex((b) => b.name.toUpperCase() === old.toUpperCase())
  if (index >= 0) list[index] = book
  else list.push(book)
  const ops: Op[] = []
  if (old.toUpperCase() !== book.name.toUpperCase()) ops.push(...renameOps(wb, old, book.name))
  ops.push({ kind: "externalLinks", links: list })
  return ops
}

/** Formulas and names that name `from` name `to` instead. */
function renameOps(wb: Workbook, from: string, to: string): Op[] {
  const books = wb.links.list
  const sheets = new Set(wb.sheets.map((s) => s.name.toUpperCase()))
  const rewrite = (text: string) => {
    let changed = false
    const out = tokenize(text).map((t) => {
      const link = linkOf(t, books, sheets)
      if (!link || link.book.toUpperCase() !== from.toUpperCase()) return t.text
      changed = true
      if (link.sheet === null) return quoteSheet(to) + "!" + link.rest
      const plain = quoteSheet(link.sheet) === link.sheet && /^[\w.]+$/.test(to)
      return (plain ? `[${to}]${link.sheet}` : `'[${to}]${link.sheet.replace(/'/g, "''")}'`) + "!" + link.rest
    })
    return changed ? out.join("") : null
  }
  const ops: Op[] = []
  wb.sheets.forEach((sheet, index) => {
    const cells: CellWrite[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      const next = rewrite(input)
      if (next) cells.push({ row, col, data: { ...sheet.cells.get(row, col)!, input: next } })
    })
    if (cells.length) ops.push({ kind: "setCells", sheet: index, cells })
  })
  const names = wb.names.all().map((n) => {
    const next = rewrite("=" + n.ref)
    return next ? { ...n, ref: next.slice(1) } : n
  })
  if (names.some((n, i) => n !== wb.names.all()[i])) ops.push({ kind: "names", names })
  return ops
}

/**
 * Break Link: every formula that uses the workbook becomes the value it shows (its format stays), and the link goes.
 * Excel's warning comes first, in the shell.
 */
export function breakLinkOps(wb: Workbook, name: string): Op[] {
  const books = wb.links.list
  const sheets = new Set(wb.sheets.map((s) => s.name.toUpperCase()))
  const uses = (text: string) =>
    tokenize(text).some((t) => linkOf(t, books, sheets)?.book.toUpperCase() === name.toUpperCase())
  const ops: Op[] = []
  wb.sheets.forEach((sheet, index) => {
    const cells: CellWrite[] = []
    sheet.cells.forEachFormula((row, col, input) => {
      if (!uses(input)) return
      const data = sheet.cells.get(row, col)!
      const { input: _input, ...rest } = data
      void _input
      cells.push({ row, col, data: rest.value === undefined ? undefined : rest })
    })
    if (cells.length) ops.push({ kind: "setCells", sheet: index, cells })
  })
  ops.push({ kind: "externalLinks", links: books.filter((b) => b.name.toUpperCase() !== name.toUpperCase()) })
  return ops
}

const rowCache = new WeakMap<Workbook, { version: number; rows: LinkRow[] }>()

/** linkRows, worked out once per change of the workbook (the ribbon asks on every redraw). */
export function cachedLinkRows(wb: Workbook): LinkRow[] {
  const hit = rowCache.get(wb)
  if (hit && hit.version === wb.version) return hit.rows
  const rows = linkRows(wb)
  rowCache.set(wb, { version: wb.version, rows })
  return rows
}
