import {
  columnName,
  NOTE_HEIGHT,
  NOTE_WIDTH,
  type Annotation,
  type CellComment,
  type CellNote,
  type Sheet,
} from "@cells/core"
import { escapeXml, escapeXmlText } from "@cells/ooxml"
import { NS_MAIN } from "./constants"
import { at, kid, kids, parseTree, textOf } from "./dom"

/**
 * Notes and comments. A note is a legacy comment (comments part) drawn by a VML shape; a 365 comment is a thread in
 * the threadedComments part, its authors in the workbook's persons part, plus a legacy comment carrying a
 * "[Threaded comment]" placeholder so older Excel versions show something.
 */

export const NS_THREADED = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

const PLACEHOLDER =
  "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924"

/** "A1" → {row, col}. */
function cellOf(ref: string | undefined): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref ?? "")
  if (!m) return null
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(m[2]) - 1, col: col - 1 }
}

const refOf = (row: number, col: number) => columnName(col) + (row + 1)

// ── Reading ──────────────────────────────────────────────────────────────────────────────────────────────────

export type CommentParts = {
  legacy?: string
  vml?: string
  threaded?: string
  /** Person id → display name (the workbook's persons part). */
  persons: ReadonlyMap<string, string>
}

/** Person id → display name, from the persons part. */
export function readPersons(xml: string): Map<string, string> {
  const out = new Map<string, string>()
  for (const p of kids(parseTree(xml).root, "person")) out.set(at(p, "id") ?? "", at(p, "displayName") ?? "")
  return out
}

/** A threaded comment's time as written (local time, no zone); kept as if it were UTC. */
function timeOf(text: string | undefined): number {
  if (!text) return 0
  const ms = Date.parse(/Z|[+-]\d\d:\d\d$/.test(text) ? text : text + "Z")
  return Number.isNaN(ms) ? 0 : ms
}

const timeText = (ms: number) =>
  new Date(ms)
    .toISOString()
    .replace(/\.(\d\d)\dZ$/, ".$1")
    .replace(/Z$/, "")

/** Shapes of the VML part by cell: shown (Show Note) and the box size in px. */
function vmlShapes(vml: string): Map<string, { shown: boolean; width?: number; height?: number }> {
  const out = new Map<string, { shown: boolean; width?: number; height?: number }>()
  for (const m of vml.matchAll(/<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/g)) {
    const body = m[2]
    if (!/ObjectType="Note"/.test(body)) continue
    const row = /<x:Row>\s*(\d+)\s*<\/x:Row>/.exec(body)
    const col = /<x:Column>\s*(\d+)\s*<\/x:Column>/.exec(body)
    if (!row || !col) continue
    const style = /style="([^"]*)"/.exec(m[1])?.[1] ?? ""
    const size = (name: string) => {
      const v = new RegExp(`(?:^|;)\\s*${name}:\\s*([\\d.]+)(pt|px)?`).exec(style)
      if (!v) return undefined
      return Math.round(v[2] === "px" ? Number(v[1]) : (Number(v[1]) * 4) / 3)
    }
    out.set(`${row[1]},${col[1]}`, {
      shown: /<x:Visible\s*\/>/.test(body),
      width: size("width"),
      height: size("height"),
    })
  }
  return out
}

/** Reads a sheet's notes and comments into `sheet.annotations`. */
export function readComments(sheet: Sheet, parts: CommentParts) {
  const threadedCells = new Set<string>()
  if (parts.threaded) {
    const root = parseTree(parts.threaded).root
    const threads = new Map<string, { row: number; col: number; comment: CellComment }>()
    for (const t of kids(root, "threadedComment")) {
      const cell = cellOf(at(t, "ref"))
      if (!cell) continue
      const post = {
        author: parts.persons.get(at(t, "personId") ?? "") ?? "",
        text: textOf(kid(t, "text") ?? t, "text").replace(/\r\n/g, "\n"),
        time: timeOf(at(t, "dT")),
      }
      const parent = at(t, "parentId")
      if (parent) {
        threads.get(parent)?.comment.posts.push(post)
        continue
      }
      const comment: CellComment = { kind: "comment", resolved: at(t, "done") === "1", posts: [post] }
      threads.set(at(t, "id") ?? "", { ...cell, comment })
    }
    for (const { row, col, comment } of threads.values()) {
      sheet.annotations.set(row, col, comment)
      threadedCells.add(`${row},${col}`)
    }
  }
  if (!parts.legacy) return
  const shapes = parts.vml ? vmlShapes(parts.vml) : new Map()
  const root = parseTree(parts.legacy).root
  const authors = kids(kid(root, "authors"), "author").map((a) => a.text)
  for (const c of kids(kid(root, "commentList"), "comment")) {
    const cell = cellOf(at(c, "ref"))
    if (!cell || threadedCells.has(`${cell.row},${cell.col}`)) continue
    const author = authors[Number(at(c, "authorId") ?? 0)] ?? ""
    // A thread's placeholder whose thread part is missing: nothing to show.
    if (author.startsWith("tc=")) continue
    let text = textOf(kid(c, "text") ?? c).replace(/\r\n/g, "\n")
    // Excel starts a note with its author's name in bold: "Name:\n".
    if (author && text.startsWith(author + ":")) text = text.slice(author.length + 1).replace(/^\n/, "")
    const shape = shapes.get(`${cell.row},${cell.col}`)
    const note: CellNote = { kind: "note", author, text, shown: shape?.shown ?? false }
    if (shape?.width && shape.width !== NOTE_WIDTH) note.width = shape.width
    if (shape?.height && shape.height !== NOTE_HEIGHT) note.height = shape.height
    sheet.annotations.set(cell.row, cell.col, note)
  }
}

// ── Writing ──────────────────────────────────────────────────────────────────────────────────────────────────

/** A stable GUID-shaped id from text (the same person or post gets the same id on every save). */
export function guid(seed: string): string {
  let h1 = 0x811c9dc5
  let h2 = 0x01000193
  let h3 = 0x9e3779b9
  let h4 = 0x85ebca6b
  for (let i = 0; i < seed.length; i++) {
    const c = seed.charCodeAt(i)
    h1 = Math.imul(h1 ^ c, 16777619)
    h2 = Math.imul(h2 ^ c, 2246822519)
    h3 = Math.imul(h3 ^ c, 3266489917)
    h4 = Math.imul(h4 ^ c, 668265263)
  }
  const hex = [h1, h2, h3, h4].map((h) => (h >>> 0).toString(16).padStart(8, "0").toUpperCase()).join("")
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}}`
}

export const personId = (name: string) => guid("person:" + name)

export type CommentsOut = {
  /** Legacy comments part (notes, and placeholders for threads). */
  legacy: string
  vml: string
  threaded: string | null
  /** Authors of the threads (for the workbook's persons part). */
  persons: Set<string>
}

/** The parts for a sheet's annotations; null when it has none. `sheetIndex` keeps the VML shape ids unique. */
export function writeComments(
  sheet: Sheet,
  sheetIndex: number,
  colPx: (col: number) => number,
  rowPx: (row: number) => number,
): CommentsOut | null {
  const all = sheet.annotations.all()
  if (!all.length) return null
  const authors: string[] = []
  const authorId = (name: string) => {
    let i = authors.indexOf(name)
    if (i < 0) i = authors.push(name) - 1
    return i
  }
  const persons = new Set<string>()
  let list = ""
  let threads = ""
  let shapes = ""
  let shapeId = 1024 * (sheetIndex + 1) + 1
  for (const { row, col, value } of all) {
    const ref = refOf(row, col)
    if (value.kind === "note") {
      const id = authorId(value.author)
      const runs = value.author
        ? `<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t>${escapeXmlText(value.author)}:</t></r>` +
          `<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">\n${escapeXmlText(value.text)}</t></r>`
        : `<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${escapeXmlText(value.text)}</t></r>`
      list += `<comment ref="${ref}" authorId="${id}" shapeId="0"><text>${runs}</text></comment>`
      shapes += noteShape(shapeId++, row, col, value, colPx, rowPx)
    } else {
      const posts = value.posts
      if (!posts.length) continue
      const rootId = guid(`thread:${sheetIndex}:${ref}:${posts[0].time}:${posts[0].author}`)
      let placeholder = PLACEHOLDER + "\n\nComment:\n    " + posts[0].text
      posts.forEach((p, i) => {
        persons.add(p.author)
        const id = i === 0 ? rootId : guid(`reply:${rootId}:${i}:${p.time}`)
        threads +=
          `<threadedComment ref="${ref}" dT="${timeText(p.time)}" personId="${personId(p.author)}" id="${id}"` +
          (i === 0 ? (value.resolved ? ` done="1"` : "") : ` parentId="${rootId}"`) +
          `><text>${escapeXmlText(p.text)}</text></threadedComment>`
        if (i > 0) placeholder += "\nReply:\n    " + p.text
      })
      const id = authorId("tc=" + rootId)
      list += `<comment ref="${ref}" authorId="${id}" shapeId="0"><text><t xml:space="preserve">${escapeXmlText(placeholder)}</t></text></comment>`
      shapes += noteShape(shapeId++, row, col, { shown: false }, colPx, rowPx)
    }
  }
  const legacy =
    `<comments xmlns="${NS_MAIN}"><authors>` +
    authors.map((a) => `<author>${escapeXmlText(a)}</author>`).join("") +
    `</authors><commentList>${list}</commentList></comments>`
  const vml =
    `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheetIndex + 1}"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">` +
    `<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    shapes +
    `</xml>`
  const threaded = threads
    ? `<ThreadedComments xmlns="${NS_THREADED}" xmlns:x="${NS_MAIN}">${threads}</ThreadedComments>`
    : null
  return { legacy, vml, threaded, persons }
}

/** The VML box of a note: placed right of the cell, one row up, as Excel places a new note. */
function noteShape(
  id: number,
  row: number,
  col: number,
  note: { shown: boolean; width?: number; height?: number },
  colPx: (col: number) => number,
  rowPx: (row: number) => number,
): string {
  const width = note.width ?? NOTE_WIDTH
  const height = note.height ?? NOTE_HEIGHT
  // Anchor: left column, x offset, top row, y offset, right column, x offset, bottom row, y offset (px).
  const c0 = col + 1
  const r0 = Math.max(0, row - 1)
  let c1 = c0
  let x = 15 + width
  while (x > colPx(c1) && c1 < 16383) x -= colPx(c1++)
  let r1 = r0
  let y = (row === 0 ? 2 : 10) + height
  while (y > rowPx(r1) && r1 < 1048575) y -= rowPx(r1++)
  const anchor = `${c0}, 15, ${r0}, ${row === 0 ? 2 : 10}, ${c1}, ${Math.round(x)}, ${r1}, ${Math.round(y)}`
  const pt = (px: number) => Math.round(px * 0.75 * 100) / 100
  return (
    `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;margin-left:0;margin-top:0;` +
    `width:${pt(width)}pt;height:${pt(height)}pt;z-index:${id};visibility:${note.shown ? "visible" : "hidden"}" ` +
    `fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
    `<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>` +
    `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor>` +
    `<x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column>${note.shown ? "<x:Visible/>" : ""}` +
    `</x:ClientData></v:shape>`
  )
}

/** The workbook's persons part for the authors of threaded comments. */
export function personsXml(names: Iterable<string>): string {
  let out = ""
  for (const name of names)
    out += `<person displayName="${escapeXml(name)}" id="${personId(name)}" userId="${escapeXml(name)}" providerId="None"/>`
  return `<personList xmlns="${NS_THREADED}" xmlns:x="${NS_MAIN}">${out}</personList>`
}

export type { Annotation }
