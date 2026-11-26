import {
  columnIndex,
  columnName,
  defaultPageSetup,
  MAX_COLS,
  MAX_ROWS,
  quoteSheet,
  type HeaderFooterKey,
  type HeaderFooterText,
  type PageSetup,
  type Range,
} from "@cells/core"
import { escapeXml, escapeXmlText } from "@cells/ooxml"
import { at, kid, kids, type XNode } from "./dom"

/**
 * A sheet's print settings in SpreadsheetML: printOptions, pageMargins, pageSetup, headerFooter, rowBreaks and
 * colBreaks in the worksheet, fitToPage in sheetPr, and the Print_Area and Print_Titles names in the workbook.
 * What Cells doesn't model (the printer settings part, dpi, copies) goes back as read.
 */

/** What a sheet's print elements held that Cells keeps without using. */
export type PageSetupExtras = {
  /** pageSetup's other attributes (r:id of the printer settings, dpi, copies…), as written. */
  pageSetup: string
  /** headerFooter's attributes and children Cells doesn't know, as read. */
  headerFooterAttrs: string
  headerFooterKids: string[]
}

export const emptyExtras = (): PageSetupExtras => ({ pageSetup: "", headerFooterAttrs: "", headerFooterKids: [] })

const PAGE_SETUP_OWN = new Set([
  "paperSize",
  "scale",
  "fitToWidth",
  "fitToHeight",
  "orientation",
  "pageOrder",
  "firstPageNumber",
  "useFirstPageNumber",
  "blackAndWhite",
])

const HEADER_FOOTER_OWN = new Set(["differentFirst", "differentOddEven", "scaleWithDoc", "alignWithMargins"])
/** headerFooter's children, in the order the schema wants them, and where each goes. */
const HEADER_FOOTER_KIDS: Record<string, HeaderFooterKey> = {
  oddHeader: "header",
  oddFooter: "footer",
  evenHeader: "evenHeader",
  evenFooter: "evenFooter",
  firstHeader: "firstHeader",
  firstFooter: "firstFooter",
}

const yes = (v: string | undefined) => v === "1" || v === "true"
const num = (v: string | undefined, fallback: number) => {
  const n = Number(v)
  return v !== undefined && Number.isFinite(n) ? n : fallback
}

/** "&LLeft&CCentre&RRight" → its three sections (a section with no code is the centre, as Excel reads it). */
export function parseHeaderFooter(text: string): HeaderFooterText {
  const out: HeaderFooterText = { left: "", center: "", right: "" }
  let section: keyof HeaderFooterText = "center"
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    if (ch === "&" && i + 1 < text.length) {
      const code = text[i + 1]
      if (code === "L" || code === "C" || code === "R") {
        section = code === "L" ? "left" : code === "C" ? "center" : "right"
        i += 2
        continue
      }
      out[section] += ch + code
      i += 2
      continue
    }
    out[section] += ch
    i++
  }
  return out
}

export function headerFooterText(hf: HeaderFooterText): string {
  return (hf.left ? "&L" + hf.left : "") + (hf.center ? "&C" + hf.center : "") + (hf.right ? "&R" + hf.right : "")
}

/** Reads one print element of a worksheet into the sheet's page setup. */
export function readPrintElement(local: string, node: XNode, setup: PageSetup, extras: PageSetupExtras) {
  switch (local) {
    case "printOptions":
      setup.gridlines = yes(at(node, "gridLines"))
      setup.headings = yes(at(node, "headings"))
      setup.centerHorizontally = yes(at(node, "horizontalCentered"))
      setup.centerVertically = yes(at(node, "verticalCentered"))
      break
    case "pageMargins": {
      const m = setup.margins
      // Excel writes inches converted from points (0.8999999999999999): 12 digits are what it meant.
      for (const k of ["left", "right", "top", "bottom", "header", "footer"] as const)
        m[k] = Number(num(at(node, k), m[k]).toPrecision(12))
      break
    }
    case "pageSetup": {
      setup.paperSize = at(node, "paperSize") !== undefined ? num(at(node, "paperSize"), 1) : null
      setup.scale = Math.min(400, Math.max(10, num(at(node, "scale"), 100)))
      setup.orientation = at(node, "orientation") === "landscape" ? "landscape" : "portrait"
      setup.order = at(node, "pageOrder") === "overThenDown" ? "overThenDown" : "downThenOver"
      setup.blackAndWhite = yes(at(node, "blackAndWhite"))
      setup.firstPageNumber = yes(at(node, "useFirstPageNumber")) ? num(at(node, "firstPageNumber"), 1) : null
      // Kept until sheetPr says whether the sheet fits to pages (sheetPr comes first; see finishPageSetup).
      setup.fitTo = { width: num(at(node, "fitToWidth"), 1), height: num(at(node, "fitToHeight"), 1) }
      let rest = ""
      for (const [name, value] of Object.entries(node.attrs))
        if (!PAGE_SETUP_OWN.has(name.replace(/^\w+:/, "")) || name.includes(":"))
          rest += ` ${name}="${escapeXml(value)}"`
      extras.pageSetup = rest
      break
    }
    case "headerFooter": {
      setup.differentFirst = yes(at(node, "differentFirst"))
      setup.differentOddEven = yes(at(node, "differentOddEven"))
      setup.scaleWithDoc = !/^(0|false)$/.test(at(node, "scaleWithDoc") ?? "")
      setup.alignWithMargins = !/^(0|false)$/.test(at(node, "alignWithMargins") ?? "")
      for (const [name, value] of Object.entries(node.attrs))
        if (!HEADER_FOOTER_OWN.has(name)) extras.headerFooterAttrs += ` ${name}="${escapeXml(value)}"`
      for (const k of node.children) {
        const key = HEADER_FOOTER_KIDS[k.local]
        if (key) setup[key] = parseHeaderFooter(k.text)
        else extras.headerFooterKids.push(`<${k.local}>${escapeXmlText(k.text)}</${k.local}>`)
      }
      break
    }
    case "rowBreaks":
    case "colBreaks": {
      const list = kids(node, "brk")
        .map((b) => num(at(b, "id"), 0))
        .filter((id) => id > 0)
      if (local === "rowBreaks") setup.rowBreaks = list
      else setup.colBreaks = list
      break
    }
  }
}

/** Once the whole worksheet is read: fit-to only counts when sheetPr turned it on. */
export function finishPageSetup(setup: PageSetup, fitToPage: boolean, sawPageSetup: boolean) {
  if (!fitToPage) setup.fitTo = null
  else if (!sawPageSetup) setup.fitTo = { width: 1, height: 1 }
}

/** The worksheet's print elements for a save ("" leaves one out). */
export function printElementsXml(setup: PageSetup, extras: PageSetupExtras | undefined): Record<string, string> {
  const d = defaultPageSetup()
  const flag = (on: boolean) => (on ? "1" : undefined)
  const options = attrs({
    horizontalCentered: flag(setup.centerHorizontally),
    verticalCentered: flag(setup.centerVertically),
    headings: flag(setup.headings),
    gridLines: flag(setup.gridlines),
  })
  const m = setup.margins
  const pageRest = extras?.pageSetup ?? ""
  const page = (orientation: string | undefined) => ({
    paperSize: setup.paperSize ?? undefined,
    scale: setup.scale !== 100 ? setup.scale : undefined,
    firstPageNumber: setup.firstPageNumber ?? undefined,
    fitToWidth: setup.fitTo && setup.fitTo.width !== 1 ? setup.fitTo.width : undefined,
    fitToHeight: setup.fitTo && setup.fitTo.height !== 1 ? setup.fitTo.height : undefined,
    pageOrder: setup.order !== d.order ? setup.order : undefined,
    orientation,
    blackAndWhite: flag(setup.blackAndWhite),
    useFirstPageNumber: setup.firstPageNumber !== null ? "1" : undefined,
  })
  // Excel names the orientation whenever it writes the element (it reads a missing one as "default", not portrait).
  const written = setup.orientation === "landscape" || !!attrs(page(undefined)) || !!pageRest
  const pageAttrs = attrs(page(written ? setup.orientation : undefined))
  const hfOwn =
    attrs({
      differentOddEven: flag(setup.differentOddEven),
      differentFirst: flag(setup.differentFirst),
      scaleWithDoc: setup.scaleWithDoc ? undefined : "0",
      alignWithMargins: setup.alignWithMargins ? undefined : "0",
    }) + (extras?.headerFooterAttrs ?? "")
  const hfInner =
    Object.entries(HEADER_FOOTER_KIDS)
      .map(([local, key]) => {
        const text = headerFooterText(setup[key])
        return text ? `<${local}>${escapeXmlText(text)}</${local}>` : ""
      })
      .join("") + (extras?.headerFooterKids ?? []).join("")
  const breaks = (local: string, list: number[], max: number) =>
    list.length
      ? `<${local} count="${list.length}" manualBreakCount="${list.length}">${list.map((id) => `<brk id="${id}" max="${max}" man="1"/>`).join("")}</${local}>`
      : ""
  return {
    printOptions: options ? `<printOptions${options}/>` : "",
    pageMargins: `<pageMargins left="${m.left}" right="${m.right}" top="${m.top}" bottom="${m.bottom}" header="${m.header}" footer="${m.footer}"/>`,
    pageSetup: pageAttrs || pageRest ? `<pageSetup${pageAttrs}${pageRest}/>` : "",
    headerFooter: hfInner || hfOwn ? `<headerFooter${hfOwn}>${hfInner}</headerFooter>` : "",
    // A row break's max is the last column (and a column break's the last row), as Excel writes them.
    rowBreaks: breaks("rowBreaks", setup.rowBreaks, MAX_COLS - 1),
    colBreaks: breaks("colBreaks", setup.colBreaks, MAX_ROWS - 1),
  }
}

function attrs(values: Record<string, string | number | undefined>): string {
  let out = ""
  for (const [k, v] of Object.entries(values)) if (v !== undefined) out += ` ${k}="${escapeXml(String(v))}"`
  return out
}

/** sheetPr with pageSetUpPr's fitToPage set or cleared (null: no sheetPr is needed). */
export function withFitToPage(sheetPr: string | null, on: boolean): string | null {
  const has = sheetPr !== null && /<(\w+:)?pageSetUpPr\b/.test(sheetPr)
  if (has) {
    return sheetPr!.replace(
      /<((?:\w+:)?pageSetUpPr)\b([^>]*?)(\/?)>/,
      (_, name: string, rest: string, close: string) => {
        const kept = rest.replace(/\s+fitToPage="[^"]*"/, "")
        return `<${name}${kept}${on ? ' fitToPage="1"' : ""}${close}>`
      },
    )
  }
  if (!on) return sheetPr
  const pr = `<pageSetUpPr fitToPage="1"/>`
  if (sheetPr === null) return `<sheetPr>${pr}</sheetPr>`
  // pageSetUpPr is sheetPr's last child.
  if (/\/>$/.test(sheetPr) && !/<\w[^>]*>[\s\S]*</.test(sheetPr.slice(1)))
    return sheetPr.replace(/\s*\/>$/, `>${pr}</sheetPr>`)
  return sheetPr.replace(/(<\/(?:\w+:)?sheetPr>)$/, `${pr}$1`)
}

/** Whether a kept sheetPr turns fit-to-page on. */
export function readsFitToPage(sheetPr: XNode): boolean {
  return yes(at(kid(sheetPr, "pageSetUpPr"), "fitToPage"))
}

// ── Print_Area and Print_Titles ──

/** The parts of a name's reference ("Sheet1!$A$1:$D$9,Sheet1!$F$1:$G$4"), without their sheet. */
function refParts(text: string): string[] {
  const parts: string[] = []
  let cur = ""
  let quoted = false
  for (const ch of text) {
    if (ch === "'") quoted = !quoted
    if (ch === "," && !quoted) {
      parts.push(cur)
      cur = ""
    } else cur += ch
  }
  parts.push(cur)
  return parts.map((p) =>
    p
      .slice(p.lastIndexOf("!") + 1)
      .replace(/\$/g, "")
      .trim(),
  )
}

/** _xlnm.Print_Area's ranges (null when it isn't plain ranges). */
export function readPrintArea(text: string): Range[] | null {
  const out: Range[] = []
  for (const p of refParts(text)) {
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(p)
    if (m) {
      const c0 = columnIndex(m[1].toUpperCase())
      const r0 = Number(m[2]) - 1
      const c1 = m[3] ? columnIndex(m[3].toUpperCase()) : c0
      const r1 = m[4] ? Number(m[4]) - 1 : r0
      out.push({ r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) })
      continue
    }
    const rows = /^(\d+):(\d+)$/.exec(p)
    if (rows) {
      out.push({ r0: Number(rows[1]) - 1, r1: Number(rows[2]) - 1, c0: 0, c1: MAX_COLS - 1 })
      continue
    }
    const cols = /^([A-Z]+):([A-Z]+)$/i.exec(p)
    if (cols) {
      out.push({
        r0: 0,
        r1: MAX_ROWS - 1,
        c0: columnIndex(cols[1].toUpperCase()),
        c1: columnIndex(cols[2].toUpperCase()),
      })
      continue
    }
    return null
  }
  return out.length ? out : null
}

/** _xlnm.Print_Titles: rows ($1:$2) and columns ($A:$B) to repeat (null when it reads as something else). */
export function readPrintTitles(text: string): { rows: [number, number] | null; cols: [number, number] | null } | null {
  let rows: [number, number] | null = null
  let cols: [number, number] | null = null
  for (const p of refParts(text)) {
    const r = /^(\d+):(\d+)$/.exec(p)
    const c = /^([A-Z]+):([A-Z]+)$/i.exec(p)
    if (r) rows = [Number(r[1]) - 1, Number(r[2]) - 1]
    else if (c) cols = [columnIndex(c[1].toUpperCase()), columnIndex(c[2].toUpperCase())]
    else return null
  }
  return { rows, cols }
}

/** The sheet's Print_Area and Print_Titles names, as Excel writes them. */
export function printNamesXml(sheetName: string, index: number, setup: PageSetup): string[] {
  const q = quoteSheet(sheetName) + "!"
  const cell = (row: number, col: number) => `$${columnName(col)}$${row + 1}`
  const out: string[] = []
  if (setup.printArea?.length) {
    const ref = setup.printArea
      .map((r) => {
        if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return `${q}$${r.r0 + 1}:$${r.r1 + 1}`
        if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return `${q}$${columnName(r.c0)}:$${columnName(r.c1)}`
        return `${q}${cell(r.r0, r.c0)}${r.r0 === r.r1 && r.c0 === r.c1 ? "" : ":" + cell(r.r1, r.c1)}`
      })
      .join(",")
    out.push(`<definedName name="_xlnm.Print_Area" localSheetId="${index}">${escapeXmlText(ref)}</definedName>`)
  }
  if (setup.titleRows || setup.titleCols) {
    const parts: string[] = []
    // Excel writes the columns first.
    if (setup.titleCols) parts.push(`${q}$${columnName(setup.titleCols[0])}:$${columnName(setup.titleCols[1])}`)
    if (setup.titleRows) parts.push(`${q}$${setup.titleRows[0] + 1}:$${setup.titleRows[1] + 1}`)
    out.push(
      `<definedName name="_xlnm.Print_Titles" localSheetId="${index}">${escapeXmlText(parts.join(","))}</definedName>`,
    )
  }
  return out
}
