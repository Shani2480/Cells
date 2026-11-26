import {
  copyTheme,
  OFFICE_EFFECTS,
  OFFICE_THEME,
  sameLook,
  SHEET_THEME_SLOTS,
  type Theme,
  type ThemeSlot,
} from "@cells/core"
import { escapeXml, readZip, writeZip } from "@cells/ooxml"
import { themeColorsFrom, DEFAULT_THEME_COLORS, type ThemeColors } from "./colors"
import { at, kid, parseTree, type XNode } from "./dom"

/**
 * The workbook's theme part (xl/theme/theme1.xml): read into the workbook's Theme, and written back with the colours
 * and fonts of Page Layout › Themes. A file's own part keeps everything else (effects, other scripts' fonts, extras).
 */

const SCHEME_ORDER: ThemeSlot[] = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
]

const find = (n: XNode, local: string): XNode | undefined =>
  n.local === local ? n : n.children.map((c) => find(c, local)).find((x) => x !== undefined)

/** The theme's colour scheme (`a:clrScheme`) in SpreadsheetML's order, for resolving cell colours. */
export function readThemeColors(xml: string | undefined): ThemeColors {
  if (!xml) return DEFAULT_THEME_COLORS
  const root = find(parseTree(xml).root, "clrScheme")
  if (!root) return DEFAULT_THEME_COLORS
  const scheme: Record<string, string> = {}
  for (const slot of root.children) {
    const color = slot.children[0]
    if (!color) continue
    const hex = color.local === "sysClr" ? at(color, "lastClr") : at(color, "val")
    if (hex && /^[0-9A-Fa-f]{6}$/.test(hex)) scheme[slot.local] = hex.toUpperCase()
  }
  return themeColorsFrom(scheme)
}

/** The file's format scheme (Effects), whitespace between tags dropped; none when it's Office's or not a:-prefixed. */
function formatScheme(xml: string): Theme["effects"] {
  const m = /<a:fmtScheme\b[^>]*>[\s\S]*?<\/a:fmtScheme>/.exec(xml)
  if (!m) return undefined
  const scheme = m[0].replace(/>\s+</g, "><")
  if (scheme === OFFICE_EFFECTS.scheme) return undefined
  return { name: /\bname="([^"]*)"/.exec(scheme)?.[1] ?? "", scheme }
}

/** The theme as the workbook keeps it: names, colours (by slot) and fonts. */
export function readTheme(xml: string | undefined, colors: ThemeColors): Theme {
  const theme = copyTheme(OFFICE_THEME)
  if (!xml) return theme
  SHEET_THEME_SLOTS.forEach((slot, i) => (theme.colors[slot] = colors[i]))
  const doc = parseTree(xml)
  theme.name = at(doc.root, "name") ?? theme.name
  const font = (local: string) => at(kid(find(doc.root, local) ?? doc.root, "latin"), "typeface")
  theme.fonts = { major: font("majorFont") || theme.fonts.major, minor: font("minorFont") || theme.fonts.minor }
  const colorsName = at(find(doc.root, "clrScheme") ?? doc.root, "name")
  const fontsName = at(find(doc.root, "fontScheme") ?? doc.root, "name")
  if (colorsName && colorsName !== theme.name) theme.colorsName = colorsName
  if (fontsName && fontsName !== theme.name) theme.fontsName = fontsName
  const effects = formatScheme(xml)
  if (effects) theme.effects = effects
  return theme
}

/** `a:clrScheme` for a theme (the dark and light 1 as the system colours Office writes when they are black and white). */
function colorSchemeXml(theme: Theme, p: string): string {
  const slot = (s: ThemeSlot) => {
    const hex = theme.colors[s].toUpperCase()
    const sys =
      s === "dk1" && hex === "000000"
        ? `<${p}sysClr val="windowText" lastClr="000000"/>`
        : s === "lt1" && hex === "FFFFFF"
          ? `<${p}sysClr val="window" lastClr="FFFFFF"/>`
          : `<${p}srgbClr val="${hex}"/>`
    return `<${p}${s}>${sys}</${p}${s}>`
  }
  const name = escapeXml(theme.colorsName ?? theme.name)
  return `<${p}clrScheme name="${name}">${SCHEME_ORDER.map(slot).join("")}</${p}clrScheme>`
}

function fontSchemeXml(theme: Theme): string {
  const font = (tag: string, face: string) =>
    `<a:${tag}><a:latin typeface="${escapeXml(face)}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`
  return (
    `<a:fontScheme name="${escapeXml(theme.fontsName ?? theme.name)}">` +
    font("majorFont", theme.fonts.major) +
    font("minorFont", theme.fonts.minor) +
    `</a:fontScheme>`
  )
}

/** A whole theme part, for a workbook that had none. */
export function themeXml(theme: Theme): string {
  return (
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(theme.name)}">` +
    `<a:themeElements>${colorSchemeXml(theme, "a:")}${fontSchemeXml(theme)}${(theme.effects ?? OFFICE_EFFECTS).scheme}</a:themeElements>` +
    `<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`
  )
}

/**
 * A file's theme part with the colours, fonts and names of `theme`; everything else stays as it was. Returns the root
 * element (the XML declaration is written with the part).
 */
export function patchThemeXml(xml: string, theme: Theme): string {
  xml = xml.replace(/^\uFEFF?<\?xml[^>]*\?>\s*/, "")
  const prefix = /<(\w+:)?theme\b/.exec(xml)?.[1] ?? ""
  let out = xml.replace(/(<(?:\w+:)?theme\b[^>]*?\bname=")[^"]*(")/, `$1${escapeXml(theme.name)}$2`)
  out = out.replace(/<(\w+:)?clrScheme\b[\s\S]*?<\/\1?clrScheme>/, colorSchemeXml(theme, prefix))
  out = out.replace(/(<(?:\w+:)?fontScheme\b[^>]*?\bname=")[^"]*(")/, `$1${escapeXml(theme.fontsName ?? theme.name)}$2`)
  // The theme family names Office's theme it came from; Excel would take that one's look.
  out = out.replace(/<(\w+:)?ext\b[^>]*>\s*<(\w+:)?themeFamily\b[^>]*\/>\s*<\/\1?ext>/, "")
  out = out.replace(/<(\w+:)?extLst>\s*<\/\1?extLst>/, "")
  // Effects: the scheme picked, or back to Office's (only when the file's is a:-prefixed, as Office writes it).
  if (prefix === "a:" || prefix === "")
    out = out.replace(/<a:fmtScheme\b[\s\S]*?<\/a:fmtScheme>/, () => (theme.effects ?? OFFICE_EFFECTS).scheme)
  for (const [tag, face] of [
    ["majorFont", theme.fonts.major],
    ["minorFont", theme.fonts.minor],
  ] as const) {
    out = out.replace(
      new RegExp(`(<(?:\\w+:)?${tag}>[\\s\\S]*?<(?:\\w+:)?latin\\b[^>]*?\\btypeface=")[^"]*(")`),
      `$1${escapeXml(face)}$2`,
    )
  }
  return out
}

/** Whether the theme needs writing: changed since it was read, or not Office's default for a new workbook. */
export function themeChanged(theme: Theme, read: Theme | undefined): boolean {
  const before = read ?? OFFICE_THEME
  return (
    !sameLook(theme, before) ||
    theme.name !== before.name ||
    (theme.colorsName ?? "") !== (before.colorsName ?? "") ||
    (theme.fontsName ?? "") !== (before.fontsName ?? "") ||
    (theme.effects?.scheme ?? "") !== (before.effects?.scheme ?? "")
  )
}

/**
 * Page Layout › Themes › Browse for Themes: the theme in an Office theme file (.thmx), or in a document that has one
 * (.xlsx, .docx, .pptx); null when there is none.
 */
export async function themeFromFile(bytes: Uint8Array): Promise<Theme | null> {
  let xml: string | undefined
  try {
    const entries = await readZip(bytes)
    const part =
      entries.find((e) => /^theme\/theme\/theme\d*\.xml$/i.test(e.name)) ??
      entries.find((e) => /^(xl|word|ppt)\/theme\/theme1\.xml$/i.test(e.name))
    xml = part ? new TextDecoder().decode(part.data) : undefined
  } catch {
    return null
  }
  if (!xml || !/clrScheme/.test(xml)) return null
  return readTheme(xml, readThemeColors(xml))
}

/**
 * Page Layout › Themes › Save Current Theme: the theme as an Office theme file (.thmx), which Excel, Word and
 * PowerPoint open with Browse for Themes.
 */
export async function themeFileBytes(theme: Theme): Promise<Uint8Array> {
  const enc = (s: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${s}`)
  const ct = "application/vnd.openxmlformats-officedocument"
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  const parts = [
    {
      name: "[Content_Types].xml",
      data: enc(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/theme/theme/themeManager.xml" ContentType="${ct}.themeManager+xml"/>` +
          `<Override PartName="/theme/theme/theme1.xml" ContentType="${ct}.theme+xml"/></Types>`,
      ),
    },
    {
      name: "_rels/.rels",
      data: enc(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${rel}/officeDocument" Target="theme/theme/themeManager.xml"/></Relationships>`,
      ),
    },
    {
      name: "theme/theme/themeManager.xml",
      data: enc(`<a:themeManager xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`),
    },
    {
      name: "theme/theme/_rels/themeManager.xml.rels",
      data: enc(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${rel}/theme" Target="theme1.xml"/></Relationships>`,
      ),
    },
    { name: "theme/theme/theme1.xml", data: enc(themeXml(theme)) },
  ]
  return writeZip(parts.map((p) => ({ ...p, stored: false })))
}
