import type { OpcPart } from "./opc"

/** What differs between Office packages: their main part, what to call it in messages, and the order of parts. */
export type PackageFormat = {
  /** File extension (lower case) → the content type its main part must have. The first is the default. */
  mainByExtension: Record<string, string>
  /** The extensions whose packages may hold macros. */
  macroExtensions: readonly string[]
  /** "an Excel workbook": used in "Cannot open a .zip package as …". */
  mainNoun: string
  /** "workbook part": used in "The package has no …". */
  mainPartNoun: string
  /** Where a part goes in the written zip, the way the Office app orders it (lower first); undefined: the default order. */
  partRank(zipName: string): number | undefined
  /** Parts the reader doesn't interpret but the app keeps on save (theme, properties, custom XML…). */
  passthroughTypes: ReadonlySet<string>
}

const COMMON_PASSTHROUGH = [
  "application/vnd.openxmlformats-officedocument.theme+xml",
  "application/vnd.openxmlformats-package.core-properties+xml",
  "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
  "application/vnd.ms-office.vbaProject",
]

export const CONTENT_TYPE_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
export const CONTENT_TYPE_WORKBOOK_MACRO = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
export const CONTENT_TYPE_WORKBOOK_TEMPLATE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
export const CONTENT_TYPE_WORKBOOK_MACRO_TEMPLATE = "application/vnd.ms-excel.template.macroEnabled.main+xml"

/** SpreadsheetML: .xlsx, .xlsm, .xltx, .xltm (Cells). */
export const SPREADSHEET_PACKAGE: PackageFormat = {
  mainByExtension: {
    xlsx: CONTENT_TYPE_WORKBOOK,
    xlsm: CONTENT_TYPE_WORKBOOK_MACRO,
    xltx: CONTENT_TYPE_WORKBOOK_TEMPLATE,
    xltm: CONTENT_TYPE_WORKBOOK_MACRO_TEMPLATE,
  },
  macroExtensions: ["xlsm", "xltm"],
  mainNoun: "an Excel workbook",
  mainPartNoun: "workbook part",
  partRank(name) {
    if (/^xl\/worksheets\//i.test(name)) return 10
    if (/^xl\/theme\//i.test(name)) return 20
    if (/^xl\/styles\.xml$/i.test(name)) return 21
    if (/^xl\/sharedStrings\.xml$/i.test(name)) return 22
    if (/^xl\//i.test(name)) return 25
    return undefined
  },
  passthroughTypes: new Set(COMMON_PASSTHROUGH),
}

export const CONTENT_TYPE_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
export const CONTENT_TYPE_DOCUMENT_MACRO = "application/vnd.ms-word.document.macroEnabled.main+xml"
export const CONTENT_TYPE_DOCUMENT_TEMPLATE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
export const CONTENT_TYPE_DOCUMENT_MACRO_TEMPLATE = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"

/** WordprocessingML: .docx, .docm, .dotx, .dotm (Inkspire), with Word's part order. */
export const WORD_PACKAGE: PackageFormat = {
  mainByExtension: {
    docx: CONTENT_TYPE_DOCUMENT,
    docm: CONTENT_TYPE_DOCUMENT_MACRO,
    dotx: CONTENT_TYPE_DOCUMENT_TEMPLATE,
    dotm: CONTENT_TYPE_DOCUMENT_MACRO_TEMPLATE,
  },
  macroExtensions: ["docm", "dotm"],
  mainNoun: "a Word document",
  mainPartNoun: "main document part",
  partRank(name) {
    if (/^word\/footnotes\.xml$/i.test(name)) return 10
    if (/^word\/endnotes\.xml$/i.test(name)) return 11
    if (/^word\/header\d*\.xml$/i.test(name)) return 12
    if (/^word\/footer\d*\.xml$/i.test(name)) return 13
    if (/^word\/media\//i.test(name)) return 14
    if (/^word\/theme\//i.test(name)) return 20
    if (/^word\/settings\.xml$/i.test(name)) return 21
    if (/^word\/numbering\.xml$/i.test(name)) return 22
    if (/^word\/styles\.xml$/i.test(name)) return 23
    if (/^word\/webSettings\.xml$/i.test(name)) return 24
    if (/^word\/fontTable\.xml$/i.test(name)) return 25
    return undefined
  },
  passthroughTypes: new Set([
    ...COMMON_PASSTHROUGH,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    "application/vnd.ms-word.vbaData+xml",
  ]),
}

/** The extensions a format opens: documents and templates, with and without macros. */
export function packageExtensions(format: PackageFormat): string[] {
  return Object.keys(format.mainByExtension)
}

/** A part the app keeps as written without reading it. */
export function isPassthroughPart(part: OpcPart, format: PackageFormat) {
  return format.passthroughTypes.has(part.contentType) || part.name.toLowerCase().startsWith("/customxml/")
}
