import {
  allowsMacros as packageAllowsMacros,
  mainContentTypeFor as packageMainType,
  openPackage as openOfficePackage,
  packageExtensions,
  SPREADSHEET_PACKAGE,
  writePackage as writeOfficePackage,
  type OpcPackage,
} from "@cells/ooxml"

// The shared OPC layer (@cells/ooxml) bound to SpreadsheetML.
export { findPart, OpcError, partKey, relationshipsFor } from "@cells/ooxml"
export type { OpcPackage, OpcPart, OpcRelationship, RelationshipGroup } from "@cells/ooxml"

/** The Excel package extensions: workbooks and templates, with and without macros. */
export const EXCEL_PACKAGE_EXTENSIONS = packageExtensions(SPREADSHEET_PACKAGE)

/** Whether a workbook with this file name may hold macros (.xlsm, .xltm). */
export const allowsMacros = (fileName: string) => packageAllowsMacros(fileName, SPREADSHEET_PACKAGE)

/** The main part's content type a workbook with this file name needs (.xlsx when the name has no Excel extension). */
export const mainContentTypeFor = (fileName: string) => packageMainType(fileName, SPREADSHEET_PACKAGE)

export const openPackage = (bytes: Uint8Array, fileName = "Book1.xlsx") =>
  openOfficePackage(bytes, fileName, SPREADSHEET_PACKAGE)

export const writePackage = (pkg: OpcPackage) => writeOfficePackage(pkg, SPREADSHEET_PACKAGE)
