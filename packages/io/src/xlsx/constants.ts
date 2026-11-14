/** SpreadsheetML names, relationship types and content types, and the tables Excel leaves implicit. */

export const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
export const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
export const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
export const NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
export const NS_XDA = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
export const REL_OFFICE_DOCUMENT = REL + "officeDocument"
export const REL_WORKSHEET = REL + "worksheet"
export const REL_STYLES = REL + "styles"
export const REL_SHARED_STRINGS = REL + "sharedStrings"
export const REL_THEME = REL + "theme"
export const REL_CALC_CHAIN = REL + "calcChain"
export const REL_METADATA = REL + "sheetMetadata"
export const REL_COMMENTS = REL + "comments"
export const REL_VML = REL + "vmlDrawing"
export const REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
export const REL_APP = REL + "extended-properties"
export const REL_THREADED = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
export const REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"

const CT = "application/vnd.openxmlformats-officedocument.spreadsheetml."
export const CT_WORKSHEET = CT + "worksheet+xml"
export const CT_STYLES = CT + "styles+xml"
export const CT_SHARED_STRINGS = CT + "sharedStrings+xml"
export const CT_METADATA = CT + "sheetMetadata+xml"
export const CT_COMMENTS = CT + "comments+xml"
export const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"
export const CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
export const CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
export const CT_APP = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
export const CT_THREADED = "application/vnd.ms-excel.threadedcomments+xml"
export const CT_PERSON = "application/vnd.ms-excel.person+xml"

/** Number formats Excel doesn't write out (en-US): id → code. Ids 5–8 and 41–44 follow the US currency. */
export const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
}

/** The legacy colour palette that `indexed` colours point into (64 and 65 are the system text and window colours). */
export const INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
] // prettier-ignore

/** Excel's built-in cell styles by name → builtinId (the ones the Cell Styles gallery offers). */
export const BUILTIN_CELL_STYLES: Record<string, number> = {
  Normal: 0,
  Comma: 3,
  Currency: 4,
  Percent: 5,
  "Comma [0]": 6,
  "Currency [0]": 7,
  Hyperlink: 8,
  "Followed Hyperlink": 9,
  Note: 10,
  "Warning Text": 11,
  Title: 15,
  "Heading 1": 16,
  "Heading 2": 17,
  "Heading 3": 18,
  "Heading 4": 19,
  Input: 20,
  Output: 21,
  Calculation: 22,
  "Check Cell": 23,
  "Linked Cell": 24,
  Total: 25,
  Good: 26,
  Bad: 27,
  Neutral: 28,
  "Explanatory Text": 53,
}

/**
 * Functions newer than Excel 2007 that a file stores with a prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`). Read
 * from what Excel 16.0 (build 17932) wrote for every function Cells has; `_xlpm.` marks LET's names.
 */
export const FUTURE_FUNCTIONS: Record<string, string> = Object.fromEntries(
  [
    "AGGREGATE", "ARABIC", "BASE", "CEILING.MATH", "CEILING.PRECISE", "CHOOSECOLS", "CHOOSEROWS", "COMBINA",
    "CONCAT", "COT", "COVARIANCE.P", "COVARIANCE.S", "CSC", "DAYS", "DECIMAL", "DROP", "EXPAND", "FLOOR.MATH",
    "FLOOR.PRECISE", "FORECAST.LINEAR", "FORMULATEXT", "HSTACK", "IFNA", "IFS", "ISFORMULA", "ISOWEEKNUM", "LET",
    "MAXIFS", "MINIFS", "MODE.SNGL", "NUMBERVALUE", "PERCENTILE.EXC", "PERCENTILE.INC", "QUARTILE.EXC",
    "QUARTILE.INC", "RANDARRAY", "RANK.AVG", "RANK.EQ", "SEC", "SEQUENCE", "SHEET", "SHEETS", "SORTBY", "STDEV.P",
    "STDEV.S", "SWITCH", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "TOCOL", "TOROW", "UNICHAR",
    "UNICODE", "UNIQUE", "VAR.P", "VAR.S", "VSTACK", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "XOR",
  ]
    .map((name): [string, string] => [name, "_xlfn."])
    .concat([
      ["FILTER", "_xlfn._xlws."],
      ["SORT", "_xlfn._xlws."],
    ]),
) // prettier-ignore
