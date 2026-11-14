/**
 * Format Cells › Number: the codes Excel (en-US) builds from each category's options, and the reverse (which category
 * and options a code shows when the dialog opens).
 */

export type Category =
  | "General"
  | "Number"
  | "Currency"
  | "Accounting"
  | "Date"
  | "Time"
  | "Percentage"
  | "Fraction"
  | "Scientific"
  | "Text"
  | "Special"
  | "Custom"

export const CATEGORIES: Category[] = [
  "General",
  "Number",
  "Currency",
  "Accounting",
  "Date",
  "Time",
  "Percentage",
  "Fraction",
  "Scientific",
  "Text",
  "Special",
  "Custom",
]

export const CATEGORY_HELP: Record<Category, string> = {
  General: "General format cells have no specific number format.",
  Number:
    "Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.",
  Currency:
    "Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.",
  Accounting: "Accounting formats line up the currency symbols and decimal points in a column.",
  Date: "Date formats display date and time serial numbers as date values. Formats that begin with an asterisk (*) respond to changes in regional date and time settings.",
  Time: "Time formats display date and time serial numbers as time values. Formats that begin with an asterisk (*) respond to changes in regional date and time settings.",
  Percentage: "Percentage formats multiply the cell value by 100 and display the result with a percent symbol.",
  Fraction: "",
  Scientific: "",
  Text: "Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.",
  Special: "Special formats are useful for tracking list and database values.",
  Custom: "Type the number format code, using one of the existing codes as a starting point.",
}

/** Negative-number choices (Number and Currency): minus, red, brackets, red brackets. */
export type Negative = 0 | 1 | 2 | 3

export type NumberOptions = { decimals: number; separator: boolean; negative: Negative; symbol: string }

export function numberCode(o: Pick<NumberOptions, "decimals" | "separator" | "negative">): string {
  const base = (o.separator ? "#,##0" : "0") + (o.decimals > 0 ? "." + "0".repeat(o.decimals) : "")
  return [base, `${base};[Red]${base}`, `${base}_);(${base})`, `${base}_);[Red](${base})`][o.negative]
}

export function currencyCode(o: Pick<NumberOptions, "decimals" | "negative" | "symbol">): string {
  const n = "#,##0" + (o.decimals > 0 ? "." + "0".repeat(o.decimals) : "")
  const s = o.symbol
  const v = `${s}${n}`
  return [v, `${v};[Red]${v}`, `${v}_);(${v})`, `${v}_);[Red](${v})`][o.negative]
}

export function accountingCode(o: Pick<NumberOptions, "decimals" | "symbol">): string {
  const n = "#,##0" + (o.decimals > 0 ? "." + "0".repeat(o.decimals) : "")
  const s = o.symbol ? `${o.symbol}* ` : "* "
  return `_(${s}${n}_);_(${s}(${n});_(${s}"-"??_);_(@_)`
}

export function percentCode(decimals: number) {
  return "0" + (decimals > 0 ? "." + "0".repeat(decimals) : "") + "%"
}

export function scientificCode(decimals: number) {
  return "0" + (decimals > 0 ? "." + "0".repeat(decimals) : "") + "E+00"
}

/** Type lists, with what Excel calls them where the code alone says too little. */
export const DATE_TYPES = [
  "m/d/yyyy",
  "[$-x-sysdate]dddd, mmmm dd, yyyy",
  "m/d;@",
  "m/d/yy;@",
  "mm/dd/yy;@",
  "[$-409]d-mmm;@",
  "[$-409]d-mmm-yy;@",
  "[$-409]dd-mmm-yy;@",
  "[$-409]mmm-yy;@",
  "[$-409]mmmm-yy;@",
  "[$-409]mmmm d, yyyy;@",
  "[$-409]m/d/yy h:mm AM/PM;@",
  "m/d/yy h:mm;@",
  "[$-409]mmmmm;@",
  "[$-409]mmmmm-yy;@",
  "m/d/yyyy;@",
  "[$-409]d-mmm-yyyy;@",
]

export const TIME_TYPES = [
  "[$-x-systime]h:mm:ss AM/PM",
  "h:mm;@",
  "[$-409]h:mm AM/PM;@",
  "h:mm:ss;@",
  "[$-409]h:mm:ss AM/PM;@",
  "mm:ss.0;@",
  "[h]:mm:ss;@",
  "[$-409]m/d/yy h:mm AM/PM;@",
  "m/d/yy h:mm;@",
]

export const FRACTION_TYPES = [
  { code: "# ?/?", label: "Up to one digit (1/4)" },
  { code: "# ??/??", label: "Up to two digits (21/25)" },
  { code: "# ???/???", label: "Up to three digits (312/943)" },
  { code: "# ?/2", label: "As halves (1/2)" },
  { code: "# ?/4", label: "As quarters (2/4)" },
  { code: "# ?/8", label: "As eighths (4/8)" },
  { code: "# ??/16", label: "As sixteenths (8/16)" },
  { code: "# ?/10", label: "As tenths (3/10)" },
  { code: "# ??/100", label: "As hundredths (30/100)" },
]

export const SPECIAL_TYPES = [
  { code: "00000", label: "Zip Code" },
  { code: "00000-0000", label: "Zip Code + 4" },
  { code: "[<=9999999]###-####;(###) ###-####", label: "Phone Number" },
  { code: "000-00-0000", label: "Social Security Number" },
]

/** Excel's built-in custom codes, as the Custom list shows them. */
export const CUSTOM_TYPES = [
  "General",
  "0",
  "0.00",
  "#,##0",
  "#,##0.00",
  "#,##0_);(#,##0)",
  "#,##0_);[Red](#,##0)",
  "#,##0.00_);(#,##0.00)",
  "#,##0.00_);[Red](#,##0.00)",
  "$#,##0_);($#,##0)",
  "$#,##0_);[Red]($#,##0)",
  "$#,##0.00_);($#,##0.00)",
  "$#,##0.00_);[Red]($#,##0.00)",
  "0%",
  "0.00%",
  "0.00E+00",
  "##0.0E+0",
  "# ?/?",
  "# ??/??",
  "m/d/yyyy",
  "d-mmm-yy",
  "d-mmm",
  "mmm-yy",
  "h:mm AM/PM",
  "h:mm:ss AM/PM",
  "h:mm",
  "h:mm:ss",
  "m/d/yyyy h:mm",
  "mm:ss",
  "mm:ss.0",
  "@",
  "[h]:mm:ss",
  '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
]

const decimalsOf = (code: string) => /\.(0+)/.exec(code)?.[1].length ?? 0

/** The category and options a format code opens on. */
export function readCode(code: string | undefined): { category: Category; options: NumberOptions; type?: string } {
  const options: NumberOptions = { decimals: 2, separator: false, negative: 0, symbol: "$" }
  if (!code || code === "General") return { category: "General", options }
  if (code === "@") return { category: "Text", options }
  for (const separator of [false, true])
    for (let d = 0; d <= 30; d++)
      for (const n of [0, 1, 2, 3] as Negative[])
        if (numberCode({ decimals: d, separator, negative: n }) === code)
          return { category: "Number", options: { ...options, decimals: d, separator, negative: n } }
  for (const symbol of ["$", ""]) {
    for (let d = 0; d <= 30; d++) {
      if (accountingCode({ decimals: d, symbol }) === code)
        return { category: "Accounting", options: { ...options, decimals: d, symbol } }
      for (const n of [0, 1, 2, 3] as Negative[])
        if (symbol && currencyCode({ decimals: d, negative: n, symbol }) === code)
          return { category: "Currency", options: { ...options, decimals: d, negative: n, symbol } }
    }
  }
  if (/^0(\.0+)?%$/.test(code)) return { category: "Percentage", options: { ...options, decimals: decimalsOf(code) } }
  if (/^0(\.0+)?E\+00$/.test(code))
    return { category: "Scientific", options: { ...options, decimals: decimalsOf(code) } }
  if (FRACTION_TYPES.some((t) => t.code === code)) return { category: "Fraction", options, type: code }
  if (DATE_TYPES.includes(code)) return { category: "Date", options, type: code }
  if (TIME_TYPES.includes(code)) return { category: "Time", options, type: code }
  if (SPECIAL_TYPES.some((t) => t.code === code)) return { category: "Special", options, type: code }
  return { category: "Custom", options, type: code }
}

/** The code the dialog's current choices make. */
export function buildCode(category: Category, o: NumberOptions, type: string | undefined): string {
  switch (category) {
    case "General":
      return "General"
    case "Number":
      return numberCode(o)
    case "Currency":
      return o.symbol ? currencyCode(o) : numberCode({ ...o, separator: true })
    case "Accounting":
      return accountingCode(o)
    case "Percentage":
      return percentCode(o.decimals)
    case "Scientific":
      return scientificCode(o.decimals)
    case "Text":
      return "@"
    case "Date":
      return type && DATE_TYPES.includes(type) ? type : DATE_TYPES[0]
    case "Time":
      return type && TIME_TYPES.includes(type) ? type : TIME_TYPES[0]
    case "Fraction":
      return type && FRACTION_TYPES.some((t) => t.code === type) ? type : FRACTION_TYPES[0].code
    case "Special":
      return type && SPECIAL_TYPES.some((t) => t.code === type) ? type : SPECIAL_TYPES[0].code
    case "Custom":
      return type || "General"
  }
}
