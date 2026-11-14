import { FUNCTION_INFO } from "./signatures"

/** Excel's function categories, in the Insert Function dialog's order. */
export const FUNCTION_CATEGORIES = [
  "Financial",
  "Date & Time",
  "Math & Trig",
  "Statistical",
  "Lookup & Reference",
  "Text",
  "Logical",
  "Information",
  "Compatibility",
] as const

export type FunctionCategory = (typeof FUNCTION_CATEGORIES)[number]

const LISTS: Partial<Record<FunctionCategory, string>> = {
  Financial: "EFFECT FV IPMT IRR NOMINAL NPER NPV PMT PPMT PV RATE SLN SYD XIRR XNPV",
  "Date & Time":
    "DATE DATEDIF DATEVALUE DAY DAYS DAYS360 EDATE EOMONTH HOUR ISOWEEKNUM MINUTE MONTH NETWORKDAYS NETWORKDAYS.INTL NOW SECOND TIME TIMEVALUE TODAY WEEKDAY WEEKNUM WORKDAY WORKDAY.INTL YEAR YEARFRAC",
  Statistical:
    "AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS CORREL COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS COVARIANCE.P COVARIANCE.S DEVSQ FORECAST FORECAST.LINEAR GEOMEAN HARMEAN INTERCEPT LARGE MAX MAXA MAXIFS MEDIAN MIN MINA MINIFS MODE.SNGL PEARSON PERCENTILE.EXC PERCENTILE.INC PERMUT QUARTILE.EXC QUARTILE.INC RANK.AVG RANK.EQ RSQ SLOPE SMALL STDEV.P STDEV.S STDEVA STDEVPA STEYX VAR.P VAR.S VARA VARPA",
  "Lookup & Reference":
    "ADDRESS AREAS CHOOSE CHOOSECOLS CHOOSEROWS COLUMN COLUMNS DROP EXPAND FILTER FORMULATEXT HLOOKUP HSTACK INDEX INDIRECT LOOKUP MATCH OFFSET ROW ROWS SORT SORTBY TAKE TOCOL TOROW TRANSPOSE UNIQUE VLOOKUP VSTACK WRAPCOLS WRAPROWS XLOOKUP XMATCH",
  Text: "CHAR CLEAN CODE CONCAT CONCATENATE DOLLAR EXACT FIND FIXED LEFT LEN LOWER MID NUMBERVALUE PROPER REPLACE REPT RIGHT SEARCH SUBSTITUTE T TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TRIM UNICHAR UNICODE UPPER VALUE",
  Logical: "AND FALSE IF IFERROR IFNA IFS LET NOT OR SWITCH TRUE XOR",
  Information:
    "ERROR.TYPE ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISREF ISTEXT N NA SHEET SHEETS TYPE",
  Compatibility: "CEILING COVAR FLOOR MODE PERCENTILE QUARTILE RANK STDEV STDEVP VAR VARP",
}

const CATEGORY = new Map<string, FunctionCategory>()
for (const [category, names] of Object.entries(LISTS))
  for (const name of names.split(" ")) CATEGORY.set(name, category as FunctionCategory)

/** A function's category; what no list names is Math & Trig. */
export function functionCategory(name: string): FunctionCategory {
  return CATEGORY.get(name.toUpperCase()) ?? "Math & Trig"
}

/** The functions of a category (all when null), sorted by name. */
export function functionsIn(category: FunctionCategory | null): string[] {
  return [...FUNCTION_INFO.keys()].filter((name) => category === null || functionCategory(name) === category)
}
