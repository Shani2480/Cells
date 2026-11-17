// A made-up corpus: workbooks built in code that cover what the xlsx reader and writer handle.
import { defaultChartSpec } from "@cells/chart"
import {
  chartRefsFor,
  consolidateOps,
  DEFAULT_VALIDATION,
  errorValue,
  hashPassword,
  Workbook,
  type CellData,
  type CFEntry,
} from "@cells/core"

const done = (wb: Workbook) => {
  wb.calc.rebuild()
  wb.calculate(true)
  return wb
}

export function budget(): Workbook {
  const wb = new Workbook(["Budget", "Notes"])
  const s = wb.sheets[0]
  const bold = wb.styles.intern({ bold: true, fill: "#FFEB9C", numberFormat: "#,##0.00" })
  const set = (r: number, c: number, d: CellData) => s.cells.set(r, c, d)
  set(0, 0, { value: "Item", style: bold })
  set(0, 1, { value: "Cost", style: bold })
  set(1, 0, { value: "Rent" })
  set(1, 1, { value: 1200.5 })
  set(2, 0, { value: "Food & drink <weekly>" })
  set(2, 1, { value: 310 })
  set(3, 0, { value: "Total" })
  set(3, 1, { value: 0, input: "=SUM(B2:B3)" })
  set(4, 0, { value: true })
  set(4, 1, { value: errorValue("#N/A") })
  set(5, 0, { value: "line one\nline two" })
  set(6, 0, { value: 0, input: '=XLOOKUP("Rent",A2:A3,B2:B3)' })
  set(7, 0, { value: 0, input: "=SEQUENCE(3)" })
  s.cols.setSize(0, 120)
  s.rows.setSize(5, 40)
  s.rows.setHidden(10, true)
  s.merges.add({ r0: 12, c0: 0, r1: 12, c1: 2 })
  s.tabColor = "#C00000"
  s.annotations.set(1, 1, { kind: "note", author: "Alex", text: "Paid on the 1st", shown: false })
  s.annotations.set(2, 1, {
    kind: "comment",
    resolved: false,
    posts: [
      { author: "Alex", text: "Too high?", time: Date.UTC(2026, 0, 5, 10, 30) },
      { author: "Sam", text: "Fine for now", time: Date.UTC(2026, 0, 5, 11, 0) },
    ],
  })
  s.conditional = [
    {
      id: 1,
      ranges: [{ r0: 1, c0: 1, r1: 2, c1: 1 }],
      rule: { kind: "cellIs", operator: "greaterThan", values: ["500"], style: { fill: "#FFC7CE" } },
      stopIfTrue: false,
    },
  ]
  wb.names.replace([{ name: "Costs", ref: "Budget!$B$2:$B$3" }])
  wb.sheets[1].cells.set(0, 0, { value: 0, input: "=Budget!B4*2" })
  return done(wb)
}

export function formats(): Workbook {
  const wb = new Workbook(["Formats"])
  const s = wb.sheets[0]
  const styles = [
    { fontName: "Georgia", fontSize: 14, italic: true, underline: "double", fontColor: "#1F4E79" },
    { strike: true, vertAlign: "superscript" },
    { hAlign: "center", vAlign: "top", wrap: true },
    { hAlign: "left", indent: 2 },
    { rotation: 45 },
    { rotation: 255 },
    { shrink: true, hAlign: "right" },
    { borderTop: { style: "thin" }, borderBottom: { style: "double", color: "#FF0000" } },
    { borderLeft: { style: "mediumDashDot", color: "#00B050" }, borderRight: { style: "hair" } },
    { borderDiagonal: { style: "thin" }, diagonalUp: true, diagonalDown: true },
    { numberFormat: "0.00%" },
    { numberFormat: "yyyy-mm-dd hh:mm" },
    { numberFormat: '_-[$€-2]\\ * #,##0.00_-;\\-[$€-2]\\ * #,##0.00_-;_-[$€-2]\\ * "-"??_-;_-@_-' },
    { numberFormat: "[Red][<0]0.0;[Blue]0.0" },
    { locked: false },
    { hiddenFormula: true },
    { cellStyle: "Good", fill: "#C6EFCE", fontColor: "#006100" },
    {
      cellStyle: "Heading 1",
      bold: true,
      fontSize: 15,
      fontColor: "#0E2841",
      borderBottom: { style: "thick", color: "#156082" },
    },
  ] as const
  styles.forEach((style, i) => {
    s.cells.set(i, 0, { value: `Style ${i + 1}`, style: wb.styles.intern(style as never) })
    s.cells.set(i, 1, { value: 45000.25 + i, style: wb.styles.intern(style as never) })
  })
  return done(wb)
}

export function conditional(): Workbook {
  const wb = new Workbook(["Rules"])
  const s = wb.sheets[0]
  for (let r = 0; r < 20; r++) {
    s.cells.set(r, 0, { value: r * 3 - 10 })
    s.cells.set(r, 1, { value: r % 4 === 0 ? "apple" : "pear" })
    s.cells.set(r, 2, { value: 45900 + r })
  }
  const col = (c: number) => [{ r0: 0, c0: c, r1: 19, c1: c }]
  const rules: CFEntry["rule"][] = [
    { kind: "cellIs", operator: "between", values: ["0", "10"], style: { bold: true } },
    { kind: "text", operator: "contains", text: "app", style: { fontColor: "#9C0006" } },
    { kind: "text", operator: "beginsWith", text: "pe", style: { italic: true } },
    { kind: "date", period: "lastMonth", style: { fill: "#FFEB9C" } },
    { kind: "duplicate", unique: false, style: { fill: "#FFC7CE" } },
    { kind: "top", bottom: false, percent: true, rank: 10, style: { fill: "#C6EFCE" } },
    { kind: "average", below: true, equal: false, style: { underline: "single" } },
    { kind: "blanks", blank: true, style: { fill: "#D9D9D9" } },
    { kind: "errors", errors: true, style: { strike: true } },
    { kind: "expression", formula: "MOD(ROW(),2)=0", style: { fill: "#DDEBF7" } },
    {
      kind: "colorScale",
      points: [
        { at: { type: "min" }, color: "#F8696B" },
        { at: { type: "percentile", value: "50" }, color: "#FFEB84" },
        { at: { type: "max" }, color: "#63BE7B" },
      ],
    },
    { kind: "dataBar", min: { type: "min" }, max: { type: "max" }, color: "#638EC6", gradient: true },
    {
      kind: "iconSet",
      set: "3Arrows",
      reverse: false,
      showValue: true,
      thresholds: [
        { type: "percent", value: "33" },
        { type: "percent", value: "67" },
      ],
    },
  ]
  const cols = [0, 1, 1, 2, 1, 0, 0, 1, 0, 0, 0, 0, 0]
  s.conditional = rules.map((rule, i) => ({ id: i + 1, ranges: col(cols[i]), rule, stopIfTrue: false }))
  return done(wb)
}

export function layout(): Workbook {
  const wb = new Workbook(["Frozen", "Hidden", "Very wide"])
  const [a, b, c] = wb.sheets
  for (let r = 0; r < 50; r++) for (let col = 0; col < 10; col++) a.cells.set(r, col, { value: r * 10 + col })
  a.cols.setHidden(3, true)
  a.cols.setSize(4, 200)
  a.colStyles.set(5, wb.styles.intern({ fill: "#FCE4D6" }))
  a.rowStyles.set(2, wb.styles.intern({ bold: true }))
  b.hidden = true
  b.cells.set(0, 0, { value: "secret" })
  c.cells.set(0, 16383, { value: "XFD1" })
  c.cells.set(1048575, 0, { value: "A1048576" })
  return done(wb)
}

export function protectedBook(): Workbook {
  const wb = new Workbook(["Locked", "Open"])
  wb.sheets[0].cells.set(0, 0, { value: "Read only" })
  wb.sheets[0].cells.set(1, 0, { value: "Type here", style: wb.styles.intern({ locked: false }) })
  wb.sheets[0].protection = {
    password: hashPassword("cells"),
    selectLocked: true,
    selectUnlocked: true,
    formatCells: true,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: true,
    insertHyperlinks: false,
    deleteColumns: false,
    deleteRows: false,
    sort: true,
    autoFilter: false,
    pivotTables: false,
    editObjects: false,
    editScenarios: false,
  }
  wb.structure.protection = { password: hashPassword("book") }
  return done(wb)
}

export function big(): Workbook {
  const wb = new Workbook(["Data"])
  const s = wb.sheets[0]
  for (let r = 0; r < 20_000; r++) {
    s.cells.set(r, 0, { value: r + 1 })
    s.cells.set(r, 1, { value: `Row ${r % 500}` })
    s.cells.set(r, 2, { value: 0, input: `=A${r + 1}*2` })
  }
  return done(wb)
}

export function empty(): Workbook {
  return done(new Workbook(["Sheet1"]))
}

export function validation(): Workbook {
  const wb = new Workbook(["Form", "Lists"])
  const [form, lists] = wb.sheets
  ;["Red", "Green", "Blue"].forEach((v, i) => lists.cells.set(i, 0, { value: v }))
  form.cells.set(0, 0, { value: "Qty" })
  form.cells.set(0, 1, { value: "Colour" })
  form.cells.set(1, 0, { value: 3 })
  const col = (c: number) => [{ r0: 1, c0: c, r1: 30, c1: c }]
  const rule = DEFAULT_VALIDATION
  form.validations = [
    {
      ranges: col(0),
      rule: { ...rule, type: "whole", formula1: "1", formula2: "10", promptTitle: "Qty", prompt: "1 to 10 & <b>" },
    },
    { ranges: col(1), rule: { ...rule, type: "list", formula1: "Lists!$A$1:$A$3", errorStyle: "warning" } },
    { ranges: col(2), rule: { ...rule, type: "list", formula1: '"Yes,No"', dropDown: false } },
    { ranges: col(3), rule: { ...rule, type: "textLength", operator: "lessThan", formula1: "8", error: "Too long" } },
    { ranges: [...col(4), ...col(6)], rule: { ...rule, type: "custom", formula1: "ISNUMBER(E2)", allowBlank: false } },
  ]
  return done(wb)
}

export function tables(): Workbook {
  const wb = new Workbook(["Orders", "Summary"])
  const [orders, summary] = wb.sheets
  const rows: (string | number)[][] = [
    ["Item", "Qty", "Unit Price", "Amount"],
    ["Pen", 2, 1.5],
    ["Book", 1, 12],
    ["Cup", 4, 3],
    ["Lamp", 1, 25],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => orders.cells.set(row, col, { value })))
  for (let row = 1; row <= 4; row++) orders.cells.set(row, 3, { value: 0, input: "=[@Qty]*[@[Unit Price]]" })
  orders.cells.set(5, 0, { value: "Total" })
  orders.cells.set(5, 1, { value: 0, input: "=SUBTOTAL(109,[Qty])" })
  orders.cells.set(5, 3, { value: 0, input: "=SUBTOTAL(109,[Amount])" })
  orders.tables = [
    {
      id: 1,
      name: "Orders",
      range: { r0: 0, c0: 0, r1: 5, c1: 3 },
      header: true,
      totals: true,
      columns: [
        { name: "Item", totalsLabel: "Total" },
        { name: "Qty", totals: "sum" },
        { name: "Unit Price" },
        { name: "Amount", totals: "sum", formula: "[@Qty]*[@[Unit Price]]" },
      ],
      style: "TableStyleLight9",
      firstColumn: true,
      lastColumn: false,
      rowStripes: true,
      columnStripes: false,
      // The table's own filter, with Qty filtered to 2 and 4 (Book and Lamp hidden; the totals leave them out).
      autoFilter: { range: { r0: 0, c0: 0, r1: 4, c1: 3 }, columns: { 1: { kind: "values", values: ["2", "4"] } } },
    },
  ]
  orders.rows.setHidden(2, true)
  orders.rows.setHidden(4, true)
  summary.cells.set(0, 0, { value: 0, input: "=SUM(Orders[Amount])" })
  summary.cells.set(1, 0, { value: 0, input: "=COUNTA(Orders[[#Headers],[Item]:[Qty]])" })
  summary.cells.set(2, 0, { value: 0, input: "=ROWS(Orders[#All])" })
  summary.cells.set(3, 0, { value: 0, input: "=Orders[[#Totals],[Qty]]" })
  return done(wb)
}

/** Two charts on a sheet of quarterly sales: a column chart of both regions and a pie of one quarter's. */
export function charts(): Workbook {
  const wb = new Workbook(["Sales"])
  const [sheet] = wb.sheets
  const rows: (string | number)[][] = [
    ["Quarter", "North", "South"],
    ["Q1", 120, 80],
    ["Q2", 150, 95],
    ["Q3", 135, 110],
    ["Q4", 170, 125],
  ]
  rows.forEach((r, row) => r.forEach((value, col) => sheet.cells.set(row, col, { value })))
  const column = { ...defaultChartSpec("column"), categories: [], series: [], title: { show: true, text: "Sales" } }
  const pie = { ...defaultChartSpec("pie"), categories: [], series: [], title: { show: true, text: "North" } }
  sheet.charts = [
    {
      id: 1,
      name: "Chart 1",
      anchor: { from: { row: 1, col: 4, x: 10, y: 5 }, to: { row: 16, col: 11, x: 0, y: 0 } },
      spec: column as unknown as Record<string, unknown>,
      refs: chartRefsFor(sheet, { r0: 0, c0: 0, r1: 4, c1: 2 }),
    },
    {
      id: 2,
      name: "Chart 2",
      anchor: { from: { row: 18, col: 4, x: 0, y: 0 }, to: { row: 33, col: 11, x: 0, y: 0 } },
      spec: pie as unknown as Record<string, unknown>,
      refs: chartRefsFor(sheet, { r0: 0, c0: 0, r1: 4, c1: 1 }),
    },
  ]
  return done(wb)
}

/** Every print setting Cells writes: two sheets, a print area with titles, breaks, fit-to, headers and footers. */
export function printing(): Workbook {
  const wb = new Workbook(["Report", "Wide"])
  const [report, wide] = wb.sheets
  for (let row = 0; row < 80; row++)
    for (let col = 0; col < 6; col++)
      report.cells.set(row, col, { value: row === 0 ? `Head ${col + 1}` : row * 10 + col })
  for (let col = 0; col < 40; col++) wide.cells.set(0, col, { value: col })
  Object.assign(report.pageSetup, {
    orientation: "landscape",
    paperSize: 9,
    scale: 85,
    margins: { left: 0.5, right: 0.5, top: 1, bottom: 0.9, header: 0.4, footer: 0.35 },
    centerHorizontally: true,
    gridlines: true,
    headings: true,
    firstPageNumber: 3,
    header: { left: "", center: "&A", right: "" },
    footer: { left: "&D", center: "", right: "Page &P of &N" },
    printArea: [{ r0: 0, c0: 0, r1: 59, c1: 5 }],
    titleRows: [0, 0],
    titleCols: [0, 0],
    rowBreaks: [30],
  })
  Object.assign(wide.pageSetup, {
    paperSize: 1,
    fitTo: { width: 1, height: 0 },
    order: "overThenDown",
    colBreaks: [10],
    blackAndWhite: true,
  })
  return done(wb)
}

/** What-If Analysis: a loan payment with a one-input and a two-input data table. */
export function whatIf(): Workbook {
  const wb = new Workbook(["Loan"])
  const s = wb.sheets[0]
  const put = (row: number, col: number, v: string | number) =>
    s.cells.set(row, col, typeof v === "string" && v.startsWith("=") ? { input: v, value: 0 } : { value: v })
  put(0, 0, "Rate")
  put(0, 1, 0.05)
  put(1, 0, "Months")
  put(1, 1, 60)
  put(2, 0, "Amount")
  put(2, 1, 10000)
  put(3, 0, "Payment")
  put(3, 1, "=PMT(B1/12,B2,-B3)")
  // One input: rates down D2:D5, the payment for each in E2:E5.
  put(0, 4, "=B4")
  ;[0.03, 0.04, 0.05, 0.06].forEach((r, i) => {
    put(i + 1, 3, r)
    put(i + 1, 4, "=TABLE(,B1)")
  })
  // Two inputs: amounts across H1:J1, months down G2:G4.
  put(0, 6, "=B4")
  ;[5000, 10000, 15000].forEach((a, i) => put(0, 7 + i, a))
  ;[36, 48, 60].forEach((m, i) => {
    put(i + 1, 6, m)
    for (let c = 0; c < 3; c++) put(i + 1, 7 + c, "=TABLE(B3,B2)")
  })
  // Scenario Manager: two sets of rate and amount.
  const scenario = (name: string, rate: number, amount: number) => ({
    name,
    comment: "Created by Alex on 1/1/2026",
    user: "Alex",
    locked: true,
    hidden: false,
    cells: [
      { row: 0, col: 1, value: rate },
      { row: 2, col: 1, value: amount },
    ],
  })
  s.scenarios = {
    scenarios: [scenario("Low", 0.03, 8000), scenario("High", 0.07, 15000)],
    current: 0,
    show: 0,
    resultCells: [{ r0: 3, c0: 1, r1: 3, c1: 1 }],
  }
  return done(wb)
}

/** Data › Consolidate with links: two regions totalled on a third sheet, detail rows grouped under each total. */
export function consolidated(): Workbook {
  const wb = new Workbook(["Total", "East", "West Side"])
  const fill = (sheet: number, rows: (string | number)[][]) =>
    rows.forEach((r, ri) => r.forEach((v, ci) => wb.sheets[sheet].cells.set(ri, ci, { value: v })))
  fill(1, [
    ["Item", "Jan", "Feb"],
    ["Pens", 4, 6],
    ["Ink", 2, 3],
  ])
  fill(2, [
    ["Item", "Feb", "Mar"],
    ["Ink", 5, 1],
    ["Paper", 9, 7],
  ])
  const sources = [
    { sheet: 1, range: { r0: 0, c0: 0, r1: 2, c1: 2 } },
    { sheet: 2, range: { r0: 0, c0: 0, r1: 2, c1: 2 } },
  ]
  const o = { fn: 0, top: true, left: true, links: true }
  const result = consolidateOps(wb.sheets, wb.styles, 0, { r0: 0, c0: 0, r1: 0, c1: 0 }, sources, o)
  if ("problem" in result) throw new Error(result.problem)
  wb.transact(result.ops, { label: "Consolidate", history: false })
  wb.sheets[0].consolidate = { ...o, refs: ["East!$A$1:$C$3", "'West Side'!$A$1:$C$3"] }
  return done(wb)
}

/** Forecast Sheet's formulas over a seasonal history (the forecast sheet's table and chart are built in the app). */
export function forecast(): Workbook {
  const wb = new Workbook(["Forecast"])
  const s = wb.sheets[0]
  const season = [5, -3, 8, -10]
  s.cells.set(0, 0, { value: "Month" })
  s.cells.set(0, 1, { value: "Sales" })
  for (let i = 0; i < 24; i++) {
    s.cells.set(i + 1, 0, { value: i + 1 })
    if (i !== 10) s.cells.set(i + 1, 1, { value: 100 + 2 * i + season[i % 4]! })
  }
  for (let i = 0; i < 4; i++) {
    const r = 26 + i
    s.cells.set(r - 1, 0, { value: 25 + i })
    s.cells.set(r - 1, 2, { input: `=FORECAST.ETS(A${r},$B$2:$B$25,$A$2:$A$25,1,1)`, value: 0 })
    s.cells.set(r - 1, 3, { input: `=FORECAST.ETS.CONFINT(A${r},$B$2:$B$25,$A$2:$A$25,0.95,1,1)`, value: 0 })
  }
  s.cells.set(0, 5, { input: "=FORECAST.ETS.SEASONALITY($B$2:$B$25,$A$2:$A$25)", value: 0 })
  s.cells.set(1, 5, { input: "=FORECAST.ETS.STAT($B$2:$B$25,$A$2:$A$25,8)", value: 0 })
  return done(wb)
}

/** Data from text files that stays linked to them (Data › From Text/CSV): a delimited and a fixed-width import. */
export function linkedText(): Workbook {
  const wb = new Workbook(["Sales Data", "Codes"])
  const [sales, codes] = wb.sheets
  const rows = [
    ["Item", "Qty", "Code"],
    ["Pens", 4, "007"],
    ["Ink", 2, "012"],
  ]
  const text = wb.styles.intern({ numberFormat: "@" })
  rows.forEach((row, r) =>
    row.forEach((value, c) => sales.cells.set(r, c, { value, ...(c === 2 && r > 0 ? { style: text } : {}) })),
  )
  sales.queryTables = [
    {
      name: "sales",
      connection: "sales",
      description: "Weekly export from the shop",
      range: { r0: 0, c0: 0, r1: 2, c1: 2 },
      source: "C:\\Users\\alex\\Documents\\sales.csv",
      options: {
        encoding: "utf-8",
        kind: "delimited",
        delimiters: [";", "|"],
        qualifier: '"',
        consecutive: false,
        breaks: [],
        startRow: 1,
        columns: ["general", "skip", "general", "text"],
        decimal: ".",
        thousands: ",",
      },
      refreshOnLoad: true,
    },
  ]
  codes.cells.set(0, 0, { value: "A1" })
  codes.cells.set(0, 1, { value: 25 })
  codes.queryTables = [
    {
      name: "codes",
      connection: "Product codes",
      description: "",
      range: { r0: 0, c0: 0, r1: 0, c1: 1 },
      source: "C:\\Users\\alex\\Documents\\codes.txt",
      options: {
        encoding: "windows-1252",
        kind: "fixed",
        delimiters: [],
        qualifier: null,
        consecutive: false,
        breaks: [4],
        startRow: 2,
        columns: ["text", "DMY"],
        decimal: ",",
        thousands: ".",
      },
      refreshOnLoad: false,
    },
  ]
  return done(wb)
}

export const samples: Record<string, () => Workbook> = {
  linkedText,
  forecast,
  consolidated,
  whatIf,
  budget,
  formats,
  conditional,
  layout,
  protected: protectedBook,
  validation,
  tables,
  charts,
  printing,
  big,
  empty,
}
