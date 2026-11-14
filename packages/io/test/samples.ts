// A made-up corpus: workbooks built in code that cover what the xlsx reader and writer handle.
import { DEFAULT_VALIDATION, errorValue, hashPassword, Workbook, type CellData, type CFEntry } from "@cells/core"

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

export const samples: Record<string, () => Workbook> = {
  budget,
  formats,
  conditional,
  layout,
  protected: protectedBook,
  validation,
  big,
  empty,
}
