import { parseInput, parseRange, type CFEntry, type CellWrite, type Op, type StylePatch } from "@cells/core"
import type { SheetController } from "@/sheet/controller"
import { embedDocument, OBJECT_ICON_SIZE } from "@/drawing/embed"
import { SIGNATURE_SIZE, signatureDocument } from "@/drawing/signature"

/**
 * Made-up data for trying the grid at scale: `?demo=100k` (100,000 rows × 20 columns) or `?demo=small`. Nothing
 * here comes from a real file.
 */
export function loadDemo(controller: SheetController, name: string | null) {
  if (!name) return
  if (name === "formats") return formatsDemo(controller)
  if (name === "objects") {
    formatsDemo(controller)
    return objectsDemo(controller)
  }
  if (name === "notes") {
    loadDemo(controller, "small")
    return notesDemo(controller)
  }
  if (name === "cf") {
    loadDemo(controller, "small")
    return conditionalDemo(controller)
  }
  const rows = name === "small" ? 30 : name === "1m" ? 1_000_000 : 100_000
  const cols = name === "1m" ? 5 : 20
  const regions = ["North", "South", "East", "West", "Central"]
  const products = ["Widget", "Gadget", "Gizmo", "Doohickey", "Sprocket", "Flange"]
  const cells: CellWrite[] = []
  const headers = ["Order", "Region", "Product", "Units", "Price", "Total"]
  for (let c = 0; c < cols; c++) {
    cells.push({ row: 0, col: c, data: { value: headers[c] ?? `Metric ${c - headers.length + 1}` } })
  }
  for (let r = 1; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let value: string | number
      switch (c) {
        case 0:
          value = `ORD-${String(r).padStart(6, "0")}`
          break
        case 1:
          value = regions[r % regions.length]
          break
        case 2:
          value = products[(r * 7) % products.length]
          break
        case 3:
          value = (r * 37) % 250
          break
        case 4:
          value = Math.round(((r * 13) % 9000) / 3) / 100
          break
        case 5:
          value = Math.round(((r * 37) % 250) * (((r * 13) % 9000) / 3)) / 100
          break
        default:
          value = ((r * (c + 3)) % 1000) / 10
      }
      cells.push({ row: r, col: c, data: { value } })
    }
  }
  controller.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "Load", history: false })
}

/** `?demo=formats`: a made-up budget that uses every kind of cell format, for checking the renderer. */
function formatsDemo(controller: SheetController) {
  const cells: CellWrite[] = []
  const put = (ref: string, text: string) => {
    const r = parseRange(ref)!
    const parsed = parseInput(text)
    if (parsed) cells.push({ row: r.r0, col: r.c0, data: { value: parsed.value, input: parsed.input } })
  }
  put("A1", "Household budget")
  put("A3", "Item")
  put("B3", "Planned")
  put("C3", "Actual")
  put("D3", "Difference")
  put("E3", "Share")
  const items: [string, number, number][] = [
    ["Rent", 1200, 1200],
    ["Groceries", 450, 512.4],
    ["Transport", 120, 96.5],
    ["Utilities", 180, 201.75],
    ["Leisure", 200, 145],
  ]
  items.forEach(([name, plan, actual], i) => {
    const r = 4 + i
    put(`A${r}`, name)
    put(`B${r}`, String(plan))
    put(`C${r}`, String(actual))
    put(`D${r}`, `=B${r}-C${r}`)
    put(`E${r}`, `=C${r}/C$10`)
  })
  put("A10", "Total")
  put("B10", "=SUM(B4:B8)")
  put("C10", "=SUM(C4:C8)")
  put("D10", "=SUM(D4:D8)")
  put("G3", "Notes")
  put("G4", "Wrapped text keeps to its column and grows downwards over several lines.")
  put("G7", "Shrinks to fit this narrow cell")
  put("H3", "Tilted")
  put("I3", "Down")
  put("J3", "Stacked")
  put("A12", "Underlined")
  put("B12", "Struck")
  put("C12", "Double")
  put("D12", "Centered across")
  put("A14", "1/15/2026")
  put("B14", "=A14+45")
  put("C14", "0.5")
  put("D14", "-1234.5")
  const ops: Op[] = [{ kind: "setCells", sheet: 0, cells }]
  const f = (ref: string, patch: StylePatch) =>
    ops.push({ kind: "format", sheet: 0, ranges: [parseRange(ref)!], patch })
  const thin = { style: "thin" as const }
  f("A1", { fontSize: 18, bold: true, fontColor: "#1f4e79" })
  f("A3:E3", { bold: true, fill: "#1f4e79", fontColor: "#ffffff", hAlign: "center", borderBottom: { style: "medium" } })
  f("B4:D10", { numberFormat: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' })
  f("D4:D8", { numberFormat: "#,##0.00;[Red]-#,##0.00" })
  f("E4:E8", { numberFormat: "0.0%", italic: true })
  f("A4:E8", { borderLeft: thin, borderRight: thin, borderTop: thin, borderBottom: thin })
  f("A5:E5", { fill: "#f2f2f2" })
  f("A7:E7", { fill: "#f2f2f2" })
  f("A10:E10", { bold: true, borderTop: thin, borderBottom: { style: "double" } })
  f("G3:J3", { bold: true })
  f("G4", { wrap: true, vAlign: "top" })
  f("G7", { shrink: true })
  f("H3", { rotation: 45 })
  f("I3", { rotation: 135 })
  f("J3", { rotation: 255 })
  f("A12", { underline: "single" })
  f("B12", { strike: true })
  f("C12", { underline: "double", fontName: "Georgia" })
  f("D12:F12", { hAlign: "centerContinuous", fill: "#fff2cc" })
  f("A14", { numberFormat: "d mmmm yyyy" })
  f("B14", { numberFormat: "dddd" })
  f("C14", { numberFormat: "# ?/?" })
  f("D14", { numberFormat: "0.00E+00", fontColor: "#7030a0" })
  ops.push({ kind: "merge", sheet: 0, ranges: [parseRange("G4:G6")!] })
  ops.push({ kind: "merge", sheet: 0, ranges: [parseRange("A1:E2")!] })
  f("A1", { hAlign: "center", vAlign: "center" })
  f("G9:J11", { borderTop: { style: "thick", color: "#c00000" }, borderBottom: { style: "dashed" } })
  f("G9:G11", {
    borderLeft: { style: "dotted" },
    borderDiagonal: { style: "thin", color: "#808080" },
    diagonalDown: true,
  })
  controller.wb.transact(ops, { label: "Load", history: false })
  controller.sheet.cols.setSize(0, 110)
  controller.sheet.cols.setSize(1, 96)
  controller.sheet.cols.setSize(2, 96)
  controller.sheet.cols.setSize(3, 96)
  controller.sheet.cols.setSize(6, 120)
  controller.sheet.cols.setSize(7, 40)
  controller.sheet.rows.setSize(2, 40)
  controller.sheet.rows.setSize(11, 24)
}

/** `?demo=cf`: the small demo with one rule of each kind, for checking how conditional formats draw. */
function conditionalDemo(controller: SheetController) {
  const r = (ref: string) => [parseRange(ref)!]
  const style = { fill: "#FFC7CE", fontColor: "#9C0006" }
  const entries: CFEntry[] = [
    {
      id: 1,
      ranges: r("B2:B30"),
      rule: { kind: "text", operator: "contains", text: "north", style },
      stopIfTrue: false,
    },
    {
      id: 2,
      ranges: r("C2:C30"),
      rule: { kind: "duplicate", unique: true, style: { fill: "#C6EFCE", fontColor: "#006100" } },
      stopIfTrue: false,
    },
    {
      id: 3,
      ranges: r("D2:D30"),
      rule: {
        kind: "dataBar",
        min: { type: "automatic" },
        max: { type: "automatic" },
        color: "#638EC6",
        gradient: true,
      },
      stopIfTrue: false,
    },
    {
      id: 4,
      ranges: r("E2:E30"),
      rule: {
        kind: "colorScale",
        points: [
          { at: { type: "min" }, color: "#F8696B" },
          { at: { type: "percentile", value: "50" }, color: "#FFEB84" },
          { at: { type: "max" }, color: "#63BE7B" },
        ],
      },
      stopIfTrue: false,
    },
    {
      id: 5,
      ranges: r("F2:F30"),
      rule: {
        kind: "iconSet",
        set: "3Arrows",
        reverse: false,
        thresholds: [
          { type: "percent", value: "33" },
          { type: "percent", value: "67" },
        ],
        showValue: true,
      },
      stopIfTrue: false,
    },
    {
      id: 6,
      ranges: r("G2:G30"),
      rule: {
        kind: "dataBar",
        min: { type: "automatic" },
        max: { type: "automatic" },
        color: "#63C384",
        gradient: false,
      },
      stopIfTrue: false,
    },
    {
      id: 7,
      ranges: r("H2:H30"),
      rule: {
        kind: "iconSet",
        set: "3TrafficLights1",
        reverse: false,
        thresholds: [
          { type: "percent", value: "33" },
          { type: "percent", value: "67" },
        ],
        showValue: true,
      },
      stopIfTrue: false,
    },
    {
      id: 8,
      ranges: r("I2:I30"),
      rule: { kind: "top", bottom: false, percent: false, rank: 5, style: { fill: "#FFEB9C", fontColor: "#9C5700" } },
      stopIfTrue: false,
    },
  ]
  controller.wb.transact([{ kind: "conditional", sheet: 0, entries }], { label: "Load", history: false })
}

/** `?demo=notes`: the small demo with a shown note, a hidden note and comment threads (made-up names). */
function notesDemo(controller: SheetController) {
  const t = Date.UTC(2026, 10, 2, 9, 30)
  controller.wb.transact(
    [
      {
        kind: "annotations",
        sheet: 0,
        writes: [
          {
            row: 2,
            col: 3,
            value: { kind: "note", author: "Alex Demo", text: "Units from the March count.", shown: true },
          },
          {
            row: 5,
            col: 4,
            value: { kind: "note", author: "Alex Demo", text: "Price before discount.", shown: false },
          },
          {
            row: 8,
            col: 5,
            value: {
              kind: "comment",
              resolved: false,
              posts: [
                { author: "Sam Example", text: "Does this total include tax?", time: t },
                { author: "Alex Demo", text: "No, tax is added on the invoice.", time: t + 3_600_000 },
              ],
            },
          },
          {
            row: 12,
            col: 1,
            value: {
              kind: "comment",
              resolved: true,
              posts: [{ author: "Sam Example", text: "Region renamed?", time: t }],
            },
          },
        ],
      },
    ],
    { label: "Demo" },
  )
  controller.wb.clearHistory()
}

/** Part 8b's objects on the formats sheet: a signature line and embedded files (made-up names, no real files). */
function objectsDemo(c: SheetController) {
  const svg = (markup: string) => `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(markup)))}`
  const at = (row: number, col: number) =>
    c.setSelection({
      ranges: [{ r0: row, c0: col, r1: row, c1: col }],
      active: { row, col },
      activeRange: 0,
      anchor: { row, col },
      focus: { row, col },
    })
  at(16, 0)
  const setup = {
    id: "{00000000-0000-4000-8000-000000000001}",
    signer: "Sam Sample",
    title: "Treasurer",
    email: "sam@example.com",
    instructions: "Check the totals before signing.",
    allowComments: false,
    showDate: true,
  }
  const { w, h } = SIGNATURE_SIZE
  c.objects.insertSignatureLine(setup, svg(signatureDocument(setup, w, h, c.wb.theme)), SIGNATURE_SIZE)
  const files = [
    { name: "Quarterly budget notes.docx", progId: "Word.Document.12" },
    { name: "Receipts.pdf", progId: "Package" },
    { name: "Shopping list.txt", progId: "Package" },
  ]
  files.forEach((f, i) => {
    at(16, 3 + i)
    const embed = { ...f, data: new TextEncoder().encode("Made-up content") }
    const { w: iw, h: ih } = OBJECT_ICON_SIZE
    c.objects.insertObject(embed, svg(embedDocument(embed, iw, ih, c.wb.theme)), OBJECT_ICON_SIZE)
  })
  at(0, 0)
}
