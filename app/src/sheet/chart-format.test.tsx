import { act } from "react"
import { createRoot } from "react-dom/client"
import { afterEach, describe, expect, it } from "vitest"
import { isChart, parseInput, selectRange } from "@cells/core"
import { openXlsx, saveXlsx } from "@cells/io"
import { ChartFormatPane } from "@/shell/ChartFormatPane"
import { InsertChartDialog } from "@/shell/InsertChartDialog"
import { recommendCharts } from "./chart-recommend"
import { SheetController } from "./controller"

/** Month, North and South in A1:C5 (mock data), the selection on it. */
function withData(
  rows = [
    ["Month", "North", "South"],
    ["Jan", "12", "8"],
    ["Feb", "19", "11"],
    ["Mar", "15", "14"],
    ["Apr", "23", "9"],
  ],
) {
  const c = new SheetController()
  c.wb.transact(
    [
      {
        kind: "setCells",
        sheet: 0,
        cells: rows.flatMap((r, row) => r.map((t, col) => ({ row, col, data: { value: parseInput(t)!.value } }))),
      },
    ],
    { label: "Load" },
  )
  c.setSelection(selectRange({ r0: 0, c0: 0, r1: rows.length - 1, c1: rows[0].length - 1 }))
  return c
}

let cleanup: (() => void) | null = null
afterEach(() => {
  cleanup?.()
  cleanup = null
  document.body.innerHTML = ""
})

async function show(node: React.ReactNode) {
  const host = document.createElement("div")
  document.body.append(host)
  const root = createRoot(host)
  await act(async () => root.render(node))
  cleanup = () => root.unmount()
}

const click = async (el: Element | undefined) => {
  expect(el).toBeTruthy()
  await act(async () => (el as HTMLElement).click())
}

describe("chart Format: the picked part", () => {
  it("formats only the picked part, resets it to the style, and undoes as one step each", () => {
    const c = withData()
    c.charts.insert()
    expect(c.charts.part).toBe("chartArea")
    c.charts.setPart("legend")
    c.charts.formatPart({ color: "#c00000", bold: true }, "Text Fill")
    c.charts.setPart("series:1")
    c.charts.formatPart({ fill: "#ffc000" }, "Shape Fill")
    const spec = c.charts.spec()!
    expect(spec.format.legend).toEqual({ color: "#c00000", bold: true })
    expect(spec.format["series:1"]).toEqual({ fill: "#ffc000" })
    c.charts.resetPart()
    expect(c.charts.spec()!.format["series:1"]).toBeUndefined()
    c.undo()
    expect(c.charts.spec()!.format["series:1"]).toEqual({ fill: "#ffc000" })
    c.undo()
    c.undo()
    expect(c.charts.spec()!.format.legend).toBeUndefined()
  })

  it("goes back to Chart Area for another chart, or when the part is gone", () => {
    const c = withData()
    const first = c.charts.insert()!
    c.charts.setPart("title")
    expect(c.charts.part).toBe("title")
    c.charts.change((s) => ({ ...s, title: { ...s.title, show: false } }))
    expect(c.charts.part).toBe("chartArea")
    c.charts.setPart("legend")
    c.setSelection(selectRange({ r0: 0, c0: 0, r1: 4, c1: 2 }))
    c.charts.insert("line")
    expect(c.charts.part).toBe("chartArea")
    c.charts.select(first.id)
    expect(c.charts.part).toBe("chartArea")
  })

  it("keeps the parts' formats through a save and reopen", async () => {
    const c = withData()
    c.charts.insert()
    c.charts.setPart("title")
    c.charts.formatPart({ color: "#2f5597", fontSize: 18 })
    c.charts.setPart("plotArea")
    c.charts.formatPart({ fill: "#deebf7", line: "#2f5597", lineWidth: 1.5 })
    const { workbook } = await openXlsx(await saveXlsx(c.wb), "x.xlsx")
    const spec = workbook.sheets[0].charts[0].spec as { format: Record<string, unknown> }
    expect(spec.format.title).toEqual({ color: "#2f5597", fontSize: 18 })
    expect(spec.format.plotArea).toEqual({ fill: "#deebf7", line: "#2f5597", lineWidth: 1.5 })
  })

  it("the Format pane offers what the part can take and changes it", async () => {
    const c = withData()
    c.charts.insert()
    c.charts.openFormatPane("legend")
    await show(<ChartFormatPane controller={c} onDone={() => {}} />)
    const headings = [...document.body.querySelectorAll("summary")].map((s) => s.textContent)
    expect(headings).toEqual(["Font", "Legend Options"])
    await click(
      [...document.body.querySelectorAll("label")].find((l) => l.textContent === "Top")?.querySelector("input") ??
        undefined,
    )
    expect(c.charts.spec()!.legend).toBe("top")
    cleanup?.()
    c.charts.setPart("series:0")
    await show(<ChartFormatPane controller={c} onDone={() => {}} />)
    expect([...document.body.querySelectorAll("summary")].map((s) => s.textContent)).toEqual([
      "Fill",
      "Line",
      "Series Options",
    ])
  })
})

describe("Recommended Charts", () => {
  it("suits the data: pairs of numbers, one small series, series of very different sizes, a long run", () => {
    expect(recommendCharts(null)[0]).toBe("column")
    expect(recommendCharts({ categories: ["1", "2"], series: [{ values: [3, 4] }], numericX: true })[0]).toBe("scatter")
    const one = recommendCharts({ categories: ["North", "South", "East"], series: [{ values: [5, 3, 2] }] })
    expect(one.slice(0, 4)).toEqual(["column", "hbar", "pie", "doughnut"])
    const apart = recommendCharts({
      categories: ["A", "B", "C"],
      series: [{ values: [1200, 1500, 900] }, { values: [0.1, 0.4, 0.2] }],
    })
    expect(apart[0]).toBe("combo-column-line-secondary")
    const months = recommendCharts({
      categories: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"],
      series: [{ values: Array(13).fill(4) }],
    })
    expect(months[0]).toBe("line")
    expect(new Set(one).size).toBe(one.length)
  })
})

describe("Insert Chart dialog", () => {
  it("inserts the recommended chart picked, drawn from the selected data", async () => {
    const c = withData()
    let closed = false
    await show(<InsertChartDialog controller={c} onClose={() => (closed = true)} />)
    const options = [...document.body.querySelectorAll('[role="option"]')].map((o) => o.getAttribute("aria-label"))
    expect(options[0]).toBe("Clustered Column")
    await click(document.body.querySelector('[aria-label="Line"][role="option"]') ?? undefined)
    await click([...document.body.querySelectorAll("button")].find((b) => b.textContent === "OK"))
    expect(closed).toBe(true)
    const chart = c.sheet.objects.find(isChart)!
    expect(c.charts.liveSpec(chart).type).toBe("line")
    expect(chart.refs?.series.map((s) => s.values)).toEqual(["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"])
  })

  it("All Charts opens on a family, and Change Chart Type changes the selected chart", async () => {
    const c = withData()
    c.charts.insert()
    await show(<InsertChartDialog controller={c} change start={{ family: "bar" }} onClose={() => {}} />)
    expect(document.body.querySelector('[role="tab"][aria-selected="true"]')?.textContent).toBe("All Charts")
    expect(document.body.querySelector(".ic-name")?.textContent).toBe("Clustered Bar")
    await click(document.body.querySelector('[aria-label="Stacked Bar"]') ?? undefined)
    await click([...document.body.querySelectorAll("button")].find((b) => b.textContent === "OK"))
    expect(c.charts.variant()).toBe("hbar-stacked")
    expect(c.sheet.charts).toHaveLength(1)
  })
})
