// PivotCharts Excel 365 made (fixtures/excel/pivot-chart.xlsx): Cells reads their PivotTable links and shows the
// points, series names and refs Excel wrote for them, from the PivotTables themselves.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { isChart, pivotChartData, type SheetChart, type Workbook } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}`, import.meta.url)))
async function parts(bytes: Uint8Array) {
  const zip = await readZip(bytes)
  return new Map(zip.map((e) => [e.normalizedName, new TextDecoder().decode(e.data)]))
}

/** Each chart's series as Excel cached them: name and its ref, the categories' ref and levels, the points. */
function excelSeries(xml: string) {
  return [...xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map(([, ser]) => {
    const tx = /<c:tx>[\s\S]*?<c:f>(.*?)<\/c:f>[\s\S]*?<c:v>(.*?)<\/c:v>/.exec(ser)!
    const cat = /<c:cat>([\s\S]*?)<\/c:cat>/.exec(ser)![1]
    const val = /<c:val>[\s\S]*?<c:f>(.*?)<\/c:f>[\s\S]*?<c:ptCount val="(\d+)"\/>([\s\S]*?)<\/c:numCache>/.exec(ser)!
    const count = Number(val[2])
    const values = Array.from({ length: count }, () => null as number | null)
    for (const [, i, v] of val[3].matchAll(/<c:pt idx="(\d+)"><c:v>(.*?)<\/c:v>/g)) values[Number(i)] = Number(v)
    const levels = [...cat.matchAll(/<c:lvl>([\s\S]*?)<\/c:lvl>/g)].map(([, lvl]) => {
      const out = Array.from({ length: count }, () => "")
      for (const [, i, v] of lvl.matchAll(/<c:pt idx="(\d+)"><c:v>(.*?)<\/c:v>/g)) out[Number(i)] = v
      return out
    })
    const plain = [...cat.matchAll(/<c:pt idx="(\d+)"><c:v>(.*?)<\/c:v>/g)].map((m) => m[2])
    return {
      name: tx[2],
      nameRef: tx[1],
      catRef: /<c:f>(.*?)<\/c:f>/.exec(cat)![1],
      valRef: val[1],
      inner: levels.length ? levels[0] : plain,
      outer: levels.slice(1).reverse(),
      values,
    }
  })
}

const chartsOf = (wb: Workbook) =>
  wb.sheets.flatMap((s) => s.objects.filter(isChart).map((c) => ({ sheet: s.name, chart: c as SheetChart })))

describe("PivotCharts", () => {
  it("reads each chart's PivotTable, and shows Excel's points, names and refs from it", async () => {
    const bytes = fixture("pivot-chart.xlsx")
    const { workbook: wb } = await openXlsx(bytes, "pivot-chart.xlsx")
    const zip = await parts(bytes)
    const charts = chartsOf(wb)
    expect(charts.map((c) => c.chart.pivot?.name)).toEqual([
      "PTTwo",
      "PTCross",
      "PTValues",
      "PTPie",
      "PTCollapsed",
      "PTFiltered",
      "PTAway",
    ])
    expect(charts[0].chart.pivot?.buttons).toEqual({
      filter: true,
      categories: true,
      data: true,
      series: false,
      visible: true,
    })
    charts.forEach(({ chart }, i) => {
      const xml = zip.get(`xl/charts/chart${i + 1}.xml`)!
      const want = excelSeries(xml)
      const sheet = wb.sheets.find((s) => s.name === chart.pivot!.sheet)!
      const table = sheet.pivots.find((t) => t.name === chart.pivot!.name)!
      const cache = wb.pivotCaches.find((c) => c.id === table.cacheId)!
      const got = pivotChartData(table, cache, sheet.name)
      const label = `${chart.pivot!.name}`
      expect(got.series.map((s) => s.name), label).toEqual(want.map((w) => w.name))
      expect(got.refs.series.map((s) => s.name), label).toEqual(want.map((w) => w.nameRef))
      expect(got.refs.series.map((s) => s.values), label).toEqual(want.map((w) => w.valRef))
      expect(got.refs.categories, label).toBe(want[0].catRef)
      expect(got.categories, label).toEqual(want[0].inner)
      // Excel writes an outer level's label only where it starts.
      got.levels.forEach((level, k) =>
        level.forEach((l, j) => {
          const w = want[0].outer[k]?.[j]
          if (w) expect(l, `${label} level ${k} at ${j}`).toBe(w)
        }),
      )
      got.series.forEach((s, k) =>
        s.values.forEach((v, j) =>
          expect(s.blanks.includes(j) ? null : v, `${label} ${s.name} ${j}`).toEqual(want[k].values[j]),
        ),
      )
    })
  })

  it("writes PivotCharts back with their PivotTable, refs and multi-level categories", async () => {
    const read = await openXlsx(fixture("pivot-chart.xlsx"), "pivot-chart.xlsx")
    const wb = read.workbook
    // Every chart written anew: a changed title.
    for (const { chart } of chartsOf(wb)) chart.spec = { ...chart.spec, title: { show: true, text: "New" } }
    const saved = await saveXlsx(wb, { source: read.source, fileName: "pivot-chart.xlsx" })
    const zip = await parts(saved)
    const before = await parts(fixture("pivot-chart.xlsx"))
    for (let i = 1; i <= 7; i++) {
      const xml = zip.get(`xl/charts/chart${i}.xml`)!
      const was = before.get(`xl/charts/chart${i}.xml`)!
      expect(xml).toContain(/<c:pivotSource>[\s\S]*?<\/c:pivotSource>/.exec(was)![0])
      const strip = (s: ReturnType<typeof excelSeries>) => s.map(({ name, nameRef, catRef, valRef, inner, values }) => ({ name, nameRef, catRef, valRef, inner, values }))
      expect(strip(excelSeries(xml)), `chart${i}`).toEqual(strip(excelSeries(was)))
    }
    const again = await openXlsx(saved, "pivot-chart.xlsx")
    expect(chartsOf(again.workbook).map((c) => c.chart.pivot)).toEqual(chartsOf(wb).map((c) => c.chart.pivot))
  })
})
