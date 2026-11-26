// Freeforms as Excel 365 saved them (fixtures/excel/freeform.xlsx): a closed triangle of straight lines
// (Freeform 1) and an open curve (Freeform 2), both custGeom.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { isChart, type SheetObject, type SheetShape } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { readCustGeom } from "../src/xlsx/cust-geom"
import { parseTree } from "../src/xlsx/dom"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(new URL("./fixtures/excel/freeform.xlsx", import.meta.url)))
const shapes = (list: SheetObject[]) => list.filter((o): o is SheetShape => !isChart(o) && o.kind === "shape")
const drawing = async (bytes: Uint8Array) =>
  new TextDecoder().decode((await readZip(bytes)).find((e) => e.normalizedName === "xl/drawings/drawing1.xml")!.data)

describe("freeforms (custGeom) in xlsx", () => {
  it("reads Excel's points, lines and curves", async () => {
    const [tri, curve] = shapes((await openXlsx(fixture(), "f.xlsx")).workbook.sheets[0].objects)
    expect(tri.geometry).toBe("custom")
    expect(tri.custom).toEqual([
      {
        w: 1905001,
        h: 1397001,
        commands: [
          { t: "M", x: 0, y: 0 },
          { t: "L", x: 1905000, y: 127000 },
          { t: "L", x: 1270000, y: 1397000 },
          { t: "Z" },
        ],
      },
    ])
    expect(curve.custom![0].commands[1]).toEqual({
      t: "C",
      x1: 508000,
      y1: 0,
      x2: 1270000,
      y2: 1270000,
      x: 2032000,
      y: 508000,
    })
  })

  it("writes a changed freeform's own outline, not a rectangle", async () => {
    const { workbook, source } = await openXlsx(fixture(), "f.xlsx")
    const tri = shapes(workbook.sheets[0].objects)[0]
    tri.custom![0].commands[1] = { t: "L", x: 1905000, y: 700000 }
    const saved = await saveXlsx(workbook, { source, fileName: "f.xlsx" })
    const xml = await drawing(saved)
    expect(xml).not.toContain(`prst="rect"`)
    const back = shapes((await openXlsx(saved, "f.xlsx")).workbook.sheets[0].objects)[0]
    const [p] = back.custom!
    const m = p.commands[1] as { x: number; y: number }
    // Written in the shape's size in EMUs: the point keeps its place in the box.
    expect(m.x / p.w).toBeCloseTo(1905000 / 1905001, 4)
    expect(m.y / p.h).toBeCloseTo(700000 / 1397001, 4)
    expect(p.commands.at(-1)).toEqual({ t: "Z" })
  })

  it("arcs become curves; guides it can't work out keep the file's own XML", () => {
    const geom = (path: string) =>
      readCustGeom(
        parseTree(
          `<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pathLst>${path}</a:pathLst></a:custGeom>`,
        ).root,
      )
    const arc = geom(
      `<a:path w="200" h="200"><a:moveTo><a:pt x="200" y="100"/></a:moveTo><a:arcTo wR="100" hR="100" stAng="0" swAng="10800000"/></a:path>`,
    )!
    const end = arc[0].commands.at(-1) as { x: number; y: number }
    expect(arc[0].commands.filter((c) => c.t === "C")).toHaveLength(2)
    expect(end.x).toBeCloseTo(0)
    expect(end.y).toBeCloseTo(100)
    expect(geom(`<a:path w="1" h="1"><a:moveTo><a:pt x="w" y="0"/></a:moveTo></a:path>`)).toBeNull()
  })
})
