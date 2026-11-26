// Pages of a sheet with its page setup: rows and columns split to the paper inside the margins, manual breaks, titles
// repeated, page order, Fit to and the print area.
import { describe, expect, it } from "vitest"
import { defaultPageSetup, paginate, pageBreaks, Workbook } from "../src"

/** A sheet with a number in every cell of rows × cols (default sizes: 20 px rows, 64 px columns). */
function sheet(rows: number, cols: number) {
  const wb = new Workbook(["Data"])
  const s = wb.sheets[0]
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) s.cells.set(r, c, { value: r * cols + c })
  return s
}

describe("pagination", () => {
  it("Letter portrait, Normal margins: 45 rows of 15 pt and 10 columns of 48 pt a page", () => {
    // Inside the margins: 7.1 × 9.5 inches = 511.2 × 684 pt.
    const s = sheet(100, 12)
    const p = paginate(s, defaultPageSetup())
    expect(p.body).toEqual({ x: 50.4, y: 54, width: 511.2, height: 684 })
    expect(p.pages.map((pg) => [pg.rows, pg.cols])).toEqual([
      [
        [0, 44],
        [0, 9],
      ],
      [
        [45, 89],
        [0, 9],
      ],
      [
        [90, 99],
        [0, 9],
      ],
      [
        [0, 44],
        [10, 11],
      ],
      [
        [45, 89],
        [10, 11],
      ],
      [
        [90, 99],
        [10, 11],
      ],
    ])
    expect(pageBreaks(p)).toEqual({ rows: [45, 90, 100], cols: [10, 12] })
  })

  it("across first, then down; numbered from the first page number", () => {
    const s = sheet(60, 12)
    const setup = { ...defaultPageSetup(), order: "overThenDown" as const, firstPageNumber: 5 }
    const p = paginate(s, setup)
    expect(p.pages.map((pg) => [pg.rows[0], pg.cols[0], pg.number])).toEqual([
      [0, 0, 5],
      [0, 10, 6],
      [45, 0, 7],
      [45, 10, 8],
    ])
  })

  it("manual breaks, hidden rows and titles that repeat", () => {
    const s = sheet(100, 3)
    s.rows.setHidden(5, true)
    const setup = { ...defaultPageSetup(), rowBreaks: [20], titleRows: [0, 1] as [number, number] }
    const p = paginate(s, setup)
    expect(p.pages.map((pg) => [pg.rows, pg.titleRows])).toEqual([
      [[0, 19], null],
      // 684 pt holds 45 rows; the titles take 2.
      [
        [20, 62],
        [0, 1],
      ],
      [
        [63, 99],
        [0, 1],
      ],
    ])
  })

  it("Fit to 1 page wide picks the largest scale that fits; the scale shrinks the rows too", () => {
    const s = sheet(40, 20)
    const setup = { ...defaultPageSetup(), fitTo: { width: 1, height: 0 } }
    const p = paginate(s, setup)
    // 20 columns are 960 pt; 511.2 pt fits them at 53%.
    expect(p.scale).toBe(0.53)
    expect(p.pages).toHaveLength(1)
    const scaled = paginate(s, { ...defaultPageSetup(), scale: 50 })
    expect(scaled.pages).toHaveLength(1)
  })

  it("prints the print area's ranges each on pages of their own, or the cells in use", () => {
    const s = sheet(10, 3)
    s.cells.set(30, 6, { value: "far" })
    expect(paginate(s, defaultPageSetup()).areas).toEqual([{ r0: 0, c0: 0, r1: 30, c1: 6 }])
    const setup = {
      ...defaultPageSetup(),
      printArea: [
        { r0: 0, c0: 0, r1: 4, c1: 1 },
        { r0: 30, c0: 6, r1: 30, c1: 6 },
      ],
    }
    const p = paginate(s, setup)
    expect(p.pages.map((pg) => [pg.area, pg.rows, pg.cols])).toEqual([
      [0, [0, 4], [0, 1]],
      [1, [30, 30], [6, 6]],
    ])
    expect(paginate(s, setup, { ignorePrintArea: true }).areas).toEqual([{ r0: 0, c0: 0, r1: 30, c1: 6 }])
    expect(paginate(new Workbook(["Empty"]).sheets[0], defaultPageSetup()).pages).toEqual([])
  })

  it("landscape A4 turns the paper", () => {
    const p = paginate(sheet(1, 1), { ...defaultPageSetup(), paperSize: 9, orientation: "landscape" })
    expect(p.paper.width).toBeCloseTo(841.89, 1)
    expect(p.paper.height).toBeCloseTo(595.28, 1)
  })

  it("a chart sheet prints on one page, its chart filling the space inside the margins", () => {
    const wb = new Workbook()
    const sheet = wb.sheets[0]
    sheet.chartSheet = true
    const setup = { ...defaultPageSetup(), orientation: "landscape" as const, firstPageNumber: 4 }
    const p = paginate(sheet, setup, { defaultPaper: 1 })
    expect(p.pages).toEqual([{ area: 0, rows: [0, 0], cols: [0, 0], titleRows: null, titleCols: null, number: 4 }])
    expect(p.paper).toEqual({ width: 792, height: 612 })
    expect(p.body).toEqual({ x: 0.7 * 72, y: 0.75 * 72, width: 792 - 1.4 * 72, height: 612 - 1.5 * 72 })
  })
})
