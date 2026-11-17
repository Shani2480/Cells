// Excel 97-2003 workbooks (.xls): files LibreOffice wrote from the made-up samples, read as their .xlsx twins read,
// and small files built here for the records LibreOffice doesn't write.
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import type { Workbook } from "@cells/core"
import { isCompoundFile, openXls, openXlsx, XlsError } from "../src"
import { rkValue } from "../src/xls/read"

const xls = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/xls/${name}.xls`, import.meta.url)))
const xlsx = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/excel/${name}.xlsx`, import.meta.url)))

/** Each cell's value and formula, by sheet, for the cells .xls can hold (65,536 rows, 256 columns). */
function contents(wb: Workbook) {
  return wb.sheets.map((s) => {
    const out: Record<string, unknown> = {}
    const used = s.cells.usedBounds()
    if (used)
      s.cells.forEach(
        { ...used, r1: Math.min(used.r1, 65535), c1: Math.min(used.c1, 255) },
        (r, c, d) => (out[`${r},${c}`] = [d.value, d.input]),
      )
    return out
  })
}

// What LibreOffice changes on the way: dynamic arrays and table references don't exist in .xls, and it saves plain
// TRUE and error cells as formulas.
const LIBREOFFICE: Record<string, string[]> = {
  budget: ["0,4,0", "0,4,1", "0,7,0"],
  features: ["0,6,3", "0,1,5"],
}

describe.each(["budget", "formats", "layout", "features"])("LibreOffice's %s.xls", (name) => {
  it("reads as its .xlsx twin", async () => {
    const a = await openXls(xls(name), `${name}.xls`)
    const b = await openXlsx(xlsx(name), `${name}.xlsx`)
    expect(a.workbook.sheets.map((s) => [s.name, s.hidden])).toEqual(b.workbook.sheets.map((s) => [s.name, s.hidden]))
    const [x, y] = [contents(a.workbook), contents(b.workbook)]
    for (const key of LIBREOFFICE[name] ?? []) {
      const [sheet, ...cell] = key.split(",")
      delete x[Number(sheet)][cell.join(",")]
      delete y[Number(sheet)][cell.join(",")]
    }
    expect(x).toEqual(y)
    expect(a.workbook.names.all()).toEqual(b.workbook.names.all())
    a.workbook.sheets.forEach((s, i) => {
      const o = b.workbook.sheets[i]
      expect(s.merges.all.map((m) => ({ ...m }))).toEqual(o.merges.all.map((m) => ({ ...m })))
      expect(s.annotations.all()).toEqual(o.annotations.all().filter((n) => n.value.kind === "note"))
      expect(a.views[i].freeze).toEqual(b.views[i].freeze)
      for (let c = 0; c < 8; c++) {
        expect(s.cols.isHidden(c)).toBe(o.cols.isHidden(c))
        expect(Math.abs(s.cols.size(c) - o.cols.size(c))).toBeLessThanOrEqual(2)
      }
      for (let r = 0; r < 20; r++) expect(s.rows.isHidden(r)).toBe(o.rows.isHidden(r))
    })
  })
})

describe("formats.xls", () => {
  it("keeps fonts, fills, borders, alignment, number formats and cell styles", async () => {
    const a = await openXls(xls("formats"))
    const b = await openXlsx(xlsx("formats"))
    const s = a.workbook.sheets[0]
    const o = b.workbook.sheets[0]
    let same = 0
    let all = 0
    o.cells.forEach(o.cells.usedBounds()!, (r, c) => {
      all++
      const x = a.workbook.styles.get(s.styleAt(r, c))
      const y = b.workbook.styles.get(o.styleAt(r, c))
      // LibreOffice spells some number formats its own way and sets the alignment of turned text.
      const { numberFormat: _x, hAlign: _hx, ...xs } = x
      const { numberFormat: _y, hAlign: _hy, ...ys } = y
      if (JSON.stringify(Object.entries(xs).sort()) === JSON.stringify(Object.entries(ys).sort())) same++
    })
    expect(same).toBe(all)
    expect(a.workbook.styles.get(s.styleAt(16, 0)).cellStyle).toBe("Good")
  })
})

describe("corpus.xls", () => {
  it("decodes shared formulas, functions, names, other sheets and arrays, and calculates them", async () => {
    const { workbook: wb, warnings } = await openXls(xls("corpus"), "corpus.xls")
    expect(warnings).toEqual([])
    expect(wb.sheets.map((s) => s.name)).toEqual(["Data", "Q1 Sales", "Long"])
    expect(wb.names.all()).toEqual([{ name: "Rate", ref: "Data!$B$1" }])
    const s = wb.sheets[0]
    const f = (r: number, c: number) => s.cells.input(r, c)
    const v = (r: number, c: number) => s.cells.value(r, c)
    expect([f(0, 2), f(5, 2), f(5, 3)]).toEqual(["=A1*B1", "=A6*B6", '=IF(A6>3,"big","small")&"!"'])
    expect([f(7, 1), f(7, 2), f(7, 3)]).toEqual([
      "=EOMONTH(A8,1)",
      '=IFERROR(1/0,"none")',
      "='Q1 Sales'!B2+SUM('Q1 Sales'!B2:B4)",
    ])
    expect([f(8, 1), f(8, 2), f(8, 3)]).toEqual([
      "=SUMPRODUCT((A1:A6>2)*(B1:B6))",
      "=ROUND(-A2^2,2)%",
      "=SUM({1,2;3,4})",
    ])
    expect([f(9, 0), f(9, 1), f(9, 2), f(9, 3)]).toEqual([
      "=VLOOKUP(4.5,A1:C6,3,FALSE)",
      '=COUNTIF(A:A,">2")',
      "=Rate*2",
      '=TEXT(A8,"dd mmm yyyy")',
    ])
    // Calculated on opening, as Excel does with an .xls (LibreOffice saved no text results).
    expect([v(5, 3), v(7, 1), v(7, 2), v(7, 3), v(8, 1), v(8, 3), v(9, 0), v(9, 2), v(9, 3)]).toEqual([
      "big!",
      46053,
      "none",
      110,
      4000098,
      10,
      4500094.5,
      2000000,
      "09 Dec 2025",
    ])
    expect(v(8, 0)).toBe("Ünïcødé — 日本語 текст")
    const long = wb.sheets[2]
    expect(long.cells.value(0, 0)).toBe("строка номер 0")
    expect(long.cells.value(897, 0)).toBe("строка номер 897")
    expect(long.cells.value(899, 0)).toBe("row text number 899")
  })
})

// ── Files built here ────────────────────────────────────────────────────────────────────────────────────────────

/** One BIFF record. */
function rec(type: number, ...parts: (number[] | Uint8Array)[]): number[] {
  const data = parts.flatMap((p) => Array.from(p))
  return [type & 0xff, type >> 8, data.length & 0xff, data.length >> 8, ...data]
}
const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff]
const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff]
const f64 = (n: number) => Array.from(new Uint8Array(new Float64Array([n]).buffer))
const str = (s: string) => [...u16(s.length), 0, ...Array.from(s, (ch) => ch.charCodeAt(0))]
const str8 = (s: string) => [s.length, 0, ...Array.from(s, (ch) => ch.charCodeAt(0))]

/** A workbook stream with one sheet holding `cells` (and `globals` records first). */
function workbookStream(cells: number[][], globals: number[][] = [], version = 0x0600): Uint8Array {
  const bof = (dt: number) => rec(0x0809, u16(version), u16(dt), u16(0), u16(0), u32(0), u32(0))
  const xf = rec(0x00e0, u16(0), u16(0), u16(0x0001), [0x20, 0, 0, 0], u32(0), u32(0), u16(0x20c0))
  const font = rec(0x0031, u16(200), u16(0), u16(0x7fff), u16(400), u16(0), [0, 0, 0, 0], str8("Arial"))
  const head = [bof(5), ...globals, font, ...Array(16).fill(xf)]
  const sheetName = "Sheet1"
  const boundLength = 4 + 4 + 2 + 2 + sheetName.length
  const before = head.reduce((n, r) => n + r.length, 0) + boundLength + 4 // + EOF
  const bound = rec(0x0085, u32(before), [0, 0], str8(sheetName))
  const out = [...head, bound, rec(0x000a), bof(0x10), ...cells, rec(0x000a)].flat()
  return new Uint8Array(out)
}

/** The smallest compound file holding `streams` (version 3, 512-byte sectors, small streams in the mini stream). */
function compoundFile(streams: Record<string, Uint8Array>): Uint8Array {
  const SECTOR = 512
  const names = Object.keys(streams)
  const big = names.filter((n) => streams[n].length >= 4096)
  const small = names.filter((n) => streams[n].length < 4096)
  // Mini stream: 64-byte sectors.
  const mini: number[] = []
  const miniFat: number[] = []
  const miniStart: Record<string, number> = {}
  for (const n of small) {
    const data = Array.from(streams[n])
    const count = Math.max(1, Math.ceil(data.length / 64))
    miniStart[n] = mini.length / 64
    for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? 0xfffffffe : miniStart[n] + i + 1)
    mini.push(...data, ...Array(count * 64 - data.length).fill(0))
  }
  const sectors: number[][] = []
  const fat: number[] = []
  const chain = (bytes: number[]) => {
    const start = sectors.length
    const count = Math.max(1, Math.ceil(bytes.length / SECTOR))
    for (let i = 0; i < count; i++) {
      sectors.push([...bytes.slice(i * SECTOR, (i + 1) * SECTOR), ...Array(SECTOR).fill(0)].slice(0, SECTOR))
      fat.push(i === count - 1 ? 0xfffffffe : start + i + 1)
    }
    return start
  }
  const bigStart: Record<string, number> = {}
  for (const n of big) bigStart[n] = chain(Array.from(streams[n]))
  const miniStreamStart = mini.length ? chain(mini) : 0xfffffffe
  const miniFatStart = miniFat.length ? chain(miniFat.flatMap(u32)) : 0xfffffffe
  const entry = (name: string, type: number, child: number, right: number, start: number, size: number) => {
    const n = Array.from(name).flatMap((ch) => u16(ch.charCodeAt(0)))
    return [
      ...[...n, 0, 0, ...Array(64).fill(0)].slice(0, 64),
      ...u16((name.length + 1) * 2),
      type,
      1,
      ...u32(0xffffffff),
      ...u32(right),
      ...u32(child),
      ...Array(16 + 4 + 16).fill(0),
      ...u32(start),
      ...u32(size),
      ...u32(0),
    ]
  }
  const dir = [
    entry("Root Entry", 5, names.length ? 1 : 0xffffffff, 0xffffffff, miniStreamStart, mini.length),
    ...names.map((n, i) =>
      entry(
        n,
        2,
        0xffffffff,
        i + 1 < names.length ? i + 2 : 0xffffffff,
        streams[n].length >= 4096 ? bigStart[n] : miniStart[n],
        streams[n].length,
      ),
    ),
  ].flat()
  const dirStart = chain(dir)
  // The FAT itself: one sector covers 128 entries.
  const fatSector = sectors.length
  fat.push(0xfffffffd)
  sectors.push([...fat.flatMap(u32), ...Array(SECTOR).fill(0xff)].slice(0, SECTOR))
  const header = [
    0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
    ...Array(16).fill(0),
    ...u16(0x3e), ...u16(3), ...u16(0xfffe), ...u16(9), ...u16(6),
    ...Array(6).fill(0),
    ...u32(0), ...u32(1), ...u32(dirStart), ...u32(0), ...u32(4096),
    ...u32(miniFatStart), ...u32(miniFat.length ? 1 : 0), ...u32(0xfffffffe), ...u32(0),
    ...u32(fatSector),
    ...Array(108 * 4).fill(0xff),
  ] // prettier-ignore
  return new Uint8Array([...header, ...sectors.flat()])
}

const book = (cells: number[][], globals: number[][] = [], version = 0x0600) =>
  compoundFile({ Workbook: workbookStream(cells, globals, version) })

describe("records LibreOffice doesn't write", () => {
  it("RK numbers: integers and doubles, either divided by 100", () => {
    expect(rkValue((123 << 2) | 0x02)).toBe(123)
    expect(rkValue((-5 << 2) | 0x02)).toBe(-5)
    expect(rkValue((1234 << 2) | 0x03)).toBe(12.34)
    // 1.5's top 32 bits.
    expect(rkValue(0x3ff80000)).toBe(1.5)
    expect(rkValue(0x3ff80000 | 0x01)).toBe(0.015)
  })

  it("reads NUMBER, RK, BOOLERR, LABEL, BLANK and a formula's text result (STRING)", async () => {
    const cell = (row: number, col: number) => [...u16(row), ...u16(col), ...u16(15)]
    const file = book([
      rec(0x0203, cell(0, 0), f64(3.25)),
      rec(0x027e, cell(0, 1), u32((42 << 2) | 0x02)),
      rec(0x0205, cell(1, 0), [1, 0]),
      rec(0x0205, cell(1, 1), [0x07, 1]),
      rec(0x0204, cell(2, 0), str("=not a formula")),
      rec(0x0201, cell(2, 1)),
      // ="a"&"b", its cached result in a STRING record.
      rec(
        0x0006,
        cell(3, 0),
        [0, 0, 0, 0, 0, 0, 0xff, 0xff],
        u16(0),
        u32(0),
        u16(9),
        [0x17, 1, 0, 0x61],
        [0x17, 1, 0, 0x62],
        [0x08],
      ),
      rec(0x0207, str("ab")),
    ])
    expect(isCompoundFile(file)).toBe(true)
    const { workbook: wb } = await openXls(file, "small.xls", { calculate: false })
    const s = wb.sheets[0]
    expect(s.cells.value(0, 0)).toBe(3.25)
    expect(s.cells.value(0, 1)).toBe(42)
    expect(s.cells.value(1, 0)).toBe(true)
    expect(s.cells.value(1, 1)).toEqual({ error: "#DIV/0!" })
    expect([s.cells.value(2, 0), s.cells.input(2, 0)]).toEqual(["=not a formula", "'=not a formula"])
    expect(s.cells.get(2, 1)).toBeUndefined()
    expect([s.cells.value(3, 0), s.cells.input(3, 0)]).toEqual(["ab", '="a"&"b"'])
  })

  it("explains password-protected, Excel 95 and damaged files", async () => {
    const kind = (bytes: Uint8Array) =>
      openXls(bytes).then(
        () => "opened",
        (e: unknown) => (e instanceof XlsError ? e.kind : String(e)),
      )
    expect(await kind(book([], [rec(0x002f, u16(1))]))).toBe("encrypted")
    expect(await kind(book([], [], 0x0500))).toBe("old")
    expect(await kind(compoundFile({ Book: new Uint8Array(64) }))).toBe("old")
    expect(await kind(compoundFile({ Other: new Uint8Array(64) }))).toBe("damaged")
    expect(await kind(new Uint8Array(600))).toBe("damaged")
  })
})

describe("outlines in .xls", () => {
  it("reads row and column groups, the collapsed mark and where summaries sit", async () => {
    const s = (await openXls(xls("outline"), "outline.xls")).workbook.sheets[0]
    expect([0, 1, 2, 3, 4, 5, 6].map((r) => s.rows.level(r))).toEqual([0, 2, 2, 1, 2, 1, 0])
    expect([s.rows.isHidden(4), s.rows.isCollapsed(5), s.cols.level(1)]).toEqual([true, true, 1])
    // LibreOffice writes summaries to the right whatever the .xlsx said (WSBOOL 0xC1).
    expect(s.outline).toEqual({ summaryBelow: true, summaryRight: true, showSymbols: true })
  })
})

describe("hyperlinks in .xls", () => {
  it("reads web, mail, file and place links as Excel wrote them, with ScreenTips", async () => {
    const read = await openXls(xls("hyperlinks"), "hyperlinks.xls")
    const twin = await openXlsx(xlsx("hyperlinks"), "hyperlinks.xlsx")
    const plain = (links: Workbook["sheets"][number]["hyperlinks"]) => links.map(({ display: _, ...l }) => l)
    expect(plain(read.workbook.sheets[0].hyperlinks)).toEqual(plain(twin.workbook.sheets[0].hyperlinks))
    expect(read.warnings.join(" ")).not.toContain("hyperlinks")
  })
})
