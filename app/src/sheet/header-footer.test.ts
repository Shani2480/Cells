import { openXlsx, saveXlsx } from "@cells/io"
import { describe, expect, it, vi } from "vitest"
import { SheetController } from "./controller"

/** 120 rows in A:C (three pages tall). */
function setup() {
  const c = new SheetController()
  c.viewWidth = 1000
  c.viewHeight = 600
  const cells = []
  for (let row = 0; row < 120; row++) for (let col = 0; col < 3; col++) cells.push({ row, col, data: { value: row } })
  c.wb.transact([{ kind: "setCells", sheet: 0, cells }], { label: "x", before: c.selection })
  c.print.defaultPaper = 1
  c.shell.showTab = vi.fn()
  return c
}

describe("Header & Footer on the page", () => {
  it("Insert › Header & Footer switches to Page Layout and edits the active cell's page", () => {
    const c = setup()
    c.goTo({ r0: 70, c0: 1, r1: 70, c1: 1 })
    c.headerFooter.open("header")
    expect(c.view).toBe("pageLayout")
    const e = c.headerFooter.editing!
    expect(e.where).toBe("header")
    expect(e.rowBand).toBeGreaterThan(0)
    expect(c.shell.showTab).toHaveBeenCalledWith("TabHeaderFooter")
  })

  it("shows codes by name, inserts elements at the caret, and keeps them as codes in one undo step", () => {
    const c = setup()
    c.print.change((s) => (s.header = { left: "", center: "Page &P", right: "" }))
    c.headerFooter.open("header", { rowBand: 0, colBand: 0 }, "center")
    expect(c.headerFooter.editing!.text.center).toBe("Page &[Page]")
    c.headerFooter.type("center", "Page &[Page] of ", 16)
    c.headerFooter.insert("&[Pages]")
    c.headerFooter.place("left", 0)
    c.headerFooter.type("left", "Sales")
    c.headerFooter.commit()
    expect(c.headerFooter.editing).toBeNull()
    expect(c.sheet.pageSetup.header).toEqual({ left: "Sales", center: "Page &P of &N", right: "" })
    c.undo()
    expect(c.sheet.pageSetup.header.center).toBe("Page &P")
  })

  it("with Different First Page and Odd & Even, each page edits the header it prints", () => {
    const c = setup()
    c.headerFooter.open("header", { rowBand: 0, colBand: 0 }, "center")
    c.headerFooter.toggleOption("differentFirst")
    c.headerFooter.toggleOption("differentOddEven")
    expect(c.sheet.pageSetup).toMatchObject({ differentFirst: true, differentOddEven: true })
    // Still editing the first page, now its own header.
    expect(c.headerFooter.editing).toMatchObject({ kind: "first", rowBand: 0 })
    c.headerFooter.type("center", "Cover")
    c.headerFooter.open("header", { rowBand: 1, colBand: 0 }, "center")
    expect(c.headerFooter.editing!.kind).toBe("even")
    c.headerFooter.type("center", "Even")
    c.headerFooter.open("header", { rowBand: 2, colBand: 0 }, "center")
    expect(c.headerFooter.editing!.kind).toBe("odd")
    c.headerFooter.type("center", "Odd")
    c.headerFooter.commit()
    const s = c.sheet.pageSetup
    expect([s.firstHeader.center, s.evenHeader.center, s.header.center]).toEqual(["Cover", "Even", "Odd"])
    // A ready-made header goes to the page's own kind.
    c.headerFooter.open("header", { rowBand: 1, colBand: 0 })
    c.headerFooter.usePreset("header", { left: "", center: "Page &P", right: "" })
    expect(c.sheet.pageSetup.evenHeader.center).toBe("Page &P")
    expect(c.sheet.pageSetup.header.center).toBe("Odd")
  })

  it("Picture puts &[Picture] at the caret and keeps the picture with the text; deleting the code drops it", () => {
    const c = setup()
    const confirm = vi.fn((_message: string, ok: () => void) => ok())
    c.shell.confirm = confirm
    const logo = { src: "data:image/png;base64,AAAA", w: 200, h: 100, name: "logo.png" }
    c.headerFooter.open("header", { rowBand: 0, colBand: 0 }, "center")
    c.headerFooter.type("center", "Acme ", 5)
    c.headerFooter.insertPicture(logo)
    expect(c.headerFooter.editing!.text.center).toBe("Acme &[Picture]")
    expect(c.headerFooter.currentPicture()).toMatchObject({ title: "logo", width: 150, height: 75 })
    // Format Picture's settings go with it.
    c.headerFooter.formatPicture({ ...c.headerFooter.currentPicture()!, color: "grayscale", width: 75, height: 37.5 })
    c.headerFooter.commit()
    const setup1 = c.sheet.pageSetup
    expect(setup1.header.center).toBe("Acme &G")
    expect(setup1.pictures.CH).toMatchObject({ color: "grayscale", width: 75 })
    // A second picture in the same section asks first.
    c.headerFooter.open("header", { rowBand: 0, colBand: 0 }, "center")
    c.headerFooter.insertPicture({ ...logo, name: "other.png" })
    expect(confirm).toHaveBeenCalledOnce()
    expect(c.headerFooter.editing!.text.center).toBe("Acme &[Picture]")
    c.headerFooter.type("center", "Acme ")
    c.headerFooter.commit()
    expect(c.sheet.pageSetup.pictures).toEqual({})
    // Undo brings the picture back with its text.
    c.undo()
    expect(c.sheet.pageSetup.pictures.CH).toMatchObject({ color: "grayscale" })
  })

  it("Esc drops the text; Go to Footer keeps what was typed in the header", () => {
    const c = setup()
    c.headerFooter.open("header")
    c.headerFooter.type("right", "Draft")
    c.headerFooter.cancel()
    expect(c.sheet.pageSetup.header.right).toBe("")
    c.headerFooter.open("header")
    c.headerFooter.type("right", "Draft")
    c.headerFooter.goTo("footer")
    expect(c.headerFooter.editing!.where).toBe("footer")
    expect(c.sheet.pageSetup.header.right).toBe("Draft")
  })

  it("a ready-made footer applies at once; leaving Page Layout or the sheet keeps the text", () => {
    const c = setup()
    c.headerFooter.open("footer")
    c.headerFooter.usePreset("footer", { left: "", center: "&A", right: "Page &P" })
    expect(c.sheet.pageSetup.footer.right).toBe("Page &P")
    expect(c.headerFooter.editing!.text.right).toBe("Page &[Page]")
    c.headerFooter.type("left", "Mock Co")
    c.setView("normal")
    expect(c.headerFooter.editing).toBeNull()
    expect(c.sheet.pageSetup.footer.left).toBe("Mock Co")
  })

  it("the ruler setting is kept per sheet and saved in Excel files", async () => {
    const c = setup()
    c.setView("pageLayout")
    c.showRuler = false
    const bytes = await saveXlsx(c.wb, { views: c.saveViews(), now: new Date(Date.UTC(2026, 0, 1)) })
    const back = await openXlsx(bytes, "Ruler.xlsx")
    expect(back.views[0]).toMatchObject({ mode: "pageLayout", showRuler: false })
    const d = new SheetController(back.workbook)
    d.loadViews(back.views, 0)
    expect(d.showRuler).toBe(false)
    // Shown (Excel's default) isn't written.
    c.showRuler = true
    const again = await openXlsx(await saveXlsx(c.wb, { views: c.saveViews() }), "R.xlsx")
    expect(again.views[0].showRuler).toBeUndefined()
  })
})
