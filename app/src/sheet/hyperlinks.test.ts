import { describe, expect, it, vi } from "vitest"
import { DEFAULT_PROTECTION } from "@cells/core"
import { SheetController } from "./controller"
import { BAD_REFERENCE, resolvePath, type LinkTarget } from "./hyperlinks"

const at = (r0: number, c0: number, r1 = r0, c1 = c0) => ({ r0, c0, r1, c1 })

function setup() {
  const c = new SheetController()
  c.viewWidth = 1000
  c.viewHeight = 600
  const followed: LinkTarget[] = []
  const alerts: string[] = []
  c.shell.followLink = (t) => void followed.push(t)
  c.shell.alert = (m) => void alerts.push(m)
  c.shell.documentPath = () => "/home/alex/Books/Budget.xlsx"
  return { c, followed, alerts }
}

const type = (c: SheetController, text: string) => {
  c.beginEdit("enter", text)
  c.commitEdit()
}

describe("Insert › Link", () => {
  it("links the selection, puts the text in an empty cell and gives it the Hyperlink style", () => {
    const { c } = setup()
    expect(c.hyperlinks.displayText()).toBe("")
    c.hyperlinks.set({ address: "https://example.com/", tooltip: "Our site", text: "Example" })
    expect(c.sheet.hyperlinks).toEqual([{ range: at(0, 0), address: "https://example.com/", tooltip: "Our site" }])
    expect(c.sheet.cells.value(0, 0)).toBe("Example")
    expect(c.wb.styles.get(c.sheet.styleAt(0, 0))).toMatchObject({ cellStyle: "Hyperlink", underline: "single" })
    // One undo step takes it all back.
    c.undo()
    expect(c.sheet.hyperlinks).toEqual([])
    expect(c.sheet.cells.get(0, 0)).toBeUndefined()
  })

  it("leaves a number or formula as it is (Text to display is <<Selection in Document>>)", () => {
    const { c } = setup()
    type(c, "=1+2")
    c.goTo(at(0, 0))
    expect(c.hyperlinks.displayText()).toBeNull()
    c.hyperlinks.set({ location: "Sheet1!C5", text: null })
    expect(c.sheet.cells.input(0, 0)).toBe("=1+2")
    expect(c.hyperlinks.activeLink()?.location).toBe("Sheet1!C5")
  })

  it("Clear Hyperlinks keeps the look; Remove Hyperlinks takes the formatting too", () => {
    const { c } = setup()
    c.hyperlinks.set({ address: "https://example.com/", text: "Site" })
    c.hyperlinks.remove(false)
    expect(c.sheet.hyperlinks).toEqual([])
    expect(c.wb.styles.get(c.sheet.styleAt(0, 0)).underline).toBe("single")
    c.undo()
    c.hyperlinks.remove(true)
    expect(c.sheet.hyperlinks).toEqual([])
    expect(c.sheet.styleAt(0, 0)).toBe(0)
    expect(c.sheet.cells.value(0, 0)).toBe("Site")
  })

  it("Clear All removes links with everything else", () => {
    const { c } = setup()
    c.hyperlinks.set({ address: "https://example.com/", text: "Site" })
    c.format.clearAll()
    expect(c.sheet.hyperlinks).toEqual([])
  })

  it("Ctrl+K opens the dialog", () => {
    const { c } = setup()
    const open = vi.fn()
    c.shell.openDialog = open
    c.keyReady({ key: "k", primary: true, shift: false, alt: false })
    expect(open).toHaveBeenCalledWith("insertLink")
  })

  it("a protected sheet takes links only when Insert hyperlinks is allowed", () => {
    const { c, alerts } = setup()
    c.sheet.protection = { ...DEFAULT_PROTECTION }
    expect(c.hyperlinks.set({ address: "https://example.com/", text: null })).toBe(false)
    expect(alerts.length).toBe(1)
    c.sheet.protection = { ...DEFAULT_PROTECTION, insertHyperlinks: true, formatCells: true }
    c.sheet.cells.set(0, 0, { value: 1 })
    expect(c.hyperlinks.set({ address: "https://example.com/", text: null })).toBe(true)
  })
})

describe("typing an address", () => {
  it("makes web and e-mail addresses links, as Excel's AutoCorrect does", () => {
    const { c } = setup()
    type(c, "www.example.com")
    expect(c.sheet.hyperlinks[0]).toEqual({ range: at(0, 0), address: "http://www.example.com" })
    c.goTo(at(1, 0))
    type(c, "alex@example.com")
    expect(c.sheet.hyperlinks[1].address).toBe("mailto:alex@example.com")
    c.goTo(at(2, 0))
    type(c, "see www.example.com")
    c.goTo(at(3, 0))
    type(c, "12.5")
    expect(c.sheet.hyperlinks).toHaveLength(2)
  })
})

describe("following links", () => {
  it("goes to a place in the workbook, or says the reference isn't valid", () => {
    const { c, alerts } = setup()
    c.sheets.insert()
    const second = c.wb.sheets[1].name
    c.activate(0)
    c.hyperlinks.set({ location: `${second}!C4`, text: "There" })
    expect(c.hyperlinks.follow(0, 0)).toBe(true)
    expect(c.sheetIndex).toBe(1)
    expect(c.selection.active).toEqual({ row: 3, col: 2 })
    c.activate(0)
    c.goTo(at(1, 0))
    c.hyperlinks.set({ location: "Nowhere!A1", text: "Bad" })
    c.hyperlinks.follow(1, 0)
    expect(alerts).toEqual([BAD_REFERENCE])
  })

  it("opens workbooks in Cells and hands the rest to the system, with paths from the workbook's folder", () => {
    const { c, followed } = setup()
    c.hyperlinks.set({ address: "../Prices.xlsx", location: "Sheet1!B2", text: "Prices" })
    c.goTo(at(1, 0))
    c.hyperlinks.set({ address: "Reports/Q1.pdf", text: "Q1" })
    c.goTo(at(2, 0))
    c.hyperlinks.set({ address: "mailto:alex@example.com", text: "Mail" })
    c.goTo(at(3, 0))
    c.sheet.cells.set(3, 0, { input: '=HYPERLINK("https://example.com/help","Help")', value: "Help" })
    for (let row = 0; row < 4; row++) c.hyperlinks.follow(row, 0)
    expect(followed).toEqual([
      { kind: "workbook", path: "/home/alex/Prices.xlsx", location: "Sheet1!B2" },
      { kind: "external", target: "/home/alex/Books/Reports/Q1.pdf", file: true },
      { kind: "external", target: "mailto:alex@example.com", file: false },
      { kind: "external", target: "https://example.com/help", file: false },
    ])
    expect(c.hyperlinks.follow(5, 5)).toBe(false)
  })

  it("resolves relative paths the way Excel stores them", () => {
    expect(resolvePath("C:\\Users\\alex\\Budget.xlsx", "..\\Reports\\Q1.pdf")).toBe("C:\\Users\\Reports\\Q1.pdf")
    expect(resolvePath("/home/alex/Budget.xlsx", "/srv/x.pdf")).toBe("/srv/x.pdf")
    expect(resolvePath(null, "Q1.pdf")).toBe("Q1.pdf")
  })
})

describe("Insert › Symbol", () => {
  it("puts the character at the caret, or opens the cell with it at the end", () => {
    const { c } = setup()
    type(c, "Price")
    c.goTo(at(0, 0))
    c.insertText("€")
    expect(c.edit?.text).toBe("Price€")
    c.setEditText("Price€", 0)
    c.insertText("©")
    expect(c.edit?.text).toBe("©Price€")
    c.commitEdit()
    expect(c.sheet.cells.value(0, 0)).toBe("©Price€")
  })
})
