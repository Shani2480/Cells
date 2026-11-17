import { describe, expect, it } from "vitest"
import {
  completeAddress,
  formulaLink,
  hyperlinkAt,
  hyperlinkKind,
  linkFromText,
  linkScreenTip,
  linkText,
  mailAddress,
  mailParts,
  withHyperlink,
  Workbook,
  type Hyperlink,
} from "../src"

const r = (r0: number, c0: number, r1 = r0, c1 = c0) => ({ r0, c0, r1, c1 })

describe("hyperlinks", () => {
  it("a new link replaces the cells' old ones and splits a larger link around them", () => {
    const links: Hyperlink[] = [{ range: r(0, 0, 2, 0), address: "https://example.com/" }]
    const next = withHyperlink(links, [r(1, 0)], { location: "#Data!A1", address: " " })
    expect(next).toEqual([
      { range: r(0, 0), address: "https://example.com/" },
      { range: r(2, 0), address: "https://example.com/" },
      { range: r(1, 0), location: "Data!A1" },
    ])
    expect(hyperlinkAt(next, 1, 0)?.location).toBe("Data!A1")
    expect(withHyperlink(next, [r(0, 0, 5, 5)], null)).toEqual([])
  })

  it("follows inserted and deleted rows", () => {
    const wb = new Workbook(["S"])
    wb.sheets[0].hyperlinks = [{ range: r(3, 1), address: "https://example.com/" }]
    wb.transact([{ kind: "insertRows", sheet: 0, at: 1, count: 2 }], { label: "Insert" })
    expect(wb.sheets[0].hyperlinks[0].range).toEqual(r(5, 1))
    wb.transact([{ kind: "deleteRows", sheet: 0, at: 5, count: 1 }], { label: "Delete" })
    expect(wb.sheets[0].hyperlinks).toEqual([])
    wb.undo()
    expect(wb.sheets[0].hyperlinks[0].range).toEqual(r(5, 1))
  })

  it("completes typed addresses and reads mail links as Excel's dialog does", () => {
    expect(completeAddress("www.example.com")).toBe("http://www.example.com")
    expect(completeAddress("alex@example.com")).toBe("mailto:alex@example.com")
    expect(completeAddress("C:\\Reports\\Q1.pdf")).toBe("C:\\Reports\\Q1.pdf")
    expect(mailAddress("alex@example.com", "Hello there")).toBe("mailto:alex@example.com?subject=Hello%20there")
    expect(mailParts("mailto:alex@example.com?subject=Hello%20there")).toEqual({
      to: "alex@example.com",
      subject: "Hello there",
    })
    expect(hyperlinkKind({ address: "mailto:a@b.c" })).toBe("mail")
    expect(hyperlinkKind({ address: "https://a.b" })).toBe("web")
    expect(hyperlinkKind({ address: "C:\\x.pdf" })).toBe("file")
    expect(hyperlinkKind({ location: "Sheet1!A1" })).toBe("place")
    expect(linkText({ address: "mailto:a@b.c" })).toBe("a@b.c")
    expect(linkText({ address: "Book.xlsx", location: "Sheet1!A1" })).toBe("Book.xlsx - Sheet1!A1")
    expect(linkScreenTip({ location: "Data!B2" })).toBe(
      "#Data!B2 - Click once to follow. Click and hold to select this cell.",
    )
    expect(linkScreenTip({ address: "https://a.b", tooltip: "Ours" })).toBe("Ours")
  })

  it("reads HYPERLINK formulas' targets", () => {
    expect(linkFromText("#Data!B2")).toEqual({ location: "Data!B2" })
    expect(linkFromText("[Prices.xlsx]Sheet1!A1")).toEqual({ address: "Prices.xlsx", location: "Sheet1!A1" })
    expect(linkFromText("C:\\Books\\Prices.xlsx#Sheet1!A1")).toEqual({
      address: "C:\\Books\\Prices.xlsx",
      location: "Sheet1!A1",
    })
    expect(linkFromText("https://example.com/a#part")).toEqual({ address: "https://example.com/a#part" })
    const wb = new Workbook(["S"])
    const s = wb.sheets[0]
    s.cells.set(0, 0, { value: "https://example.com/" })
    s.cells.set(0, 1, { input: '=HYPERLINK(A1&"help","Help")', value: "Help" })
    s.cells.set(0, 2, { input: "=A1", value: "https://example.com/" })
    expect(formulaLink(wb.sheets, 0, 0, 1, 0)).toEqual({ address: "https://example.com/help" })
    expect(formulaLink(wb.sheets, 0, 0, 2, 0)).toBeNull()
  })
})
