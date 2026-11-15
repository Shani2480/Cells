import { describe, expect, it } from "vitest"
import { a1Range, openLink, parseOpenRequest } from "./link"

describe("cells:// links", () => {
  it("reads the file and range", () => {
    expect(parseOpenRequest("cells://open?file=%2Fhome%2Falex%2FBudget%202026.xlsx&range=Sheet1!A1%3AD10")).toEqual({
      path: "/home/alex/Budget 2026.xlsx",
      range: "Sheet1!A1:D10",
    })
    expect(parseOpenRequest("cells://open?file=C%3A%5CUsers%5Calex%5CBudget.xlsx")).toEqual({
      path: "C:\\Users\\alex\\Budget.xlsx",
    })
    expect(parseOpenRequest("CELLS:open?file=%2Fsrv%2Fa.xlsx")).toEqual({ path: "/srv/a.xlsx" })
  })

  it("takes Word's R1C1 ranges, and sheet names in quotes", () => {
    expect(a1Range("Sheet1!R1C1:R10C4")).toBe("Sheet1!A1:D10")
    expect(a1Range("'My Sheet'!R2C28")).toBe("'My Sheet'!AB2")
    expect(a1Range("Totals")).toBe("Totals")
  })

  it("refuses other actions, relative paths and other schemes", () => {
    expect(parseOpenRequest("cells://delete?file=%2Fa.xlsx")).toBeNull()
    expect(parseOpenRequest("cells://open?file=Budget.xlsx")).toBeNull()
    expect(parseOpenRequest("cells://open")).toBeNull()
    expect(parseOpenRequest("https://example.com/a.xlsx")).toBeNull()
    expect(parseOpenRequest("/home/alex/Budget.xlsx")).toEqual({ path: "/home/alex/Budget.xlsx" })
  })

  it("writes links that read back", () => {
    const link = openLink("/home/alex/Q1 Budget.xlsx", "'Sales 2026'!B2:C9")
    expect(link).toBe("cells://open?file=%2Fhome%2Falex%2FQ1%20Budget.xlsx&range=%27Sales%202026%27%21B2%3AC9")
    expect(parseOpenRequest(link)).toEqual({ path: "/home/alex/Q1 Budget.xlsx", range: "'Sales 2026'!B2:C9" })
  })
})
