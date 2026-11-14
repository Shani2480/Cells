// Flash Fill: the patterns Excel's own Flash Fill examples show (names, initials, emails, phone numbers, case).
import { describe, expect, it } from "vitest"
import { flashFill, flashTokens } from "../src"

const fill = (rows: string[][], outputs: (string | null)[]) => flashFill(rows, outputs)

describe("Flash Fill", () => {
  it("splits words, numbers and punctuation", () => {
    expect(flashTokens("McDonald, Ann-Marie 42")).toEqual(["Mc", "Donald", ",", " ", "Ann", "-", "Marie", " ", "42"])
  })

  it("first and last names from one example; the last name counts from the end", () => {
    const rows = [["Nancy Davolio"], ["Andrew Fuller"], ["Janet Mary Leverling"]]
    expect(fill(rows, ["Nancy", null, null])).toEqual([null, "Andrew", "Janet"])
    expect(fill(rows, ["Davolio", null, null])).toEqual([null, "Fuller", "Leverling"])
  })

  it("rearranges with typed text, initials and case", () => {
    const rows = [["Nancy Davolio"], ["andrew fuller"]]
    expect(fill(rows, ["Davolio, Nancy", null])).toEqual([null, "fuller, andrew"])
    expect(fill(rows, ["N. Davolio", null])).toEqual([null, "a. fuller"])
    expect(fill(rows, ["NANCY", null])).toEqual([null, "ANDREW"])
    expect(fill([["nancy davolio"], ["ANDREW FULLER"]], ["Nancy Davolio", null])).toEqual([null, "Andrew Fuller"])
  })

  it("joins two columns, and builds email addresses", () => {
    const rows = [
      ["Nancy", "Davolio"],
      ["Andrew", "Fuller"],
    ]
    expect(fill(rows, ["nancy.davolio@example.com", null])).toEqual([null, "andrew.fuller@example.com"])
    expect(fill(rows, ["Nancy Davolio", null])).toEqual([null, "Andrew Fuller"])
  })

  it("pulls digits out of phone numbers; a second example settles which part", () => {
    const rows = [["(206) 555-0100"], ["(425) 555-0199"], ["(360) 555-0123"]]
    expect(fill(rows, ["2065550100", null, null])).toEqual([null, "4255550199", "3605550123"])
    expect(fill(rows, ["0100", null, null])).toEqual([null, "0199", "0123"])
  })

  it("uses every example, and gives up when none fits", () => {
    const rows = [["a-b"], ["c d"], ["e-f"]]
    expect(fill([["Seattle, WA"], ["Austin, TX"], ["Boise, ID"]], ["WA", "TX", null])).toEqual([null, null, "ID"])
    expect(fill(rows, ["x", null, null])).toEqual([null, "x", "x"])
    expect(fill(rows, ["zz", "q", null])).toBeNull()
    expect(fill([["Ann"], [""]], ["Ann", null])).toEqual([null, null])
  })
})
