import { describe, expect, it } from "vitest"
import { decodeTSV, encodeHTML, encodeTSV } from "../src/clipboard"

describe("clipboard text", () => {
  it("writes Excel's TSV", () => {
    expect(
      encodeTSV([
        ["a", "b"],
        ["1", ""],
      ]),
    ).toBe("a\tb\r\n1\t\r\n")
    expect(encodeTSV([['say "hi"', "two\nlines", "tab\there"]])).toBe('"say ""hi"""\t"two\nlines"\t"tab\there"\r\n')
  })

  it("reads TSV with quotes, CRLF and a final line break", () => {
    expect(decodeTSV("a\tb\r\n1\t\r\n")).toEqual([
      ["a", "b"],
      ["1", ""],
    ])
    expect(decodeTSV('"say ""hi"""\t"two\nlines"\r\n')).toEqual([['say "hi"', "two\nlines"]])
    expect(decodeTSV("x")).toEqual([["x"]])
    expect(decodeTSV("a\n\nb")).toEqual([["a"], [""], ["b"]])
    // A quote that does not close a field is literal text.
    expect(decodeTSV('5" pipe\tx')).toEqual([['5" pipe', "x"]])
    expect(decodeTSV('"open\tx')).toEqual([['"open', "x"]])
  })

  it("round-trips any table", () => {
    const rows = [
      ["", "a\tb", '"q"'],
      ["line\r\nbreak", "plain", ""],
    ]
    expect(decodeTSV(encodeTSV(rows))).toEqual(rows)
  })

  it("writes an HTML table", () => {
    expect(encodeHTML([["<a>", "x\ny"]])).toContain("<td>&lt;a&gt;</td><td>x<br>y</td>")
  })
})
