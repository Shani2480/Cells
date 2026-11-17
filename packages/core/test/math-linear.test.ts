import { describe, expect, it } from "vitest"
import { BUILTIN_EQUATIONS, MATH_STRUCTURES, mathText, parseLinear, toLinear, type MathList } from "../src"

const tree = (text: string) => JSON.parse(JSON.stringify(parseLinear(text))) as MathList

describe("Equations in Office's linear format", () => {
  it("builds up fractions, scripts and roots as Office does", () => {
    expect(tree("x=(−b±√(b^2−4ac))/2a")).toEqual([
      { t: "r", text: "x" },
      { t: "r", text: "=" },
      {
        t: "f",
        num: [
          { t: "r", text: "−" },
          { t: "r", text: "b" },
          { t: "r", text: "±" },
          {
            t: "rad",
            deg: [],
            e: [
              { t: "sSup", e: [{ t: "r", text: "b" }], sup: [{ t: "r", text: "2" }] },
              { t: "r", text: "−" },
              { t: "r", text: "4ac" },
            ],
          },
        ],
        den: [{ t: "r", text: "2a" }],
      },
    ])
    // A space ends a fraction: ½ times (α+β), not 1 over 2(α+β).
    const half = tree("1/2 (α+β)")
    expect(half.map((n) => n.t)).toEqual(["f", "d"])
    expect(tree("√(3&x)")).toEqual([{ t: "rad", deg: [{ t: "r", text: "3" }], e: [{ t: "r", text: "x" }] }])
  })

  it("reads n-ary operators, functions, matrices and one-sided brackets", () => {
    const [sum] = tree("∑_(k=0)^n▒〖x^k〗")
    expect(sum).toMatchObject({
      t: "nary",
      chr: "∑",
      limLoc: "undOvr",
      sub: [{ text: "k" }, { text: "=" }, { text: "0" }],
    })
    expect(tree("∫_0^1▒x")[0]).toMatchObject({ t: "nary", limLoc: "subSup" })
    expect(tree("sin⁡〖2x〗")).toEqual([
      { t: "func", name: [{ t: "r", text: "sin", sty: "p" }], e: [{ t: "r", text: "2x" }] },
    ])
    expect(tree("■(a&b@c&d)")[0]).toMatchObject({
      t: "m",
      rows: [
        [[{ text: "a" }], [{ text: "b" }]],
        [[{ text: "c" }], [{ text: "d" }]],
      ],
    })
    expect(tree("{■(1@2)┤")[0]).toMatchObject({ t: "d", beg: "{", end: "" })
    expect(tree("‖v‖")).toEqual([{ t: "d", beg: "‖", end: "‖", items: [[{ t: "r", text: "v" }]] }])
    expect(tree("\\alpha+\\infty")).toEqual([
      { t: "r", text: "α" },
      { t: "r", text: "+" },
      { t: "r", text: "∞" },
    ])
  })

  it("prints the tree back so that it builds up the same, for Office's equations and every structure", () => {
    const all = [
      ...BUILTIN_EQUATIONS.map((e) => e.linear),
      ...MATH_STRUCTURES.flatMap((s) => s.items.map((i) => i.linear)),
    ]
    for (const text of all) {
      const once = tree(text)
      expect(tree(toLinear(once)), text).toEqual(once)
    }
    // Letters print in math italic, as Office shows them; upright names stay upright.
    expect(toLinear(tree("x=(−b±√(b^2−4ac))/2a"))).toBe("𝑥=(−𝑏±√(𝑏^2−4𝑎𝑐))/2𝑎")
    expect(toLinear(tree("2 sin⁡〖1/2 (α+β)〗"))).toBe("2 sin⁡〖1/2 (𝛼+𝛽)〗")
    expect(mathText(tree("a/b"))).toBe("a/b")
  })
})
