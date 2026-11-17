import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { defaultTextBox, parseLinear, Workbook, type MathList, type SheetShape } from "@cells/core"
import { readZip, writeZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"
import { parseTree } from "../src/xlsx/dom"
import { readMath } from "../src/xlsx/omml"

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)))
const drawing = async (bytes: Uint8Array) =>
  new TextDecoder().decode((await readZip(bytes)).find((e) => e.normalizedName.endsWith("drawings/drawing1.xml"))!.data)

/** Plain runs joined and minus signs as hyphens: Word and the parser split text into runs differently. */
function normal(list: MathList): unknown[] {
  const out: Record<string, unknown>[] = []
  for (const n of list) {
    const last = out[out.length - 1]
    if (n.t === "r" && !n.sty) {
      const text = n.text.replace(/−/g, "-")
      if (last?.t === "r" && !last.sty) last.text += text
      else out.push({ t: "r", text })
      continue
    }
    const copy: Record<string, unknown> = { ...n }
    for (const [k, v] of Object.entries(copy)) {
      if (!Array.isArray(v)) continue
      const lists = v as unknown[]
      copy[k] = Array.isArray(lists[0])
        ? lists.map((x) => (Array.isArray((x as unknown[])[0]) ? (x as MathList[]).map(normal) : normal(x as MathList)))
        : normal(v as MathList)
    }
    out.push(copy)
  }
  return out
}

describe("Insert › Equation in xlsx", () => {
  it("reads Office Math as Word writes Office's built-in equations, the same as their linear form", () => {
    const rows = readFileSync(join(__dirname, "fixtures/office/equations.tsv"), "utf8").trim().split("\n")
    expect(rows).toHaveLength(9)
    for (const row of rows) {
      const [name, linear, xml] = row.split("\t")
      const doc = parseTree(
        xml!.replace(
          "<m:oMathPara>",
          `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="w">`,
        ),
      )
      const [read] = readMath(doc, doc.root)
      expect(normal(read!), name).toEqual(normal(parseLinear(linear!)))
    }
  })

  it("opens Excel's equation as an equation and saves it back exactly while it is unchanged", async () => {
    const opened = await openXlsx(fixture("excel/equation.xlsx"), "equation.xlsx")
    expect(opened.warnings).toEqual([])
    const shape = opened.workbook.sheets[0]!.objects[0] as SheetShape
    expect(shape.textBox).toBe(true)
    const [para] = shape.text!.paragraphs
    expect(para!.math).toHaveLength(1)
    expect(normal(para!.math![0]!)).toEqual(normal(parseLinear("x=(−b±√(b^2−4ac))/2a")))
    const before = await drawing(fixture("excel/equation.xlsx"))
    const again = await drawing(await saveXlsx(opened.workbook, { source: opened.source }))
    const choice = (xml: string) => /<mc:AlternateContent[\s\S]*<\/mc:AlternateContent>/.exec(xml)![0]
    expect(choice(again)).toBe(choice(before))
  })

  it("writes a new equation in Excel 2010's form, with its linear text for older readers", async () => {
    const wb = new Workbook()
    const box = defaultTextBox()
    box.text = {
      paragraphs: [{ runs: [{ text: "", size: 14 }], math: [parseLinear("e^(iπ)+1=0")] }],
      autoFit: true,
      wrap: false,
    }
    wb.sheets[0]!.objects.push({
      kind: "shape",
      id: 3,
      name: "TextBox 3",
      geometry: "rect",
      textBox: true,
      anchor: { from: { col: 1, row: 1, x: 0, y: 0 }, to: { col: 3, row: 3, x: 0, y: 0 } },
      ...box,
    })
    const bytes = await saveXlsx(wb)
    const xml = await drawing(bytes)
    expect(xml).toMatch(/<mc:Choice xmlns:a14="[^"]+" Requires="a14"><xdr:sp\b/)
    expect(xml).toContain(`<a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">`)
    // Letters in math italic, the size kept in each run, minus signs as Office writes them.
    expect(xml).toContain(`<m:t>𝑒</m:t>`)
    expect(xml).toMatch(/<m:sup><m:r><a:rPr lang="en-US" sz="1400" i="1">/)
    expect(/<mc:Fallback>[\s\S]*?<a:t>([^<]*)<\/a:t>/.exec(xml)?.[1]).toBe("𝑒^(𝑖𝜋)+1=0")
    const ids = [...xml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map((m) => m[1])
    expect(ids).toHaveLength(2)
    expect(ids[0]).toBe(ids[1])

    const back = await openXlsx(bytes, "new.xlsx")
    const para = (back.workbook.sheets[0]!.objects[0] as SheetShape).text!.paragraphs[0]!
    expect(para.runs[0]?.size).toBe(14)
    expect(normal(para.math![0]!)).toEqual(normal(parseLinear("e^(iπ)+1=0")))
  })

  it("keeps text written beside an equation in the same paragraph, as the equation's normal text", async () => {
    const bytes = fixture("excel/equation.xlsx")
    const entries = await readZip(bytes)
    const part = entries.find((e) => e.normalizedName.endsWith("drawings/drawing1.xml"))!
    const xml = new TextDecoder()
      .decode(part.data)
      .replace("<a:pPr/><a14:m>", `<a:pPr/><a:r><a:rPr lang="en-US"/><a:t>Roots: </a:t></a:r><a14:m>`)
    const patched = await writeZip(
      entries.map((e) => ({ name: e.name, data: e === part ? new TextEncoder().encode(xml) : e.data })),
    )
    const read = await openXlsx(patched, "mixed.xlsx")
    const line = (read.workbook.sheets[0]!.objects[0] as SheetShape).text!.paragraphs[0]!.math![0]!
    expect(line[0]).toEqual({ t: "r", text: "Roots: ", sty: "p" })
    expect(line.length).toBeGreaterThan(3)
  })
})
