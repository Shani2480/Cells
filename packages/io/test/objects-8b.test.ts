import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { isChart, SIGNING_INSTRUCTIONS, Workbook, type SheetPicture } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures/excel", `${name}.xlsx`)))
const text = (entries: { name: string; data: Uint8Array }[], name: string) =>
  new TextDecoder().decode(entries.find((e) => e.name === name)?.data)
const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="

const lines = (wb: Workbook, sheet = 0) =>
  wb.sheets[sheet].objects.filter((o): o is SheetPicture => !isChart(o) && o.kind === "picture" && !!o.signature)

describe("signature lines (Insert › Signature Line)", () => {
  it("reads Excel's line with its signer, and writes it back to the VML drawing", async () => {
    const read = await openXlsx(fixture("sigline"), "sigline.xlsx")
    expect(read.warnings.join()).not.toMatch(/Form controls/)
    const [line] = lines(read.workbook)
    expect(line.signature).toMatchObject({
      id: "{997C6EAC-CB60-40DE-8C1E-82E1F5C0E7E7}",
      provider: "{000CD6A4-0000-0000-C000-000000000046}",
      signer: "Alex Example",
      title: "Manager",
      email: "alex@example.com",
      instructions: "Check the totals before signing.",
      allowComments: true,
      showDate: true,
    })
    expect(line.anchor).toEqual({ from: { col: 0, x: 0, row: 0, y: 0 }, to: { col: 2, x: 0, row: 8, y: 0 } })
    expect(line.src).toMatch(/^data:image\/x-emf;base64,/)

    const entries = await readZip(await saveXlsx(read.workbook, { source: read.source }))
    const vml = text(entries, "xl/drawings/vmlDrawing1.vml")
    expect(vml).toMatch(/<o:signatureline [^>]*o:suggestedsigner="Alex Example"[^>]*issignatureline="t"\/>/)
    expect(vml).toContain("<x:Anchor>0, 0, 0, 0, 2, 0, 8, 0</x:Anchor>")
    expect(text(entries, "xl/drawings/_rels/vmlDrawing1.vml.rels")).toMatch(/media\/image\d+\.emf"/)
    // Not in a DrawingML drawing as well.
    expect(entries.some((e) => e.name.startsWith("xl/drawings/drawing"))).toBe(false)
    const again = await openXlsx(await saveXlsx(read.workbook, { source: read.source }), "x.xlsx")
    expect(lines(again.workbook)[0].signature).toEqual(line.signature)
  })

  it("a new line goes in a VML drawing next to the sheet's notes", async () => {
    const wb = new Workbook()
    wb.sheets[0].annotations.set(1, 1, { kind: "note", author: "Alex", text: "Check", shown: false })
    wb.sheets[0].objects.push({
      kind: "picture",
      id: 1,
      name: "Picture 1",
      anchor: { from: { col: 3, x: 0, row: 2, y: 0 }, to: { col: 6, x: 0, row: 7, y: 0 } },
      src: PNG,
      signature: {
        id: "{11111111-2222-3333-4444-555555555555}",
        signer: "Sam Sample",
        title: "",
        email: "",
        instructions: SIGNING_INSTRUCTIONS,
        allowComments: false,
        showDate: false,
      },
    })
    const bytes = await saveXlsx(wb)
    const vml = text(await readZip(bytes), "xl/drawings/vmlDrawing1.vml")
    expect(vml.match(/<v:shape /g)).toHaveLength(2)
    expect(vml).toContain('showsigndate="f"')
    const again = await openXlsx(bytes, "x.xlsx")
    expect(again.workbook.sheets[0].annotations.all()).toHaveLength(1)
    expect(lines(again.workbook)[0].signature).toMatchObject({ signer: "Sam Sample", showDate: false })
  })
})
