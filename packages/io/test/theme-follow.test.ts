import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { OFFICE_THEME, themeColorHex } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = () => new Uint8Array(readFileSync(join(__dirname, "fixtures/excel/theme-follow.xlsx")))

/** Excel's Facet colours (Page Layout › Themes), as a theme to change to. */
const FACET = {
  ...OFFICE_THEME,
  name: "Facet",
  colors: { ...OFFICE_THEME.colors, accent1: "90C226", accent2: "54A021", dk2: "2C3C43", lt2: "EBEBEB" },
}

describe("conditional formats and charts that follow the theme", () => {
  it("reads a rule's theme colours and a chart's theme colours as links, as Excel wrote them", async () => {
    const { workbook: wb } = await openXlsx(fixture(), "theme-follow.xlsx")
    const rule = wb.sheets[0].conditional[0].rule as { style: Record<string, unknown> }
    expect(rule.style).toMatchObject({ fillThemeColor: { theme: 5 }, fontThemeColor: { theme: 1 } })
    expect(rule.style.fill).toBe(`#${OFFICE_THEME.colors.accent2}`)
    const spec = wb.sheets[0].charts[0].spec as { format: Record<string, { fill?: string }> }
    expect(spec.format["series:0"].fill).toBe("scheme:accent1/shade=76000")
    expect(spec.format["series:1"].fill).toBe("scheme:accent1/tint=77000")
  })

  it("a new theme recolours the rule; the links are written back", async () => {
    const read = await openXlsx(fixture(), "theme-follow.xlsx")
    const wb = read.workbook
    wb.transact([{ kind: "theme", theme: FACET }], { label: "Theme" })
    const rule = wb.sheets[0].conditional[0].rule as { style: { fill?: string; fontColor?: string } }
    expect(rule.style.fill).toBe(themeColorHex({ theme: 5 }, FACET))
    expect(rule.style.fill).toBe("#54A021")
    const entries = await readZip(await saveXlsx(wb, { source: read.source }))
    const text = (name: string) => new TextDecoder().decode(entries.find((e) => e.name === name)!.data)
    expect(text("xl/styles.xml")).toMatch(
      /<dxf><font><color theme="1"\/><\/font><fill><patternFill><bgColor theme="5"\/>/,
    )
    expect(text("xl/charts/chart1.xml")).toContain('<a:schemeClr val="accent1"><a:shade val="76000"/></a:schemeClr>')
  })
})
