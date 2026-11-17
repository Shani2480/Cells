import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { copyTheme, OFFICE_THEME, type Theme } from "@cells/core"
import { readZip } from "@cells/ooxml"
import { openXlsx, saveXlsx } from "../src"

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures/excel", `${name}.xlsx`)))

/** Office's Facet theme (Document Themes 16\Facet.thmx). */
const FACET: Theme = {
  name: "Facet",
  colors: {
    dk1: "000000",
    lt1: "FFFFFF",
    dk2: "2C3C43",
    lt2: "EBEBEB",
    accent1: "90C226",
    accent2: "54A021",
    accent3: "E6B91E",
    accent4: "E76618",
    accent5: "C42F1A",
    accent6: "918655",
    hlink: "99CA3C",
    folHlink: "B9D181",
  },
  fonts: { major: "Trebuchet MS", minor: "Trebuchet MS" },
}

/** Excel's Range colours (BGR numbers) as "#RRGGBB". */
const bgr = (n: number) =>
  "#" + [n & 255, (n >> 8) & 255, (n >> 16) & 255].map((v) => v.toString(16).padStart(2, "0").toUpperCase()).join("")

/**
 * What Excel 365 showed for A1:A8 of theme.xlsx after Workbook.ApplyTheme(Facet.thmx): row, Interior.Color,
 * Font.Color, Font.Name, bottom border colour (0: none).
 */
const EXCEL_FACET = [
  [1, 2540176, 0, "Trebuchet MS", 0],
  [2, 7594943, 0, "Trebuchet MS", 0],
  [3, 16777215, 4406316, "Trebuchet MS", 0],
  [4, 16777215, 4406316, "Trebuchet MS", 0],
  [5, 16777215, 4406316, "Trebuchet MS", 2540176],
  [6, 16777215, 4220268, "Trebuchet MS", 0],
  [7, 255, 0, "Trebuchet MS", 0],
  [8, 16777215, 0, "Trebuchet MS", 0],
] as const

describe("Page Layout › Themes", () => {
  it("keeps Excel's theme colours and fonts, and a new theme restyles them as Excel does", async () => {
    const read = await openXlsx(fixture("theme"), "theme.xlsx")
    const wb = read.workbook
    expect(wb.theme.name).toBe("Office Theme")
    const style = (row: number) => wb.styles.get(wb.sheets[0].styleAt(row - 1, 0))
    expect(style(1)).toMatchObject({ fill: "#156082", fillThemeColor: { theme: 4 } })
    expect(style(4)).toMatchObject({ fontName: "Aptos Display", fontScheme: "major", fontThemeColor: { theme: 3 } })

    wb.transact([{ kind: "theme", theme: FACET }], { label: "Theme" })
    for (const [row, fill, font, name, border] of EXCEL_FACET) {
      const s = style(row)
      expect(s.fill ?? "#FFFFFF", `A${row} fill`).toBe(bgr(fill))
      expect(s.fontColor ?? "#000000", `A${row} font colour`).toBe(bgr(font))
      expect(s.fontName ?? wb.theme.fonts.minor, `A${row} font`).toBe(name)
      if (border) expect(s.borderBottom?.color, `A${row} border`).toBe(bgr(border))
    }

    // Saved: the theme part has Facet's colours and fonts, the styles keep their theme references.
    const bytes = await saveXlsx(wb, { source: read.source })
    const entries = await readZip(bytes)
    const text = (name: string) => new TextDecoder().decode(entries.find((e) => e.name === name)!.data)
    const theme = text("xl/theme/theme1.xml")
    expect(theme).toMatch(/<a:clrScheme name="Facet"><a:dk1><a:sysClr val="windowText" lastClr="000000"\/>/)
    expect(theme).toMatch(/<a:majorFont><a:latin typeface="Trebuchet MS"/)
    expect(theme).toMatch(/<a:fmtScheme name="Office">/) // the file's own effects stay
    const styles = text("xl/styles.xml")
    expect(styles).toMatch(/<fgColor theme="4" tint="0.39997558519241921"\/>/)
    expect(styles).toMatch(/<name val="Trebuchet MS"\/><family val="2"\/><scheme val="major"\/>/)
    const again = await openXlsx(bytes, "theme.xlsx")
    expect(again.workbook.theme).toMatchObject({ name: "Facet", fonts: FACET.fonts })
    expect(again.workbook.styles.get(again.workbook.sheets[0].styleAt(0, 0)).fill).toBe("#90C226")
  })

  it("writes a theme part for a new workbook only when its theme isn't Office's", async () => {
    const names = async (theme?: Theme) => {
      const read = await openXlsx(await saveXlsx(new (await import("@cells/core")).Workbook()), "x.xlsx")
      const wb = read.workbook
      if (theme) wb.transact([{ kind: "theme", theme }], { label: "Theme" })
      const entries = await readZip(await saveXlsx(wb, { source: read.source }))
      return entries.map((e) => e.name)
    }
    expect(await names()).not.toContain("xl/theme/theme1.xml")
    expect(await names({ ...copyTheme(OFFICE_THEME), ...FACET })).toContain("xl/theme/theme1.xml")
  })
})
