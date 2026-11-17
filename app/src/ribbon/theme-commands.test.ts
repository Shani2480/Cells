import { OFFICE_THEME, parseRange, selectRange, tableStyleDef } from "@cells/core"
import { describe, expect, it } from "vitest"
import { fontFor, setBodyFont } from "@/grid/cell-look"
import { SheetController } from "@/sheet/controller"
import { themeAccents } from "@/sheet/charts"
import { CELL_STYLES, cellStylePatch } from "./cell-styles"
import { themeGrid } from "./colors"
import { OFFICE_COLOR_SETS, OFFICE_FONT_SETS, OFFICE_THEMES } from "./theme-commands"

const facet = OFFICE_THEMES.find((t) => t.name === "Facet")!

describe("Page Layout › Themes in the app", () => {
  it("lists Office's themes, colour sets and font sets with the default first", () => {
    expect(OFFICE_THEMES[0].name).toBe("Office Theme")
    expect(OFFICE_THEMES).toHaveLength(11)
    expect(facet.colors.accent1).toBe("90C226")
    expect(OFFICE_COLOR_SETS.map((s) => s.label).slice(0, 4)).toEqual([
      "Office",
      "Office 2007 - 2010",
      "Office 2013 - 2022",
      "Grayscale",
    ])
    expect(OFFICE_FONT_SETS.find((s) => s.label === "Office 2013 - 2022")).toMatchObject({
      major: "Calibri Light",
      minor: "Calibri",
    })
  })

  it("the colour menu shows the theme's colours with Excel's tints, and hands back their references", () => {
    const office = themeGrid(OFFICE_THEME)
    expect(office[0][4]).toMatchObject({ name: "Dark Teal, Accent 1", hex: "#156082", ref: { theme: 4 } })
    expect(office[1][4]).toMatchObject({ hex: "#C0E6F5", ref: { theme: 4, tint: 0.7999816888943144 } })
    expect(themeGrid(facet)[0][4]).toMatchObject({ hex: "#90C226", ref: { theme: 4 } })
  })

  it("theme-coloured cells, cell styles, table styles, charts and the body font follow a new theme", () => {
    const c = new SheetController()
    c.setSelection(selectRange(parseRange("A1")!))
    c.format.fill("#156082", { theme: 4 })
    c.setSelection(selectRange(parseRange("A2")!))
    c.format.apply(cellStylePatch(CELL_STYLES.find((s) => s.name === "Heading 1")!), "Cell Style")
    c.transact([{ kind: "theme", theme: facet }], { label: "Theme", before: c.selection })

    const style = (row: number) => c.wb.styles.get(c.sheet.styleAt(row, 0))
    expect(style(0).fill).toBe("#90C226")
    expect(style(1)).toMatchObject({ fontColor: "#2C3C43", borderBottom: { color: "#90C226" } })
    // A style applied under the new theme takes its colours straight away.
    expect(
      cellStylePatch(
        CELL_STYLES.find((s) => s.name === "Accent1")!,
        facet,
      ),
    ).toMatchObject({
      fill: "#90C226",
      fillThemeColor: { theme: 4 },
    })
    expect(tableStyleDef("TableStyleMedium2", facet)?.headerRow?.fill).toBe("#90C226")
    expect(tableStyleDef("TableStyleMedium2")?.headerRow?.fill).toBe("#156082")
    expect(themeAccents(c.wb.theme)[0]).toBe("#90C226")

    setBodyFont(c.wb.theme.fonts.minor)
    expect(fontFor({}).font).toMatch(/^[\d.]+px "Trebuchet MS", /)
    setBodyFont("Aptos Narrow")

    c.wb.undo()
    expect(style(0).fill).toBe("#156082")
  })
})
