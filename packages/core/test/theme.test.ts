import { describe, expect, it } from "vitest"
import { copyTheme, OFFICE_THEME, themeColorHex, Workbook, type Theme } from "../src"

const BLUE: Theme = {
  ...copyTheme(OFFICE_THEME),
  name: "Custom",
  colors: { ...OFFICE_THEME.colors, accent1: "4472C4", dk2: "44546A" },
  fonts: { major: "Calibri Light", minor: "Calibri" },
}

describe("themes", () => {
  it("resolves theme colours with Excel's tint", () => {
    expect(themeColorHex({ theme: 4 }, OFFICE_THEME)).toBe("#156082")
    expect(themeColorHex({ theme: 4, tint: 0.7999816888943144 }, BLUE)).toBe("#D9E1F2") // Excel's "Blue, Accent 1, Lighter 80%"
    expect(themeColorHex({ theme: 1 }, BLUE)).toBe("#000000")
  })

  it("a theme change restyles theme colours and fonts, keeps the rest, and undoes", () => {
    const wb = new Workbook()
    const themed = wb.styles.intern({
      fill: "#156082",
      fillThemeColor: { theme: 4 },
      fontColor: "#0E2841",
      fontThemeColor: { theme: 3 },
      fontName: "Aptos Display",
      fontScheme: "major",
      borderTop: { style: "thin", color: "#156082", theme: { theme: 4 } },
    })
    const plain = wb.styles.intern({ fill: "#156082", fontName: "Arial" })
    wb.transact([{ kind: "theme", theme: BLUE }], { label: "Theme" })
    expect(wb.theme.name).toBe("Custom")
    expect(wb.styles.get(themed)).toMatchObject({
      fill: "#4472C4",
      fontColor: "#44546A",
      fontName: "Calibri Light",
      borderTop: { color: "#4472C4" },
    })
    expect(wb.styles.get(plain)).toEqual({ fill: "#156082", fontName: "Arial" })
    // Interning finds the restyled format under its new look.
    expect(wb.styles.intern({ ...wb.styles.get(themed) })).toBe(themed)
    wb.undo()
    expect(wb.theme.name).toBe("Office Theme")
    expect(wb.styles.get(themed)).toMatchObject({ fill: "#156082", fontName: "Aptos Display" })
  })

  it("a colour or font set on its own leaves the theme", () => {
    const wb = new Workbook()
    const id = wb.styles.intern({
      fill: "#156082",
      fillThemeColor: { theme: 4 },
      fontName: "Aptos Display",
      fontScheme: "major",
    })
    const red = wb.styles.derive(id, { fill: "#FF0000", fontName: "Arial" })
    expect(wb.styles.get(red)).toEqual({ fill: "#FF0000", fontName: "Arial" })
    const accent2 = wb.styles.derive(id, { fill: "#E97132", fillThemeColor: { theme: 5 } })
    expect(wb.styles.get(accent2).fillThemeColor).toEqual({ theme: 5 })
  })
})
