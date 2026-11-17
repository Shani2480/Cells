import { describe, expect, it } from "vitest"
import { applyTint, copyTheme, OFFICE_THEME, themeColorHex, Workbook, type Theme } from "../src"

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

  it("tints like Excel 365 (Interior.TintAndShade on the Office theme's colours)", () => {
    const samples: [string, number, string][] = [
      ["E8E8E8", -0.5, "747474"],
      ["E8E8E8", -0.25, "ADADAD"],
      ["E8E8E8", 0.2, "ECECEC"],
      ["E8E8E8", 0.4, "F0F0F0"],
      ["E8E8E8", 0.6, "F5F5F5"],
      ["E8E8E8", 0.8, "FAFAFA"],
      ["E8E8E8", -0.1, "D0D0D0"],
      ["E8E8E8", 0.9, "FCFCFC"],
      ["0E2841", -0.5, "061320"],
      ["0E2841", -0.25, "0B1D2F"],
      ["0E2841", 0.2, "1D5287"],
      ["0E2841", 0.4, "2D7DCE"],
      ["0E2841", 0.6, "70A8E0"],
      ["0E2841", 0.8, "B8D3EF"],
      ["0E2841", -0.1, "0D2339"],
      ["0E2841", 0.9, "DAE9F8"],
      ["156082", -0.5, "0B3040"],
      ["156082", -0.25, "104861"],
      ["156082", 0.2, "1F8EBE"],
      ["156082", 0.4, "44B3E1"],
      ["156082", 0.6, "83CCEB"],
      ["156082", 0.8, "C0E6F5"],
      ["156082", -0.1, "135673"],
      ["156082", 0.9, "E0F2FA"],
      ["E97132", -0.5, "7E350E"],
      ["E97132", -0.25, "BE5014"],
      ["E97132", 0.2, "ED8D5A"],
      ["E97132", 0.4, "F1A983"],
      ["E97132", 0.6, "F7C7AC"],
      ["E97132", 0.8, "FBE2D5"],
      ["E97132", -0.1, "E46018"],
      ["E97132", 0.9, "FDF0EA"],
      ["196B24", -0.5, "0D3512"],
      ["196B24", -0.25, "12501A"],
      ["196B24", 0.2, "27A737"],
      ["196B24", 0.4, "47D359"],
      ["196B24", 0.6, "83E28E"],
      ["196B24", 0.8, "C1F0C8"],
      ["196B24", -0.1, "165F20"],
      ["196B24", 0.9, "E0F8E3"],
      ["0F9ED5", -0.5, "074F69"],
      ["0F9ED5", -0.25, "0C769E"],
      ["0F9ED5", 0.2, "2BBAF0"],
      ["0F9ED5", 0.4, "61CBF3"],
      ["0F9ED5", 0.6, "94DCF8"],
      ["0F9ED5", 0.8, "CAEDFB"],
      ["0F9ED5", -0.1, "0E8DBE"],
      ["0F9ED5", 0.9, "E3F5FD"],
      ["A02B93", -0.5, "51154A"],
      ["A02B93", -0.25, "782170"],
      ["A02B93", 0.2, "CB3DBD"],
      ["A02B93", 0.4, "D86DCD"],
      ["A02B93", 0.6, "E49EDD"],
      ["A02B93", 0.8, "F2CEEF"],
      ["A02B93", -0.1, "8F2785"],
      ["A02B93", 0.9, "F9E6F7"],
      ["4EA72E", -0.5, "275317"],
      ["4EA72E", -0.25, "3C7D22"],
      ["4EA72E", 0.2, "6ACB45"],
      ["4EA72E", 0.4, "8ED973"],
      ["4EA72E", 0.6, "B5E6A2"],
      ["4EA72E", 0.8, "DAF2D0"],
      ["4EA72E", -0.1, "479629"],
      ["4EA72E", 0.9, "ECF9E8"],
    ]
    for (const [hex, t, want] of samples) expect(applyTint(hex, t), `${hex} ${t}`).toBe(want)
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
